"""Private Office RPC. Optional format/UNO libraries execute only in this guest."""
import base64
import json
import os
from pathlib import Path
import re
import socketserver
import subprocess
import sys
import tempfile
import time
import uuid

import uno
import unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.document.MacroExecMode import NEVER_EXECUTE
from com.sun.star.document import XDocumentEventListener
from display import OfficeDisplay
import artifact_editor
import documents

ROOT = Path('/var/lib/a0-office/files')
ROOT.mkdir(parents=True, exist_ok=True)
MAX_MESSAGE = 700 * 1024 * 1024
FORMATS = {'odt','ods','odp','docx','xlsx','pptx'}
FILTERS = {'odt':'writer8','docx':'Office Open XML Text','ods':'calc8','xlsx':'Calc MS Excel 2007 XML',
           'odp':'impress8','pptx':'Impress MS PowerPoint 2007 XML'}
display = OfficeDisplay('shared')
components = {}
listeners = {}
desktop = None
processes = []


class DocumentEvents(unohelper.Base, XDocumentEventListener):
    def __init__(self, path):
        self.path, self.closed = path, False

    def documentEventOccured(self, event):
        if event.EventName in {'OnSaveDone', 'OnSaveAsDone'}:
            self.path = Path(uno.fileUrlToSystemPath(event.Source.getURL())).resolve()

    def disposing(self, event):
        self.closed = True


def open_component(identifier, path, readonly=False):
    component = load(path, readonly=readonly)
    listener = DocumentEvents(path)
    component.addDocumentEventListener(listener)
    components[identifier], listeners[identifier] = component, listener
    return component


def properties(**values):
    return tuple(PropertyValue(Name=key, Value=value) for key, value in values.items())


def ensure_office():
    global desktop
    if desktop is not None:
        return desktop
    display.ensure_display()
    env = {**display._env(), 'SAL_USE_VCLPLUGIN':'gtk3', 'XDG_CONFIG_HOME': display.gtk_config}
    log = open('/run/a0-office/office.log', 'ab')
    processes.append(subprocess.Popen(['matchbox-window-manager', '-use_titlebar', 'no'], env=env, stdout=log, stderr=log))
    processes.append(subprocess.Popen(['dbus-run-session', '--', 'libreoffice', '--invisible', '--norestore', '--nodefault',
        '--nofirststartwizard', '-env:UserInstallation=file:///var/lib/a0-office/profile',
        '--accept=pipe,name=a0-office;urp;StarOffice.ServiceManager'], env=env, stdout=log, stderr=log))
    log.close()
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        try:
            context = resolver.resolve('uno:pipe,name=a0-office;urp;StarOffice.ComponentContext')
            provider = context.ServiceManager.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', context)
            config = provider.createInstanceWithArguments('com.sun.star.configuration.ConfigurationUpdateAccess',
                properties(nodepath='/org.openoffice.Office.Common/Misc'))
            config.setPropertyValue('ShowTipOfTheDay', False)
            config.commitChanges()
            desktop = context.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', context)
            return desktop
        except Exception:
            time.sleep(.2)
    raise RuntimeError('LibreOffice did not start; see sandbox office.log')


def document_path(request):
    identifier = str(request['file_id'])
    if not re.fullmatch(r'[a-zA-Z0-9_-]{1,80}', identifier):
        raise ValueError('Invalid document identifier')
    name = str(request['name'])
    if name != Path(name).name or Path(name).suffix.lower().lstrip('.') not in FORMATS:
        raise ValueError('Invalid Office filename')
    path = ROOT / identifier / name
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def load(path, visible=True, readonly=False):
    component = ensure_office().loadComponentFromURL(path.as_uri(), '_blank', 0,
        properties(Hidden=not visible, ReadOnly=readonly, MacroExecutionMode=NEVER_EXECUTE, UpdateDocMode=0))
    if component is None:
        raise RuntimeError('LibreOffice could not load this document')
    if visible:
        window = component.getCurrentController().getFrame().getContainerWindow()
        window.setPosSize(0, 0, display.width, display.height, 15)
        window.setVisible(True)
        window.toFront()
    return component


def saved(identifier, force=True):
    component = components.get(identifier)
    if component is None:
        return None
    listener = listeners[identifier]
    if not listener.closed and force and component.isModified():
        component.store()
    path = listener.path if listener.closed else Path(uno.fileUrlToSystemPath(component.getURL())).resolve()
    if not path.is_relative_to(ROOT) or path.suffix.lstrip('.').lower() not in FORMATS:
        raise ValueError('Save Office files inside the sandbox Documents folder, or use the Office Save As button')
    return {'dirty':False if listener.closed else component.isModified(), 'closed':listener.closed,
            'name':path.name, 'data':base64.b64encode(path.read_bytes()).decode()}


def handle(request):
    action = request.get('action')
    identifier = request.get('file_id', '')
    if action == 'template':
        data = documents.template_bytes(request.get('kind','document'), request['ext'], request.get('title',''), request.get('content',''))
        return {'data':base64.b64encode(data).decode()}
    if action == 'status':
        return {'open_ids':list(components), 'display':bool(desktop), 'viewer': display.ensure_viewer() if desktop else None}
    if action == 'resize':
        result = display.resize(request['width'], request['height'])
        for component in components.values():
            controller = component.getCurrentController()
            if controller:
                controller.getFrame().getContainerWindow().setPosSize(0,0,display.width,display.height,15)
        return result
    if action == 'retarget':
        if identifier in components:
            components[request['new_id']] = components.pop(identifier)
            listeners[request['new_id']] = listeners.pop(identifier)
        return {'ok':True}
    if action in {'save','sync','close'}:
        result = saved(identifier, force=action != 'sync') if request.get('save',True) else None
        if (action == 'close' or (result and result.get('closed'))) and identifier in components:
            component = components.pop(identifier)
            listener = listeners.pop(identifier)
            if not listener.closed:
                # Discard means discard; avoid an invisible LibreOffice save prompt.
                component.setModified(False)
                component.close(True)
        return result or {'open':False}
    if action not in {'open','read','edit','export'}:
        raise ValueError('Unsupported Office action')
    path = document_path(request)
    data = base64.b64decode(request['data'], validate=True)
    existing = components.get(identifier)
    reopen = existing is not None and action == 'edit'
    if existing is not None and action == 'open':
        frame = existing.getCurrentController().getFrame()
        frame.activate()
        window = frame.getContainerWindow()
        window.toFront()
        window.setFocus()
        return {'viewer':display.ensure_viewer()}
    if reopen:
        if existing.isModified():
            raise RuntimeError('Document changed during edit; save and retry')
        existing.close(True)
        components.pop(identifier)
        listeners.pop(identifier)
    path.write_bytes(data)
    doc = {'file_id': identifier, 'path':str(path), 'extension':path.suffix.lstrip('.')}
    documents._current[identifier] = doc
    if action == 'open':
        open_component(identifier, path, readonly=request.get('mode') == 'view')
        return {'viewer':display.ensure_viewer()}
    if action == 'read':
        return artifact_editor.read_artifact(doc, max_chars=int(request.get('max_chars') or 12000))
    if action == 'edit':
        _, details = artifact_editor.edit_artifact(doc, **request.get('edit', {}))
        if reopen:
            open_component(identifier, path)
        return {'data':base64.b64encode(path.read_bytes()).decode(), 'details':details}
    target = str(request.get('target_format','')).lower().lstrip('.')
    if target not in {*FORMATS, 'pdf'}:
        raise ValueError('Unsupported export format')
    component = load(path, visible=False) if existing is None else existing
    try:
        with tempfile.TemporaryDirectory() as directory:
            target_path = Path(directory) / ('export.' + target)
            if target == 'pdf':
                kind = 'calc' if doc['extension'] in {'ods','xlsx'} else 'impress' if doc['extension'] in {'odp','pptx'} else 'writer'
                filter_name = kind + '_pdf_Export'
            else:
                filter_name = FILTERS[target]
            component.storeToURL(target_path.as_uri(), properties(FilterName=filter_name, Overwrite=True))
            return {'data':base64.b64encode(target_path.read_bytes()).decode()}
    finally:
        if existing is None:
            component.close(True)


class Handler(socketserver.StreamRequestHandler):
    def handle(self):
        try:
            raw = self.rfile.readline(MAX_MESSAGE + 1)
            if len(raw) > MAX_MESSAGE:
                raise ValueError('Office request is too large')
            result = {'ok':True, 'result':handle(json.loads(raw))}
        except Exception as exc:
            result = {'ok':False, 'error':str(exc)}
        self.wfile.write(json.dumps(result, ensure_ascii=False).encode() + b'\n')


if __name__ == '__main__':
    socket = Path('/run/a0-office/control.sock')
    socket.unlink(missing_ok=True)
    with socketserver.UnixStreamServer(str(socket), Handler) as server:
        socket.chmod(0o600)
        ready = Path('/var/lib/a0-office/ready.json')
        ready.write_text(json.dumps({'boot':sys.argv[1]}))
        subprocess.run(['python','/plugin/sandbox/progress.py','ready','Office is ready'],check=True)
        server.serve_forever()
