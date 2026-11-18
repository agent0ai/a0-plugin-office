import asyncio
import uuid
from pathlib import Path
from helpers.api import ApiHandler, Request
from helpers import plugins
from usr.plugins.office.helpers import document_store, libreoffice
from usr.plugins.office.helpers.runtime import get_runtime

FORMATS = {'odt','ods','odp','docx','xlsx','pptx'}

class OfficeSession(ApiHandler):
    async def process(self, input: dict, request: Request) -> dict:
        if plugins.get_toggle_state('office') != 'enabled':
            return {'ok':False,'error':'Office is disabled'}
        if input.get('action') == 'status':
            return {'ok':True, **get_runtime().status()}
        try:
            return await asyncio.to_thread(dispatch, input)
        except Exception as exc:
            return {'ok':False,'error':str(exc)}


def dispatch(input):
    action = str(input.get('action') or 'open')
    runtime = get_runtime()
    context_id = str(input.get('context_id') or input.get('ctxid') or '')
    if action == 'sync':
        return {'ok':True, **(runtime.sync(save=False) if runtime.ready else runtime.status())}
    if action == 'stop':
        runtime.stop()
        return {'ok':True, **runtime.status()}
    if action == 'resize':
        return runtime.call('resize',width=input.get('width'),height=input.get('height'))
    if action == 'browse':
        path = document_store.normalize_path(input.get('path') or document_store.default_open_path(context_id),context_id=context_id)
        entries=[]
        for child in sorted(path.iterdir() if path.exists() else [],key=lambda item:(not item.is_dir(),item.name.lower())):
            if child.name == '.office-drafts':
                continue
            if child.is_dir() or child.suffix.lower().lstrip('.') in FORMATS:
                entries.append({'name':child.name,'path':str(child),'directory':child.is_dir()})
        try:
            parent=document_store.normalize_path(path.parent,context_id=context_id)
        except PermissionError:
            parent=path
        return {'ok':True,'path':str(path),'parent':str(parent),'entries':entries}
    if action == 'home':
        return {'ok':True,'path':document_store.default_open_path(context_id)}
    with runtime.lock:
        if action == 'create':
            fmt = str(input.get('format') or 'odt')
            if fmt not in FORMATS:
                raise ValueError('Unsupported Office format')
            path = input.get('path','')
            if input.get('draft'):
                path = str(document_store.document_home(context_id) / '.office-drafts' / uuid.uuid4().hex / ('Untitled.' + fmt))
            doc = document_store.create_document(kind=input.get('kind','document'),title='' if input.get('draft') else input.get('title') or 'Untitled',
                fmt=fmt, content=input.get('content',''),path=path,context_id=context_id)
        else:
            doc = document_store.get_document(input['file_id']) if input.get('file_id') else document_store.register_document(input.get('path',''), context_id=context_id)
        if doc['extension'] not in FORMATS:
            raise ValueError('Text documents use the Editor surface')
        identifier = doc['file_id']
        if action in {'create','open'}:
            result = runtime.open(doc,mode=input.get('mode','edit'))
            return {**result, 'documents':runtime.status()['documents']}
        if action == 'close':
            runtime.close(identifier,save=input.get('save',True))
            return {'ok':True, **runtime.status()}
        if action == 'versions':
            return {'ok':True,'versions':document_store.version_history(identifier)}
        synced = runtime.sync(identifier)
        identifier = synced.get('document', {}).get('file_id', identifier)
        doc = document_store.get_document(identifier)
        if action == 'save':
            return {'ok':True,'document':doc}
        if action == 'rename':
            target = document_store.normalize_path(input['target_path'],context_id=context_id)
            was_open = identifier in runtime.opened
            if was_open:
                runtime.close(identifier)
            updated = document_store.rename_document(identifier,target,context_id=context_id)
            if was_open:
                return runtime.open(updated)
            return {'ok':True,'document':updated}
        if action == 'save_as':
            target = document_store.normalize_path(input['target_path'],context_id=context_id)
            if target.suffix.lstrip('.') != doc['extension']:
                raise ValueError('Save As keeps the format; use Export to convert')
            if target.exists():
                raise FileExistsError(str(target))
            if doc.get('draft'):
                runtime.close(identifier)
                updated = document_store.rename_document(identifier,target,context_id=context_id)
            else:
                document_store._write_atomic(target,Path(doc['path']).read_bytes(),source=Path(doc['path']))
                updated = document_store.register_document(target,context_id=context_id)
            result = runtime.open(updated)
            if not doc.get('draft'):
                runtime.close(identifier)
            return {**result, 'documents':runtime.status()['documents']}
        if action == 'export':
            return libreoffice.convert_document(doc['path'],input.get('target_format','pdf'))
        if action == 'restore':
            was_open = identifier in runtime.opened
            if was_open:
                runtime.close(identifier)
            restored = document_store.restore_version(identifier,int(input['version_id']))
            if was_open:
                return runtime.open(restored)
            return {'ok':True,'document':restored}
    raise ValueError('Unsupported Office action')
