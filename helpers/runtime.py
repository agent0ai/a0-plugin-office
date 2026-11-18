"""Lazy community-plugin sandbox using Agent Zero's public manager API."""
import atexit
import base64
import hashlib
import json
from pathlib import Path
import socket
import shutil
import threading
import time
import uuid
from filelock import FileLock
from helpers import files, virtual_desktop
from helpers.sandbox import get_manager, SandboxMount, SandboxError

PLUGIN = Path(__file__).resolve().parents[1]
NAME = 'service-office'
MAX_BYTES = 512 * 1024 * 1024


def ipc_directory(manager):
    scope = hashlib.sha256(str(manager.config.instances_dir).encode()).hexdigest()[:16]
    return Path('/run') / ('a0-office-' + scope)


class OfficeRuntime:
    def __init__(self):
        self.lock = threading.RLock()
        self.progress_lock = threading.Lock()
        self.progress = {'state':'idle', 'message':'Office starts on first use.'}
        self.manager = None
        self.ready = False
        self.boot = ''
        self.token = 'office-' + uuid.uuid4().hex
        self.opened = {}
        self.socket_dir = None

    def _progress(self, state, message):
        with self.progress_lock:
            self.progress = {'state':state, 'message':message}

    def status(self):
        with self.progress_lock:
            result = dict(self.progress)
        if result['state'] == 'starting' and self.manager:
            try:
                guest = json.loads(self.manager.read_file(NAME, 'var/lib/a0-office/progress.json', limit=4096))
                if guest.get('boot') == self.boot and guest.get('state') != 'ready':
                    result.update(guest)
            except (OSError, ValueError, SandboxError):
                pass
        return {**result, 'documents':[dict(item['document'], dirty=item.get('dirty',False), save_requested=item.get('save_requested',False)) for item in list(self.opened.values())]}

    def ensure(self):
        if self.ready and self.manager and self.manager.is_running(NAME):
            return
        self._progress('starting', 'Preparing Office. First use installs LibreOffice and can take several minutes.')
        manager = self.manager = get_manager()
        bundle = Path('/usr/local/lib/tuxee-gui')
        if not (bundle / 'bin/tuxee-Xvnc').is_file():
            raise RuntimeError('Office needs the shared Tuxee display bundle supplied by current Agent Zero v3 images.')
        manager.config.instances_dir.mkdir(parents=True, exist_ok=True)
        with FileLock(manager.config.instances_dir / '.office-setup.lock', timeout=1800):
            self.boot = uuid.uuid4().hex
            command = ['sh','/plugin/sandbox/boot.sh',self.boot]
            support = Path(files.get_abs_path('usr/office/runtime-support'))
            support.mkdir(parents=True,exist_ok=True)
            shutil.copyfile(files.get_abs_path('helpers/native_display.py'),support / 'native_display.py')
            self.socket_dir = ipc_directory(manager)
            self.socket_dir.mkdir(mode=0o700,parents=True,exist_ok=True)
            # Mount guest code only: ACL updates on extension files trigger A0's plugin reloader.
            mounts = [SandboxMount(PLUGIN / 'sandbox','/plugin/sandbox'), SandboxMount(bundle,'/opt/tuxee-gui'),
                      SandboxMount(support, '/support'), SandboxMount(self.socket_dir,'/run/a0-office',False)]
            if (manager.config.instances_dir / NAME / 'metadata.json').exists():
                manager.stop(NAME)
                manager.set_mounts(NAME, mounts)
                manager.set_command(NAME, command)
                manager.start(NAME)
            else:
                manager.create(NAME, mounts=mounts, command=command, service=True, share_usr=False,
                               standalone=False, sandbox_type='permanent', ownership={'plugin':'office'})
            deadline = time.monotonic() + 1800
            while time.monotonic() < deadline:
                try:
                    ready = json.loads(manager.read_file(NAME,'var/lib/a0-office/ready.json',limit=4096))
                    if ready.get('boot') == self.boot:
                        self.ready = True
                        self._progress('ready','Office is ready')
                        return
                except (OSError,ValueError,SandboxError):
                    pass
                if not manager.is_running(NAME):
                    try:
                        detail = manager.read_file(NAME,'var/lib/a0-office/startup.log',limit=6000)
                    except (OSError,SandboxError):
                        detail = 'Sandbox stopped before readiness.'
                    raise RuntimeError('Office setup failed: ' + detail)
                time.sleep(.25)
            manager.stop(NAME)
            raise TimeoutError('Office installation exceeded 30 minutes. Retry to resume.')

    def call(self, action, **payload):
        with self.lock:
            try:
                self.ensure()
                response = self.request({'action':action, **payload})
                if not response.get('ok'):
                    raise RuntimeError(response.get('error') or 'Office request failed')
                self._progress('ready','Office is ready')
                return response['result']
            except Exception as exc:
                self._progress('error',str(exc))
                raise

    def request(self, payload):
        # One private socket avoids launching PRoot and walking mount ACLs for every small request.
        data = json.dumps(payload).encode() + b'\n'
        limit = 700 * 1024 * 1024
        if len(data) > limit:
            raise ValueError('Office request exceeds the size limit')
        with socket.socket(socket.AF_UNIX) as connection:
            connection.settimeout(150)
            connection.connect(str(self.socket_dir / 'control.sock'))
            connection.sendall(data)
            with connection.makefile('rb') as stream:
                response = stream.readline(limit + 1)
        if len(response) > limit:
            raise ValueError('Office response exceeds the size limit')
        return json.loads(response)

    def payload(self, doc):
        path = Path(doc['path'])
        if path.stat().st_size > MAX_BYTES:
            raise ValueError('Office file exceeds the 512 MiB limit')
        return {'file_id':doc['file_id'], 'name':path.name, 'data':base64.b64encode(path.read_bytes()).decode()}

    def open(self, doc, mode='edit'):
        from usr.plugins.office.helpers.viewer_gateway import register
        register()
        with self.lock:
            identifier = doc['file_id']
            if identifier in self.opened:
                self.sync(identifier)
            result = self.call('open', **self.payload(doc), mode=mode)
            current = Path(doc['path']).read_bytes()
            self.opened.setdefault(identifier, {'document':doc, 'hash':hashlib.sha256(current).hexdigest(), 'dirty':False})
            viewer = result['viewer']
            virtual_desktop.register_session(token=self.token, host='127.0.0.1', port=viewer['port'], owner='office-vnc',
                title='Office', password=viewer['password'], resize=lambda width,height:self.call('resize',width=width,height=height))
            return {'ok':True, 'document':doc, 'viewer_url':f'/browser/vnc/session/{self.token}/index.html', 'token':self.token}

    def sync(self, file_id='', save=True):
        from usr.plugins.office.helpers import document_store
        with self.lock:
            selected = None
            for identifier in [file_id] if file_id else list(self.opened):
                record = self.opened.get(identifier)
                if not record:
                    continue
                result = self.call('save' if save else 'sync', file_id=identifier)
                if 'data' not in result:
                    continue
                data = base64.b64decode(result['data'], validate=True)
                doc = record['document']
                path = Path(doc['path'])
                current = path.read_bytes()
                if hashlib.sha256(current).hexdigest() != record['hash']:
                    raise RuntimeError('The original file changed outside Office. Close without saving and reopen it before continuing.')
                if result['name'] != path.name:
                    target = document_store.normalize_path(path.with_name(result['name']))
                    if target.exists():
                        raise FileExistsError(f'Save As target exists: {target}')
                    document_store._write_atomic(target, data, source=path)
                    updated = document_store.register_document(target)
                    self.call('retarget',file_id=identifier,new_id=updated['file_id'])
                    del self.opened[identifier]
                    self.opened[updated['file_id']] = record
                else:
                    updated = document_store.replace_document_bytes(identifier,data,actor='office:native')
                record.update(document=updated, hash=hashlib.sha256(data).hexdigest(), dirty=result.get('dirty',False), save_requested=bool(doc.get('draft') and result.get('save_requested')))
                if result.get('closed'):
                    self.opened.pop(updated['file_id'],None)
                selected = updated
            return {**self.status(), **({'document':selected} if file_id and selected else {})}

    def edited(self, doc):
        record = self.opened.get(doc['file_id'])
        if record:
            record.update(document=doc, hash=hashlib.sha256(Path(doc['path']).read_bytes()).hexdigest())

    def close(self, file_id, save=True):
        with self.lock:
            if file_id not in self.opened:
                return
            if save:
                result = self.sync(file_id)
                file_id = result.get('document', {}).get('file_id', file_id)
            self.call('close',file_id=file_id,save=save)
            self.opened.pop(file_id,None)

    def stop(self):
        with self.lock:
            if self.ready:
                self.sync()
            virtual_desktop.unregister_session(self.token)
            if self.manager and (self.manager.config.instances_dir / NAME / 'metadata.json').exists():
                self.manager.stop(NAME)
            if self.socket_dir and self.socket_dir.exists():
                # Retain the mount directory: the manager validates it before restart/delete.
                for child in self.socket_dir.iterdir():
                    if child.is_dir() and not child.is_symlink():
                        shutil.rmtree(child)
                    else:
                        child.unlink(missing_ok=True)
            self.ready = False
            self.opened.clear()
            self._progress('idle','Office stopped; installed packages are retained.')


_runtime = OfficeRuntime()
def get_runtime():
    return _runtime

atexit.register(_runtime.stop)
