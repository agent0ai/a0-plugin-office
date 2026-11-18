"""Run in an A0 checkout with this plugin installed under usr/plugins/office."""
import base64
import hashlib
from pathlib import Path
from types import SimpleNamespace

import pytest
from usr.plugins.office.helpers import document_store as ds
from usr.plugins.office.helpers.runtime import OfficeRuntime
from usr.plugins.office.helpers.viewer_gateway import OfficeVncGateway


@pytest.fixture
def documents(tmp_path, monkeypatch):
    state = tmp_path / 'state'
    monkeypatch.setattr(ds, 'STATE_DIR', state)
    monkeypatch.setattr(ds, 'BACKUP_DIR', state / 'versions')
    monkeypatch.setattr(ds, 'DB_PATH', state / 'documents.sqlite3')
    monkeypatch.setattr(ds, 'allowed_roots', lambda *a, **kw: [tmp_path])
    def create(name='sample.odt', data=b'original'):
        path = tmp_path / name
        path.write_bytes(data)
        return ds.register_document(path)
    return create


def native_runtime(doc, data, name=None):
    runtime = OfficeRuntime()
    runtime.opened[doc['file_id']] = {'document':doc, 'hash':hashlib.sha256(Path(doc['path']).read_bytes()).hexdigest()}
    calls=[]
    def call(action, **payload):
        calls.append((action,payload))
        return {'data':base64.b64encode(data).decode(), 'name':name or doc['basename'], 'dirty':False}
    runtime.call=call
    return runtime,calls


def test_versions_restore_and_file_permissions(documents):
    doc=documents()
    path=Path(doc['path']); path.chmod(0o640)
    ds.replace_document_bytes(doc['file_id'],b'changed')
    version=ds.version_history(doc['file_id'])[0]
    assert Path(version['path']).stat().st_mode & 0o777 == 0o600
    ds.restore_version(doc['file_id'],version['id'])
    assert path.read_bytes()==b'original'
    assert path.stat().st_mode & 0o777 == 0o640


def test_native_sync_preserves_external_changes(documents):
    doc=documents(); runtime,_=native_runtime(doc,b'native')
    before=ds.version_history(doc['file_id'])
    path=Path(doc['path']); path.write_bytes(b'external')
    with pytest.raises(RuntimeError, match='outside Office'):
        runtime.sync(doc['file_id'])
    assert path.read_bytes()==b'external'
    assert ds.version_history(doc['file_id'])==before
    runtime.close(doc['file_id'],save=False)
    assert not runtime.opened


def test_native_save_as_and_close_use_new_identifier(documents):
    doc=documents(); runtime,calls=native_runtime(doc,b'new file','copy.odt')
    runtime.close(doc['file_id'])
    assert Path(doc['path']).read_bytes()==b'original'
    assert Path(doc['path']).with_name('copy.odt').read_bytes()==b'new file'
    new_id=[args['new_id'] for action,args in calls if action=='retarget'][0]
    assert calls[-1]==('close',{'file_id':new_id,'save':True})
    assert not runtime.opened


def test_native_save_as_refuses_existing_file(documents):
    doc=documents(); other=documents('copy.odt',b'keep')
    runtime,_=native_runtime(doc,b'new','copy.odt')
    with pytest.raises(FileExistsError): runtime.sync(doc['file_id'])
    assert Path(other['path']).read_bytes()==b'keep'


def test_native_poll_does_not_create_redundant_versions(documents):
    doc=documents(); runtime,_=native_runtime(doc,b'original')
    before=ds.version_history(doc['file_id'])
    runtime.sync(save=False); runtime.sync(save=False)
    assert ds.version_history(doc['file_id'])==before


def test_scoped_paths_reject_escape_and_symlink(documents,tmp_path):
    with pytest.raises(PermissionError): ds.normalize_path('/etc/passwd')
    (tmp_path/'escape').symlink_to('/etc')
    with pytest.raises(PermissionError): ds.normalize_path(tmp_path/'escape/passwd')


def test_failed_atomic_replace_preserves_original(documents,monkeypatch):
    doc=documents()
    def fail(*a): raise OSError('disk error')
    monkeypatch.setattr(ds.os,'replace',fail)
    with pytest.raises(OSError): ds.replace_document_bytes(doc['file_id'],b'new')
    assert Path(doc['path']).read_bytes()==b'original'
    assert not list(Path(doc['path']).parent.glob('.*.tmp'))


def test_viewer_admission(monkeypatch):
    from usr.plugins.office.helpers import viewer_gateway as gateway
    viewer=OfficeVncGateway(None)
    monkeypatch.setattr(gateway.plugins,'get_toggle_state',lambda name:'enabled' if name=='office' else 'disabled')
    monkeypatch.setattr(viewer,'is_authorized',lambda scope:True)
    scope={'type':'websocket','headers':[(b'host',b'localhost:5000'),(b'origin',b'http://localhost:5000')]}
    assert viewer.allowed(scope)  # Browser plugin may be disabled.
    assert not viewer.allowed({**scope,'headers':[(b'host',b'localhost:5000'),(b'origin',b'https://other.test')]})
    assert not viewer.allowed({**scope,'headers':[(b'host',b'localhost:5000')]})
    monkeypatch.setattr(viewer,'is_authorized',lambda scope:False)
    assert not viewer.allowed(scope)


def test_viewer_only_accepts_owned_loopback_tokens(monkeypatch):
    from usr.plugins.office.helpers import viewer_gateway as gateway
    viewer=OfficeVncGateway(None)
    scope={'path':'/browser/vnc/session/office-test/ws'}
    for owner,host in [('browser-vnc','127.0.0.1'),('office-vnc','192.0.2.1')]:
        monkeypatch.setattr(gateway.virtual_desktop,'proxy_for_token',lambda token:SimpleNamespace(owner=owner,host=host))
        assert viewer.endpoint(scope)==(None,'')


def test_setup_retry_clears_error_state(monkeypatch):
    runtime=OfficeRuntime()
    runtime.ensure=lambda:None
    runtime._progress('error','previous failure')
    runtime.request=lambda payload:{'ok':True,'result':{}}
    runtime.call('status')
    assert runtime.status()['state']=='ready'


def test_registered_path_is_revalidated(documents):
    doc=documents()
    path=Path(doc['path']); path.unlink(); path.symlink_to('/etc/passwd')
    with pytest.raises(PermissionError): ds.get_document(doc['file_id'])


def test_stop_clears_rpc_files_but_retains_mount(tmp_path):
    runtime=OfficeRuntime()
    runtime.socket_dir=tmp_path/'ipc'; runtime.socket_dir.mkdir()
    (runtime.socket_dir/'control.sock').touch()
    (runtime.socket_dir/'office-display').mkdir()
    metadata=tmp_path/'service-office'/'metadata.json'
    metadata.parent.mkdir(); metadata.write_text('{}')
    stopped=[]
    runtime.manager=SimpleNamespace(config=SimpleNamespace(instances_dir=tmp_path),stop=stopped.append)
    runtime.stop()
    assert stopped==['service-office']
    assert runtime.socket_dir.is_dir() and not list(runtime.socket_dir.iterdir())


def test_export_result_identifies_the_exported_file(documents,monkeypatch):
    import json
    from usr.plugins.office.tools.office_artifact import OfficeArtifact
    from usr.plugins.office.helpers import libreoffice
    original=documents()
    exported=documents('sample.pdf',b'%PDF-1.7\n')
    monkeypatch.setattr(libreoffice,'convert_document',lambda *a:{'ok':True,'path':exported['path']})
    tool=OfficeArtifact(agent=None,name='office_artifact',method=None,args={},message='',loop_data=None)
    result=tool._execute(action='export',file_id=original['file_id'],target_format='pdf')
    payload=json.loads(result.message)
    assert payload['document']['file_id']==exported['file_id']
    assert payload['document']['extension']=='pdf'
    assert result.additional['format']=='pdf'
