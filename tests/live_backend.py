"""Explicit integration check; creates test documents in a disposable A0 container.
Run from /a0: python -m usr.plugins.office.tests.live_backend
Stop Office in the WebUI before running (this owns its own runtime).
"""
import asyncio
from concurrent.futures import ThreadPoolExecutor
import json
import subprocess
from pathlib import Path
import time
from usr.plugins.office.helpers import artifact_editor, document_store as ds, libreoffice
from usr.plugins.office.helpers.runtime import get_runtime
from usr.plugins.office.api.office_session import dispatch
from usr.plugins.office.tools.office_artifact import OfficeArtifact


def main():
    runtime=get_runtime()
    started=time.monotonic()
    with ThreadPoolExecutor(3) as pool:
        jobs=[pool.submit(runtime.call,'template',ext='odt',title='Concurrent',content='Hello') for _ in range(3)]
        while not all(job.done() for job in jobs):
            now=time.monotonic(); runtime.status(); assert time.monotonic()-now<1
            time.sleep(.1)
        assert all(job.result()['data'] for job in jobs)
    print('Concurrent ready, seconds:',round(time.monotonic()-started,2),flush=True)
    for ext,kind,content in [('odt','document','Hello Office'),('docx','document','Hello Office'),
        ('ods','spreadsheet','Month\tRevenue\nJanuary\t120\nFebruary\t180'),
        ('xlsx','spreadsheet','Month\tRevenue\nJanuary\t120\nFebruary\t180'),
        ('odp','presentation','Hello Office'),('pptx','presentation','Hello Office')]:
        doc=ds.create_document(kind=kind,title='Backend '+ext,fmt=ext,content=content)
        assert artifact_editor.read_artifact(doc)
        if kind=='document': edit={'operation':'append_text','content':'Verified change'}
        elif kind=='spreadsheet': edit={'operation':'set_cells','cells':{'B4':'=SUM(B2:B3)'}}
        else: edit={'operation':'append_slide','slides':[{'title':'Verified slide','body':'Hello from Office'}]}
        doc,details=artifact_editor.edit_artifact(doc,**edit)
        assert details['changed'],details
        assert 'Verified' in json.dumps(artifact_editor.read_artifact(doc)) or kind=='spreadsheet'
        if ext=='xlsx':
            doc,details=artifact_editor.edit_artifact(doc,operation='create_chart',chart={'type':'column','title':'Revenue','data_range':'A1:B3','categories':'A2:A3'})
            assert details['changed']
        runtime.open(doc)
        runtime.sync(doc['file_id'])
        pdf=libreoffice.convert_document(doc['path'],'pdf')
        assert Path(pdf['path']).read_bytes().startswith(b'%PDF')
        paired={'odt':'docx','docx':'odt','ods':'xlsx','xlsx':'ods','odp':'pptx','pptx':'odp'}[ext]
        converted=libreoffice.convert_document(doc['path'],paired)
        assert artifact_editor.read_artifact(ds.register_document(converted['path']))
        versions=ds.version_history(doc['file_id'])
        assert versions
        dispatch({'action':'restore','file_id':doc['file_id'],'version_id':versions[-1]['id']})
        runtime.close(doc['file_id'])
        print('PASS',ext,'create/read/edit/native open/PDF/restore/close',flush=True)
    tool=OfficeArtifact(agent=None,name='office_artifact',method=None,args={},message='',loop_data=None)
    result=asyncio.run(tool.execute(action='create',title='Tool document',content='Created through tool',format='odt'))
    result=json.loads(result.message); assert result['ok']
    identifier=result['document']['file_id']
    renamed=dispatch({'action':'rename','file_id':identifier,'target_path':str(Path(result['document']['path']).with_name('Renamed '+identifier[:8]+'.odt'))})
    runtime.open(renamed['document'])
    copied=dispatch({'action':'save_as','file_id':identifier,'target_path':str(Path(result['document']['path']).with_name('Copy '+identifier[:8]+'.odt'))})
    assert copied['document']['file_id']!=identifier
    print('PASS tool dispatch, rename and Save As',flush=True)
    native=copied['document']
    target='/var/lib/a0-office/files/'+native['file_id']+'/Native Save As '+identifier[:8]+'.odt'
    code="""import uno
from com.sun.star.beans import PropertyValue
local=uno.getComponentContext()
resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
ctx=resolver.resolve('uno:pipe,name=a0-office;urp;StarOffice.ComponentContext')
desktop=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
for doc in desktop.getComponents():
    if doc.getURL().endswith(SUFFIX):
        doc.storeAsURL(uno.systemPathToFileUrl(TARGET),(PropertyValue(Name='FilterName',Value='writer8'),))
        break
else: raise AssertionError('Native document not found')
""".replace('SUFFIX',repr('/'+native['file_id']+'/'+Path(native['path']).name.replace(' ','%20'))).replace('TARGET',repr(target))
    subprocess.run(runtime.manager.process_command('service-office',['/usr/bin/python3','-c',code]),check=True,timeout=30)
    synced=runtime.sync(native['file_id'])['document']
    assert synced['file_id']!=native['file_id'] and Path(synced['path']).name==Path(target).name
    assert Path(native['path']).exists()
    runtime.close(synced['file_id'])
    print('PASS native UNO Save As preserves original and registers new host file',flush=True)
    runtime.stop()
    start=time.monotonic(); runtime.call('template',ext='odt',title='Warm',content='Warm')
    print('Warm restart seconds:',round(time.monotonic()-start,2),flush=True)
    print('Sandbox private bytes:',runtime.manager.disk_usage('service-office'),flush=True)
    runtime.stop()

if __name__=='__main__': main()
