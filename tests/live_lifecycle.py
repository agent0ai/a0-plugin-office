"""Destructive test for a DISPOSABLE A0 container: uninstall Office, then install a ZIP.
Usage: python /tmp/live_lifecycle.py URL /tmp/office.zip
Copy this script outside the plugin before running, since uninstall deletes the plugin.
"""
import asyncio
import base64
import hashlib
from pathlib import Path
import sys
from playwright.async_api import async_playwright

async def main():
    url,package=sys.argv[1:]
    payload=base64.b64encode(Path(package).read_bytes()).decode()
    from helpers.sandbox import get_manager
    manager=get_manager()
    documents=Path('/a0/usr/workdir/documents')
    snapshots={p:hashlib.sha256(p.read_bytes()).hexdigest() for p in documents.glob('*') if p.is_file()}
    history=Path('/a0/usr/office/documents/documents.sqlite3')
    assert snapshots and history.exists()
    async with async_playwright() as p:
        browser=await p.chromium.launch(headless=True,args=['--no-sandbox'])
        page=await browser.new_page()
        await page.goto(url)
        await page.wait_for_function('!!globalThis.Alpine')
        await page.wait_for_timeout(2000)
        async def api(endpoint,body):
            result=await page.evaluate("async ([endpoint,body])=>{const {callJsonApi}=await import('/js/api.js');return callJsonApi(endpoint,body)}",[endpoint,body])
            assert result.get('ok') is not False,result
            return result
        # Exercise transport admission against an actual registered viewer.
        doc=next(path for path in snapshots if path.suffix=='.odt')
        opened=await api('plugins/office/office_session',{'action':'open','path':str(doc)})
        credentials=opened['viewer_url'].replace('index.html','credentials')
        assert (await page.request.get(url+credentials)).status==200
        assert (await page.request.get(url+credentials,headers={'Origin':'https://other.example'})).status==403
        assert (await page.request.post(url+'/api/plugins/office/office_session',headers={'X-CSRF-Token':'invalid'},data={'action':'stop'})).status==403
        await api('plugins/office/office_session',{'action':'stop'})
        print('PASS live same-origin viewer admission and CSRF rejection',flush=True)
        await api('plugins',{'action':'delete_plugin','plugin_name':'office'})
        assert not Path('/a0/usr/plugins/office').exists()
        assert not (manager.config.instances_dir/'service-office').exists()
        assert history.exists()
        assert all(path.exists() and hashlib.sha256(path.read_bytes()).hexdigest()==digest for path,digest in snapshots.items())
        print('PASS uninstall removes plugin and sandbox, preserves documents and history',flush=True)
        installed=await page.evaluate("""async data=>{
            const {fetchApi}=await import('/js/api.js');
            const form=new FormData();form.append('action','install_zip');
            form.append('plugin_file',new File([Uint8Array.from(atob(data),ch=>ch.charCodeAt(0))],'office.zip'));
            return (await fetchApi('plugins/_plugin_installer/plugin_install',{method:'POST',body:form})).json();
        }""",payload)
        assert installed.get('success'),installed
        assert Path('/a0/usr/plugins/office/plugin.yaml').exists()
        result=await api('plugins/office/office_session',{'action':'status'})
        assert result['state']=='idle'
        assert not (manager.config.instances_dir/'service-office').exists()
        print('PASS normal ZIP installation, install hook, API discovery, lazy dependency setup',flush=True)
        await browser.close()

if __name__=='__main__': asyncio.run(main())
