"""Explicit Playwright smoke test in a disposable A0 container with a running WebUI.
Usage: python -m usr.plugins.office.tests.live_ui http://localhost:<discovered-port> /tmp/office-shots
Playwright and its Chromium shell are test-only dependencies, not plugin dependencies.
"""
import asyncio
from pathlib import Path
import sys
import zipfile
from playwright.async_api import async_playwright

async def main():
    url=sys.argv[1]; output=Path(sys.argv[2]); output.mkdir(parents=True,exist_ok=True)
    async with async_playwright() as p:
        browser=await p.chromium.launch(headless=True,args=['--no-sandbox'])
        page=await browser.new_page(viewport={'width':1440,'height':1000})
        errors=[]; page.on('pageerror',lambda error:errors.append(str(error)))
        await page.goto(url)
        await page.wait_for_function('!!globalThis.Alpine')
        await page.wait_for_timeout(3500)
        await page.evaluate("async()=>{const {callJsonApi}=await import('/js/api.js');return callJsonApi('plugins',{action:'toggle_plugin',plugin_name:'office',enabled:true})}")
        await page.reload()
        await page.wait_for_function('!!globalThis.Alpine')
        await page.wait_for_timeout(3500)
        await page.evaluate("async()=>{const s=await import('/js/surfaces.js');await s.open('office')}")
        await page.wait_for_function("!!globalThis.Alpine?.store('office')")
        await page.locator('.office-panel:visible').wait_for()
        async def api(payload):
            result=await page.evaluate("async payload=>{const {callJsonApi}=await import('/js/api.js');return callJsonApi('plugins/office/office_session',payload)}",payload)
            assert result.get('ok') is not False,result
            return result
        async def create(fmt,title,content):
            await page.wait_for_function("!Alpine.store('office').busy")
            result=await page.evaluate("async payload=>Alpine.store('office').run(payload)",{'action':'create','format':fmt,'title':title,'content':content})
            assert result,result
            await page.wait_for_timeout(1000)
            frame=next(f for f in page.frames if '/browser/vnc/' in f.url)
            await frame.wait_for_function('window.a0OfficeViewer?.connected',timeout=30000)
            await page.wait_for_timeout(1000)
            return result['document'],frame
        writer,frame=await create('odt','Office for Agent Zero','Native Writer, Calc and Impress\nA dedicated sandbox. A shared display runtime.\nDocuments stay in your A0 workspace.\nAgent edits and native editing work together.')
        await frame.evaluate("""()=>{const r=window.a0OfficeViewer.rfb;r.sendKey(0xffe3,'ControlLeft',true);r.sendKey(0xff57,'End');r.sendKey(0xffe3,'ControlLeft',false);r.sendKey(0xff0d,'Enter');window.a0OfficeViewer.paste('Verified clipboard and keyboard input.');}""")
        await page.wait_for_timeout(1200)
        await api({'action':'save','file_id':writer['file_id']})
        with zipfile.ZipFile(writer['path']) as z:
            assert b'Verified clipboard and keyboard input.' in z.read('content.xml')
        print('PASS VNC connection, keyboard, clipboard, native save',flush=True)
        await page.screenshot(path=str(output/'writer.png'))
        calc,_=await create('ods','Quarterly budget','Category\tPlanned\tActual\nEngineering\t12000\t11500\nDesign\t5000\t4800\nOperations\t4000\t3900\nMarketing\t3000\t2750\nTotal\t24000\t22950')
        await page.screenshot(path=str(output/'calc.png'))
        impress,_=await create('odp','Office without the extra desktop','Office for Agent Zero\nWriter, Calc and Impress\nDependencies install on first use\nOne shared Office sandbox\nYour documents stay in your workspace')
        await page.screenshot(path=str(output/'impress.png'))
        print('PASS Writer, Calc and Impress rendered',flush=True)
        # Docking must retain or reconnect the live viewer and resize the desktop.
        await page.evaluate("async()=>{const {store}=await import('/components/sidebar/chats/chats-store.js');await store.newChat()}")
        await page.wait_for_function("Alpine.store('rightCanvas').shouldRender()")
        assert await page.evaluate("async()=>{const s=await import('/js/surfaces.js');return s.dock('office')}")
        await page.wait_for_timeout(2500)
        await page.set_viewport_size({'width':1600,'height':1000})
        await page.wait_for_timeout(1800)
        frame=next(f for f in page.frames if '/browser/vnc/' in f.url)
        await frame.wait_for_function('window.a0OfficeViewer?.connected',timeout=30000)
        await page.screenshot(path=str(output/'docked.png'))
        # Browser's toggle must not gate Office's display.
        await page.evaluate("async()=>{const {callJsonApi}=await import('/js/api.js');return callJsonApi('plugins',{action:'toggle_plugin',plugin_name:'_browser',enabled:false})}")
        await frame.goto(frame.url)
        await frame.wait_for_function('window.a0OfficeViewer?.connected',timeout=30000)
        print('PASS docking, resize, reconnect, Browser disabled',flush=True)
        await api({'action':'open','file_id':writer['file_id']})
        await page.wait_for_timeout(600)
        await frame.evaluate("""()=>{const r=window.a0OfficeViewer.rfb;r.sendKey(0xffe3,'ControlLeft',true);r.sendKey(0x77,'KeyW');r.sendKey(0xffe3,'ControlLeft',false);}""")
        await page.wait_for_timeout(1200)
        result=await api({'action':'sync'})
        assert writer['file_id'] not in [doc['file_id'] for doc in result['documents']],result
        print('PASS native window close synchronizes tabs',flush=True)
        await page.evaluate("async()=>{const {callJsonApi}=await import('/js/api.js');return callJsonApi('plugins',{action:'toggle_plugin',plugin_name:'office',enabled:false})}")
        from helpers.sandbox import get_manager
        assert not get_manager().is_running('service-office')
        denied=await page.evaluate("async()=>{const {callJsonApi}=await import('/js/api.js');return callJsonApi('plugins/office/office_session',{action:'status'})}")
        assert denied.get('ok') is False
        await page.evaluate("async()=>{const {callJsonApi}=await import('/js/api.js');return callJsonApi('plugins',{action:'toggle_plugin',plugin_name:'office',enabled:true})}")
        await api({'action':'open','file_id':calc['file_id']})
        await api({'action':'stop'})
        print('PASS disable stops sandbox, blocks API, re-enable reuses installation',flush=True)
        assert not errors,errors
        await browser.close()

if __name__=='__main__': asyncio.run(main())
