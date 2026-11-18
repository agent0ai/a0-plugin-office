"""Real UI regression check; run only in a disposable A0 container.
Temporarily fails the guest installer to verify error reporting and retry.
Usage: python -m usr.plugins.office.tests.live_ui_fixes URL SCREENSHOT_DIRECTORY
"""
import asyncio
from pathlib import Path
import sys
import time
import zipfile
from playwright.async_api import async_playwright


async def main():
    suffix=str(int(time.time()))
    review=f'Office review {suffix}.odt'
    copy=f'Office copy {suffix}.odt'
    output=Path(sys.argv[2]); output.mkdir(parents=True,exist_ok=True)
    installer=Path(__file__).resolve().parents[1]/'sandbox/install.sh'
    original=installer.read_text()
    # Current A0 redirects registered main.html paths into Canvas. Exercise the
    # exact standalone markup too, as used by plugin previews/older v3 builds.
    standalone=installer.parents[1]/'webui/standalone-test.html'
    standalone.write_text((standalone.parent/'main.html').read_text())
    async with async_playwright() as p:
        browser=await p.chromium.launch(headless=True,args=['--no-sandbox'])
        page=await browser.new_page(viewport={'width':1440,'height':1000})
        errors=[]; page.on('pageerror',lambda error:errors.append(str(error)))
        await page.goto(sys.argv[1])
        await page.wait_for_function('!!globalThis.Alpine')
        await page.wait_for_timeout(3000)
        await page.evaluate("async()=>{const {callJsonApi}=await import('/js/api.js');return callJsonApi('plugins',{action:'toggle_plugin',plugin_name:'office',enabled:true})}")
        await page.reload()
        await page.wait_for_function("typeof openModal==='function' && !!globalThis.Alpine?.store('rightCanvas')",timeout=60000)
        await page.wait_for_timeout(1500)
        await page.evaluate("async()=>{const {callJsonApi}=await import('/js/api.js');return callJsonApi('plugins/office/office_session',{action:'stop'})}")
        await page.evaluate("async()=>{const {store}=await import('/components/sidebar/chats/chats-store.js');store.deselectChat()}")
        # This is the direct plugin modal entry, not just the unified surface path.
        await page.evaluate("()=>{void openModal('/plugins/office/webui/standalone-test.html')}")
        panel=page.locator('.office-panel:visible')
        await panel.wait_for()
        await page.wait_for_timeout(1800)
        assert (await panel.bounding_box())['height']>600
        assert await panel.get_by_label('Document name',exact=True).count()==0
        assert not await panel.get_by_role('button',name='Save',exact=True).is_visible()
        dock=page.locator('.office-modal .surface-dock-button')
        assert await dock.is_disabled()
        await page.screenshot(path=str(output/'empty.png'))

        async def new_writer():
            await panel.get_by_role('button',name='File',exact=True).click()
            await panel.get_by_role('button',name='New',exact=True).click()
            await panel.get_by_role('button',name='Writer document',exact=True).filter(visible=True).click()

        installer.write_text('#!/bin/sh\necho "Test installation failure: mirror unavailable"\nexit 42\n')
        try:
            await page.wait_for_timeout(2000)
            await new_writer()
            await page.wait_for_function("!Alpine.store('office').busy",timeout=900000)
            status=await page.evaluate("Alpine.store('office').startup")
            assert status['state']=='error' and 'mirror unavailable' in status['message'],status
            assert not await panel.locator('.office-progress').is_visible()
            await page.locator('.toast-message').filter(has_text='mirror unavailable').first.wait_for()
            await page.wait_for_timeout(300)
            await page.screenshot(path=str(output/'install-error.png'))
        finally:
            installer.write_text(original)
        await page.wait_for_timeout(2000)
        while await page.locator('.toast-dismiss:visible').count():
            await page.locator('.toast-dismiss:visible').first.click()
        print('PASS direct modal sizing, empty toolbar, disabled docking, actual failed setup and visible cause',flush=True)
        await new_writer()
        await page.wait_for_timeout(1500)
        await page.screenshot(path=str(output/'installing.png'))
        await page.wait_for_function("!Alpine.store('office').busy",timeout=900000)
        state=await page.evaluate("({docs:Alpine.store('office').documents,error:Alpine.store('office').lastError})")
        assert state['docs'] and state['docs'][0]['draft'],state
        await page.wait_for_function("document.querySelector('.office-host iframe')?.contentWindow?.a0OfficeViewer?.connected",timeout=30000)
        frame=next(f for f in page.frames if '/browser/vnc/' in f.url)
        await frame.wait_for_function('window.a0OfficeViewer?.connected',timeout=30000)
        await page.wait_for_timeout(1500)
        assert await panel.locator('.office-save').evaluate("el=>[...el.childNodes].filter(n=>n.nodeType===Node.TEXT_NODE).every(n=>!n.textContent.trim())")
        with zipfile.ZipFile(state['docs'][0]['path']) as archive:
            from xml.etree import ElementTree
            assert not ''.join(ElementTree.fromstring(archive.read('content.xml')).itertext()).strip()
        await frame.evaluate("text=>window.a0OfficeViewer.paste(text)", "Office for Agent Zero\nA familiar document workflow.\nVerified in the real application.")
        await page.wait_for_timeout(1000)
        await panel.get_by_role('button',name='File',exact=True).click()
        await panel.get_by_role('button',name='New',exact=True).click()
        await page.screenshot(path=str(output/'file-new.png'))
        await page.keyboard.press('Escape')
        await page.keyboard.press('Escape')
        await panel.get_by_role('button',name='Save',exact=True).filter(visible=True).click()
        filename=page.get_by_role('textbox',name='File name',exact=True)
        await filename.wait_for()
        await filename.fill(review)
        await page.screenshot(path=str(output/'first-save.png'))
        await page.get_by_role('button',name='Cancel',exact=True).filter(visible=True).click()
        assert await page.evaluate("Alpine.store('office').documents[0].draft")
        # Ctrl+S inside the native viewer must use the same first-save dialog.
        await frame.locator('canvas').click(position={'x':400,'y':300})
        await page.keyboard.press('Control+s')
        await filename.wait_for()
        await filename.fill(review)
        await page.get_by_role('button',name='Save Here',exact=True).click()
        await filename.wait_for(state='hidden')
        await page.wait_for_function("name=>Alpine.store('office').documents.some(d=>d.basename===name&&!d.draft)",arg=review)
        doc=await page.evaluate("Alpine.store('office').documents[0]")
        with zipfile.ZipFile(doc['path']) as archive:
            assert b'Verified in the real application' in archive.read('content.xml')
        await page.wait_for_timeout(1500)
        await page.mouse.move(5,5)
        await page.wait_for_timeout(800)
        await page.screenshot(path=str(output/'writer.png'))
        print('PASS setup retry, native editing, first-save cancellation, Ctrl+S, saved bytes',flush=True)
        await panel.get_by_role('button',name='File',exact=True).click()
        await panel.get_by_role('button',name='Save As',exact=True).click()
        await filename.wait_for()
        await filename.fill(copy)
        await page.get_by_role('button',name='Save Here',exact=True).click()
        await filename.wait_for(state='hidden')
        assert Path(doc['path']).exists()
        assert await page.evaluate("Alpine.store('office').documents.length")==1
        await panel.get_by_role('button',name='File',exact=True).click()
        await panel.get_by_role('button',name='Open',exact=True).click()
        picker=page.locator('.office-open-modal:visible')
        await picker.wait_for()
        assert (await picker.bounding_box())['height']>400
        await page.screenshot(path=str(output/'open.png'))
        await page.set_viewport_size({'width':480,'height':800})
        await page.wait_for_timeout(1200)
        assert (await picker.bounding_box())['width']<=480
        await page.screenshot(path=str(output/'open-narrow.png'))
        await picker.get_by_role('button',name=review,exact=True).click()
        await picker.wait_for(state='hidden')
        await page.set_viewport_size({'width':1440,'height':1000})
        # Draft close must respect cancellation, discard, and first-save naming.
        await new_writer()
        await page.wait_for_function("!Alpine.store('office').busy")
        draft_id=await page.evaluate("Alpine.store('office').active")
        async def close_active():
            await panel.get_by_role('button',name='File',exact=True).click()
            await panel.get_by_role('button',name='Close document',exact=True).click()
        await close_active()
        await page.get_by_role('button',name='Cancel',exact=True).filter(visible=True).click()
        assert await page.evaluate("id=>Alpine.store('office').documents.some(d=>d.file_id===id)",draft_id)
        await close_active()
        await page.get_by_role('button',name="Don't Save",exact=True).click()
        await page.wait_for_function("id=>!Alpine.store('office').documents.some(d=>d.file_id===id)",arg=draft_id)
        await page.wait_for_function("!Alpine.store('office').busy")
        await new_writer()
        await page.wait_for_function("!Alpine.store('office').busy")
        draft_id=await page.evaluate("Alpine.store('office').active")
        await close_active()
        await page.locator('.modal-footer').get_by_role('button',name='Save',exact=True).filter(visible=True).click()
        await filename.wait_for()
        await filename.fill(f'Closed draft {suffix}.odt')
        await page.get_by_role('button',name='Save Here',exact=True).click()
        await page.wait_for_function("id=>!Alpine.store('office').documents.some(d=>d.file_id===id)",arg=draft_id)
        await page.wait_for_function("!Alpine.store('office').busy")
        print('PASS draft close Cancel, Don’t Save, and Save',flush=True)
        await page.evaluate("async()=>{const {store}=await import('/components/sidebar/chats/chats-store.js');await store.newChat()}")
        await page.wait_for_function("Alpine.store('rightCanvas').shouldRender()")
        assert await dock.is_enabled()
        await page.evaluate("async()=>{await closeModal('/plugins/office/webui/standalone-test.html');const s=await import('/js/surfaces.js');await s.open('office')}")
        dock=page.get_by_role('button',name='Attach to sidebar',exact=True)
        await dock.wait_for()
        await dock.click()
        await page.wait_for_function("Alpine.store('rightCanvas').isOpen&&!Alpine.store('rightCanvas').isDetached")
        await page.wait_for_timeout(2000)
        assert (await panel.bounding_box())['height']>400
        await page.screenshot(path=str(output/'docked.png'))
        print('PASS Save As preserves original, usable Open dialog wide/narrow, real dock button',flush=True)
        assert not errors,errors
        await browser.close()
    standalone.unlink(missing_ok=True)


if __name__=='__main__': asyncio.run(main())
