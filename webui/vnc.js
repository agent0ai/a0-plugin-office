import RFB from '/plugins/_browser/webui/vendor/novnc/core/rfb.js';
const screen = document.getElementById('screen');
const status = document.getElementById('status');
let rfb, timer, stopped = false;
const viewer = window.a0OfficeViewer = {connected:false, paste(text) {
  if (!rfb || !viewer.connected) return;
  rfb.clipboardPasteFrom(text);
  rfb.sendKey(0xffe3,'ControlLeft',true);
  rfb.sendKey(0x76,'KeyV');
  rfb.sendKey(0xffe3,'ControlLeft',false);
}};
async function connect() {
  if (stopped) return;
  try {
    const response = await fetch('./credentials',{credentials:'same-origin',cache:'no-store'});
    if (!response.ok) throw new Error('Office session is unavailable');
    const credentials = await response.json();
    const url = new URL('./ws', location.href);
    url.protocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
    const client = rfb = new RFB(screen,url.href,{credentials});
    viewer.rfb = client;
    client.scaleViewport = true;
    client.resizeSession = false;
    client.addEventListener('connect',()=>{viewer.connected=true;status.hidden=true;client.focus();});
    client.addEventListener('disconnect',()=>{
      viewer.connected=false;status.hidden=false;status.textContent='Office disconnected. Reconnecting…';
      if (!stopped) timer=setTimeout(connect,1500);
    });
    client.addEventListener('clipboard',({detail})=>{
      if (document.hasFocus()) void navigator.clipboard?.writeText(detail.text).catch(()=>{});
    });
  } catch (error) {
    status.hidden=false;status.textContent=error.message;
    if (!stopped) timer=setTimeout(connect,1500);
  }
}
document.addEventListener('keydown',async event=>{
  if (!viewer.connected || !((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='v')) return;
  event.preventDefault();event.stopImmediatePropagation();
  try {viewer.paste(await navigator.clipboard.readText());}
  catch {status.hidden=false;status.textContent='Allow clipboard access to paste into Office';}
},true);
window.addEventListener('pagehide',()=>{stopped=true;clearTimeout(timer);rfb?.disconnect();});
void connect();
