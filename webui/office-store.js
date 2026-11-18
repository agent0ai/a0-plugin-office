import {createStore} from '/js/AlpineStore.js';
import {callJsonApi} from '/js/api.js';
import {showConfirmDialog} from '/js/confirmDialog.js';
import {toastFrontendError, toastFrontendInfo} from '/components/notifications/notification-store.js';
import {store as fileBrowser} from '/components/modals/file-browser/file-browser-store.js';
let frame, host, poll, resize, observer, polling=false;
const api = async payload => {
  const result=await callJsonApi('plugins/office/office_session',{context_id:globalThis.getContext?.()||'',...payload});
  if(result?.ok===false) throw new Error(result.error||'Office request failed');
  return result;
};
function park(){let node=document.getElementById('office-keepalive');if(!node){node=document.createElement('div');node.id='office-keepalive';node.hidden=true;document.body.append(node);}return node;}
function move(target){if(!frame||!target||frame.parentElement===target)return;if(target.moveBefore&&frame.isConnected)target.moveBefore(frame,null);else target.append(frame);}
export const store=createStore('office',{
  lastError:'',busy:false,startup:{state:'idle'},documents:[],active:'',viewerUrl:'',title:'Untitled',picker:false,entries:[],browsePath:'',parentPath:'',versions:[],showHistory:false,
  get exportFormats(){const ext=this.documents.find(doc=>doc.file_id===this.active)?.extension;return ['pdf',...(['odt','docx'].includes(ext)?['odt','docx']:['ods','xlsx'].includes(ext)?['ods','xlsx']:['odp','pptx']).filter(fmt=>fmt!==ext)];},
  async mount(panel,options={}){
    if(!panel.getBoundingClientRect().width)return;
    host=panel.querySelector('[data-office-host]');move(host);
    observer?.disconnect();observer=new ResizeObserver(()=>{clearTimeout(resize);resize=setTimeout(()=>this.resize(),250);});observer.observe(host);
    clearInterval(poll);poll=setInterval(()=>this.poll(),1500);await this.poll();
  },
  detach(panel){if(host&&panel.contains(host)){move(park());host=null;observer?.disconnect();clearInterval(poll);}},
  async poll(){if(polling)return;polling=true;try{
    const result=await api({action:this.busy?'status':'sync'});this.startup=result;this.documents=result.documents||this.documents;if(!this.documents.some(doc=>doc.file_id===this.active))this.active=this.documents[0]?.file_id||'';
  }catch(error){if(!this.busy)this.error(error);}finally{polling=false;}},
  error(error){const message=error.message||String(error);if(message!==this.lastError)toastFrontendError(message,'Office');this.lastError=message;},
  async run(payload){if(this.busy)return null;this.busy=true;try{
    const result=await api(payload);this.lastError='';
    if(result.document){this.active=result.document.file_id;this.title=result.document.basename.replace(/\.[^.]+$/,'');}
    if(result.viewer_url){this.viewerUrl=result.viewer_url;if(!frame){frame=document.createElement('iframe');frame.title='Interactive LibreOffice';frame.allow='clipboard-read; clipboard-write';}if(frame.getAttribute('src')!==result.viewer_url)frame.src=result.viewer_url;move(host||park());}
    if(result.path&&payload.action==='export')toastFrontendInfo('Exported to '+result.path,'Office');
    return result;
  }catch(error){this.error(error);return null;}finally{this.busy=false;await this.poll();void this.resize();}},
  async create(format){return this.run({action:'create',format,title:this.title||'Untitled',kind:format==='ods'?'spreadsheet':format==='odp'?'presentation':'document'});},
  async open(payload={}){this.picker=false;return this.run({action:'open',...payload});},
  async command(action,extra={}){return this.run({action,file_id:this.active,...extra});},
  async discard(){if(await showConfirmDialog({title:'Discard pending Office edits?',message:'Close this document without writing its pending edits to the original file.',confirmText:'Discard and close'}))await this.close(this.active,false);},
  async close(file_id,save=true){if(!await this.run({action:'close',file_id,save}))return;if(this.active===file_id)this.active=this.documents[0]?.file_id||'';if(this.active)await this.open({file_id:this.active});},
  async resize(){if(!host||!this.viewerUrl)return;const box=host.getBoundingClientRect();if(box.width<100||box.height<100)return;try{await api({action:'resize',width:Math.round(box.width),height:Math.round(box.height)});}catch{}},
  async browse(path=''){try{const result=await api({action:'browse',path});this.browsePath=result.path;this.parentPath=result.parent;this.entries=result.entries;this.picker=true;}catch(error){this.error(error);}},
  async saveAs(){const doc=this.documents.find(item=>item.file_id===this.active);if(!doc)return;await fileBrowser.openSaveAsPicker(doc.path.slice(0,doc.path.lastIndexOf('/')),{filename:doc.basename,onConfirm:async({path})=>Boolean(await this.command('save_as',{target_path:path}))});},
  async rename(){const doc=this.documents.find(item=>item.file_id===this.active);if(!doc)return;await fileBrowser.openSaveAsPicker(doc.path.slice(0,doc.path.lastIndexOf('/')),{filename:doc.basename,onConfirm:async({path})=>Boolean(await this.command('rename',{target_path:path}))});},
  async history(){const result=await this.command('versions');if(result){this.versions=result.versions;this.showHistory=true;}},
  async restore(version_id){if(await this.command('restore',{version_id}))this.showHistory=false;},
});
