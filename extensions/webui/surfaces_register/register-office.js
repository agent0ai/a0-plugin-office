import {store} from '/plugins/office/webui/office-store.js';
export default function registerOffice(surfaces){surfaces.registerSurface({id:'office',title:'Office',icon:'description',order:35,modalPath:'/plugins/office/webui/main.html',
  async open(payload={}){
    for(let i=0;i<100;i++){const panel=[...document.querySelectorAll('[data-surface-id="office"] .office-panel')].find(panel=>panel.getBoundingClientRect().width>0);if(panel){await store.mount(panel);break;}await new Promise(resolve=>setTimeout(resolve,50));}
    if(payload.path||payload.file_id)await store.open(payload);
  },async close(){const panel=[...document.querySelectorAll('[data-surface-id="office"] .office-panel')].find(panel=>panel.getBoundingClientRect().width>0);if(panel)store.detach(panel);}
});}
