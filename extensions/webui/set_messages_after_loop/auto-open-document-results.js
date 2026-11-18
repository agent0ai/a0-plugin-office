import {open} from '/js/surfaces.js';
import {normalizeDocumentMetadata, parseDocumentResult} from '../lib/document-actions.js';
const seen=new Set();
export default function(context){
  if(context?.historyEmpty)return;
  for(const {args} of context?.results||[]){
    if(seen.has(args?.id))continue;
    const result=parseDocumentResult(args?.content);
    const tool=args?.kvps?._tool_name||args?._tool_name||result?._tool_name;
    if(tool!=='office_artifact')continue;
    const doc=normalizeDocumentMetadata(args,result);
    if(doc.format==='pdf')continue;
    if(doc.action==='open'||doc.open_in_canvas||doc.open_in_desktop){
      seen.add(args.id);
      void open('office',doc);
    }
  }
}
