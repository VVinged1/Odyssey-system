import {ITEM_POOL_PREFIX,validateBlueprint,blueprintText} from './itemPool.js';
export const ITEM_POOL_CACHE_PREFIX='odyssey.item-pool.v2:';
const timestamp=row=>Number.isFinite(row?.updatedAt)?row.updatedAt:0;
export function poolRecordWins(incoming,current) {
 if(!current)return true;const delta=timestamp(incoming)-timestamp(current);if(delta)return delta>0;
 return incoming.deleted===true&&current.deleted!==true;
}
export function mergePoolRecords(current,...sources) {
 const result={...current};for(const source of sources)for(const [key,row]of Object.entries(source||{})) {
  if(!key.startsWith(ITEM_POOL_PREFIX)||!row||typeof row!=='object')continue;
  if(row.deleted!==true){try{validateBlueprint(row.blueprint);}catch{continue;}}
  if(poolRecordWins(row,result[key]))result[key]=structuredClone(row);
 }return result;
}
// Room metadata has a 16 KB total limit. The catalog lives in scene metadata,
// backed up per room in this browser so it can follow the GM to another scene.
export function createCampaignPoolStore({room,scene,storage,roomId,requireGM,canSync,onChange}) {
 const cacheKey=ITEM_POOL_CACHE_PREFIX+roomId,warnings=new Map();let records={},clock=0,lastNotice='',readRequest=0,sceneRevision=0;
 const warn=(key,message)=>message?warnings.set(key,message):warnings.delete(key);
 let backupReadable=true;
 try{const raw=storage.getItem(cacheKey);if(raw){const cache=JSON.parse(raw);if(cache?.version!==2||!cache.records||typeof cache.records!=='object')throw Error('Invalid item pool backup.');records=mergePoolRecords({},cache.records);}}
 catch{backupReadable=false;warn('backup','Browser item-pool backup could not be read; the saved copy has been kept.');}
 const backup=()=>{if(!backupReadable)return;try{const json=JSON.stringify({version:2,records});if(new TextEncoder().encode(json).length>4*1024*1024)throw Error('Item pool backup exceeds 4 MB.');storage.setItem(cacheKey,json);warn('backup','');}catch(e){warn('backup','Browser backup failed: '+e.message+' Export JSON to keep a copy.');}};
 const notify=()=>{const warning=[...warnings.values()].join(' '),signature=blueprintText(records)+'|'+warning;if(signature!==lastNotice){lastNotice=signature;onChange(structuredClone(records),warning);}};
 const merge=(metadata)=>{records=mergePoolRecords(records,metadata);for(const row of Object.values(records))clock=Math.max(clock,timestamp(row));backup();notify();};
 room.onMetadataChange(metadata=>merge(metadata));
 scene.onMetadataChange(metadata=>{sceneRevision++;merge(metadata);});
 let syncQueue=Promise.resolve();
 const persist=()=>{
  const task=syncQueue.catch(()=>{}).then(async()=>{
   await requireGM();const metadata=await scene.getMetadata(),patch={};merge(metadata);
   for(const [key,row]of Object.entries(records))if(poolRecordWins(row,metadata[key]))patch[key]=row;
   if(Object.keys(patch).length)await scene.setMetadata(patch);warn('sharing','');notify();
  });syncQueue=task;return task;
 };
 async function read(){
  const request=++readRequest,before=sceneRevision;
  const results=await Promise.allSettled([room.getMetadata(),scene.getMetadata()]);
  if(request!==readRequest)return structuredClone(records);
  if(results[0].status==='fulfilled'){merge(results[0].value);warn('room','');}else warn('room','Legacy room pool could not be read: '+results[0].reason?.message);
  if(results[1].status==='fulfilled'){if(before===sceneRevision)merge(results[1].value);warn('scene','');}else warn('scene','Scene item pool could not be read: '+results[1].reason?.message);
  if(canSync()){try{await persist();}catch(e){warn('sharing','Pool sharing failed: '+e.message+' Available items are retained in this browser.');}}
  backup();notify();return structuredClone(records);
 }
 scene.onReadyChange(ready=>{if(ready)void read();});
 return {
  read,
  async write(patch){await requireGM();const next={};for(const [key,row]of Object.entries(patch)){if(!key.startsWith(ITEM_POOL_PREFIX))throw Error('Invalid pool metadata key.');if(row.deleted!==true)validateBlueprint(row.blueprint);next[key]={...structuredClone(row),updatedAt:clock=Math.max(Date.now(),clock+1)};}merge(next);try{await persist();}catch(e){warn('sharing','Pool sharing failed: '+e.message+' Available items are retained in this browser.');notify();throw e;}},
  refreshBackup:()=>{backup();notify();},
  metadata:()=>structuredClone(records),
 };
}
