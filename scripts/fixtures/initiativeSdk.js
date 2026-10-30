import {normalizeIconPaths} from '../../node_modules/@owlbear-rodeo/sdk/lib/common/normalize.js';
import {produceWithPatches,enablePatches,setAutoFreeze} from 'immer';
enablePatches();setAutoFreeze(false);
// Browser test boundary only. Production bundles never import this SDK fixture.
const metadataKey='com.codex.body-hp/data';
const base={enabled:true,body:{Torso:{current:3,max:3,armor:6},Head:{current:1,max:1,armor:2},'L.Arm':{current:2,max:2,armor:1},'R.Arm':{current:2,max:2,armor:1}},odyssey:{owner:{playerId:'owner',playerName:'Alex'},skills:{Melee:4,Parry:3,Shoot:4,Medicine:3},skillCategories:{Shoot:'combat',Medicine:'applied'},skillLabels:{Shoot:'Precision Shot'},skillPerks:{Shoot:['Accurate']},attributes:{Strength:15,Agility:12,Magic:11},weapons:{melee:[{name:'Blade',damage:5,accuracy:2}],ranged:[{id:'rifle',name:'Rifle',capacity:6,loaded:4,loadedAmmoId:'ball',ammoIds:['ball']}]},ammunition:[{id:'ball',name:'Ball',damage:12,penetration:5,quantity:20}]}};
if(globalThis.odysseyInitialInventory)base.odyssey.inventory=structuredClone(globalThis.odysseyInitialInventory);
const token=(id,name,x)=>({id,type:'IMAGE',layer:'CHARACTER',name,visible:true,position:{x,y:100},scale:{x:1,y:1},rotation:0,image:{url:'',width:100,height:100,mime:'image/png'},grid:{dpi:100,offset:{x:0,y:0}},metadata:{[metadataKey]:structuredClone(base)}});
const state={role:'GM',id:'owner',name:'Alex',color:'#54aaff',selection:['hero'],items:[token('hero','Odyssey Ranger',100),token('enemy','Training Target',300)],local:[],room:structuredClone(globalThis.odysseyInitialRoomMetadata||{}),sceneMetadata:structuredClone(globalThis.odysseyInitialSceneMetadata||{}),sceneMetadataCallbacks:[],sceneReadyCallbacks:[],playerCallbacks:[],itemCallbacks:[],localCallbacks:[],roomCallbacks:[],popovers:[],accountScenes:[],writes:0};
state.items[1].metadata[metadataKey].odyssey.owner={playerId:'other',playerName:'Other'};
const emit=(callbacks,value)=>callbacks.forEach(fn=>fn(structuredClone(value)));
const subscribe=list=>fn=>{list.push(fn);return()=>list.splice(list.indexOf(fn),1);};
state.notifications=[];state.contextMenus={};state.broadcasts=[];state.broadcastCallbacks=[];
const noop=async()=>{};
const player=()=>({role:state.role,id:state.id,name:state.name,color:state.color,selection:state.selection});
function collection(key,callbacks){return {getItems:async()=>structuredClone(state[key]),onChange:subscribe(callbacks),getItemBounds:async ids=>{const point=state[key].find(i=>ids.includes(i.id))?.position||{x:100,y:100};return {width:100,height:100,center:point,min:{x:point.x-50,y:point.y-50},max:{x:point.x+50,y:point.y+50}};},updateItems:async(ids,updater)=>{const previous=state[key].filter(i=>ids.includes(i.id)),[items]=produceWithPatches(previous,updater);const updates=new Map(items.map(i=>[i.id,i]));state[key]=state[key].map(i=>updates.get(i.id)||i);state.writes++;if(key==='items'&&state.afterItemWrite)await state.afterItemWrite(structuredClone(items));queueMicrotask(()=>emit(callbacks,state[key]));},addItems:async items=>{state[key].push(...structuredClone(items));queueMicrotask(()=>emit(callbacks,state[key]));},deleteItems:async ids=>{state[key]=state[key].filter(i=>!ids.includes(i.id));queueMicrotask(()=>emit(callbacks,state[key]));}};}
const OBR={notification:{show:async(message,variant)=>{state.notifications.push({message,variant});return String(state.notifications.length);},close:noop},isReady:true,isAvailable:true,onReady:fn=>{setTimeout(fn,0);return()=>{};},
 player:{getRole:async()=>state.role,getId:async()=>state.id,getName:async()=>state.name,getColor:async()=>state.color,getSelection:async()=>state.selection,onChange:subscribe(state.playerCallbacks),select:async ids=>{state.selection=ids;emit(state.playerCallbacks,player());}},
 party:{getPlayers:async()=>[{id:'owner',name:'Alex'},{id:'other',name:'Other'}],onChange:()=>()=>{}},
 room:{id:'room',getMetadata:async()=>state.room,setMetadata:async patch=>{const next={...state.room,...patch};if(new TextEncoder().encode(JSON.stringify(next)).length>=16384)throw Error('Room metadata exceeds 16 KB.');Object.assign(state.room,patch);emit(state.roomCallbacks,state.room);},onMetadataChange:subscribe(state.roomCallbacks)},
 scene:{id:'scene',isReady:async()=>true,onReadyChange:subscribe(state.sceneReadyCallbacks),getMetadata:async()=>structuredClone(state.sceneMetadata),setMetadata:async patch=>{if(state.sceneWriteError)throw Error(state.sceneWriteError);Object.assign(state.sceneMetadata,patch);emit(state.sceneMetadataCallbacks,state.sceneMetadata);},onMetadataChange:subscribe(state.sceneMetadataCallbacks),items:collection('items',state.itemCallbacks),local:collection('local',state.localCallbacks),grid:{getDpi:async()=>100,getScale:async()=>({parsed:{multiplier:1,unit:'m',digits:0}}),getDistance:async(a,b)=>Math.max(Math.abs(a.x-b.x),Math.abs(a.y-b.y))/100,onChange:()=>()=>{}}},
 viewport:{getWidth:async()=>1000,getHeight:async()=>800,getScale:async()=>1,transformPoint:async p=>p,inverseTransformPoint:async p=>p,animateTo:noop},
 action:{setWidth:async n=>{state.actionWidth=n;},setHeight:async n=>{state.actionHeight=n;},setBadgeText:noop},contextMenu:{remove:async id=>{delete state.contextMenus[id];},create:async menu=>{const icons=normalizeIconPaths(menu.icons);for(const icon of icons){if(!['http:','https:'].includes(new URL(icon.icon).protocol))throw Error('Invalid SDK menu icon URL.');}state.contextMenus[menu.id]={...menu,icons};}},
 broadcast:{sendMessage:async(channel,data)=>{state.broadcasts.push({channel,data:structuredClone(data)});},onMessage:(channel,fn)=>{const row={channel,fn};state.broadcastCallbacks.push(row);return()=>state.broadcastCallbacks.splice(state.broadcastCallbacks.indexOf(row),1);}},tool:{remove:noop,removeMode:noop,create:noop,createMode:async mode=>{state.targetMode=mode;},getActiveTool:async()=>'move',getActiveToolMode:async()=>'default',activateTool:noop,activateMode:noop},
 popover:{close:noop,open:async v=>state.popovers.push(v)},
 assets:{uploadScenes:async scenes=>{state.accountScenes=structuredClone(scenes);},downloadScenes:async()=>structuredClone(state.accountScenes)},
};
function builder(type){const data={id:crypto.randomUUID(),type,metadata:{},position:{x:0,y:0},layer:'DRAWING'};const proxy=new Proxy({},{get(_,key){if(key==='build')return()=>structuredClone(data);return value=>{data[key]=value;return proxy;};}});return proxy;}
export const buildText=()=>builder('TEXT');
export const buildPath=()=>builder('PATH');
export const buildRuler=()=>builder('RULER');
export const buildImage=()=>builder('IMAGE');
export function buildSceneUpload(){const data={name:'',items:[]};const proxy=new Proxy({},{get(_,key){if(key==='build')return()=>structuredClone(data);return value=>{data[key]=value;return proxy;};}});return proxy;}
export const Command={MOVE:'MOVE',LINE:'LINE',CLOSE:'CLOSE',QUAD:'QUAD',CUBIC:'CUBIC'};
export const Math2={subtract:(a,b)=>({x:a.x-b.x,y:a.y-b.y}),multiply:(a,n)=>({x:a.x*n,y:a.y*n})};
export const isImage=item=>item?.type==='IMAGE';
window.odysseyTest={state,receiveBroadcast(channel,data){for(const row of state.broadcastCallbacks)if(row.channel===channel)row.fn({data:structuredClone(data)});},setRole(role,id='owner'){state.role=role;state.id=id;emit(state.playerCallbacks,player());},select(id){state.selection=[id];emit(state.playerCallbacks,player());},refresh(){emit(state.itemCallbacks,state.items);},setRoom(patch){Object.assign(state.room,patch);emit(state.roomCallbacks,state.room);},data(id='hero'){return state.items.find(i=>i.id===id).metadata[metadataKey];}};
export default OBR;
