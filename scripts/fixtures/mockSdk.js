// Browser test boundary only. Production bundles never import this SDK fixture.
const metadataKey='com.codex.body-hp/data';
const base={enabled:true,body:{Torso:{current:3,max:3,armor:6},Head:{current:1,max:1,armor:2},'L.Arm':{current:2,max:2,armor:1},'R.Arm':{current:2,max:2,armor:1}},odyssey:{owner:{playerId:'owner',playerName:'Alex'},skills:{Melee:4,Parry:3,Shoot:9,Medicine:6},skillCategories:{Shoot:'combat',Medicine:'applied'},skillLabels:{Shoot:'Precision Shot'},skillPerks:{Shoot:['Accurate']},attributes:{Strength:15,Agility:12,Magic:11},weapons:{melee:[{name:'Blade',damage:5,accuracy:2}],ranged:[{id:'rifle',name:'Rifle',capacity:6,loaded:4,loadedAmmoId:'ball',ammoIds:['ball']}]},ammunition:[{id:'ball',name:'Ball',damage:12,penetration:5,quantity:20}]}};
const token=(id,name,x)=>({id,type:'IMAGE',layer:'CHARACTER',name,visible:true,position:{x,y:100},scale:{x:1,y:1},rotation:0,image:{url:'',width:100,height:100,mime:'image/png'},grid:{dpi:100,offset:{x:0,y:0}},metadata:{[metadataKey]:structuredClone(base)}});
const state={role:'GM',id:'owner',name:'Alex',color:'#54aaff',selection:['hero'],items:[token('hero','Odyssey Ranger',100),token('enemy','Training Target',300)],local:[],room:{},playerCallbacks:[],itemCallbacks:[],localCallbacks:[],roomCallbacks:[],popovers:[],accountScenes:[],writes:0};
state.items[1].metadata[metadataKey].odyssey.owner={playerId:'other',playerName:'Other'};
const emit=(callbacks,value)=>callbacks.forEach(fn=>fn(structuredClone(value)));
const subscribe=list=>fn=>{list.push(fn);return()=>list.splice(list.indexOf(fn),1);};
const noop=async()=>{};
const player=()=>({role:state.role,id:state.id,name:state.name,color:state.color,selection:state.selection});
function collection(key,callbacks){return {getItems:async()=>structuredClone(state[key]),onChange:subscribe(callbacks),getItemBounds:async()=>({width:100,height:100,center:{x:100,y:100},min:{x:50,y:50},max:{x:150,y:150}}),updateItems:async(ids,updater)=>{const items=state[key].filter(i=>ids.includes(i.id));updater(items);state.writes++;queueMicrotask(()=>emit(callbacks,state[key]));},addItems:async items=>{state[key].push(...structuredClone(items));queueMicrotask(()=>emit(callbacks,state[key]));},deleteItems:async ids=>{state[key]=state[key].filter(i=>!ids.includes(i.id));queueMicrotask(()=>emit(callbacks,state[key]));}};}
const OBR={isReady:true,isAvailable:true,onReady:fn=>{setTimeout(fn,0);return()=>{};},
 player:{getRole:async()=>state.role,getId:async()=>state.id,getName:async()=>state.name,getColor:async()=>state.color,getSelection:async()=>state.selection,onChange:subscribe(state.playerCallbacks),select:async ids=>{state.selection=ids;emit(state.playerCallbacks,player());}},
 party:{getPlayers:async()=>[{id:'owner',name:'Alex'},{id:'other',name:'Other'}],onChange:()=>()=>{}},
 room:{id:'room',getMetadata:async()=>state.room,setMetadata:async patch=>{Object.assign(state.room,patch);emit(state.roomCallbacks,state.room);},onMetadataChange:subscribe(state.roomCallbacks)},
 scene:{id:'scene',isReady:async()=>true,onReadyChange:()=>()=>{},items:collection('items',state.itemCallbacks),local:collection('local',state.localCallbacks),grid:{getDpi:async()=>100,getScale:async()=>({parsed:{multiplier:1,unit:'m'}})}},
 action:{setWidth:noop,setHeight:noop,setBadgeText:noop},contextMenu:{remove:noop,create:noop},
 broadcast:{sendMessage:noop,onMessage:()=>()=>{}},tool:{remove:noop,removeMode:noop,create:noop,createMode:async mode=>{state.targetMode=mode;},getActiveTool:async()=>'move',getActiveToolMode:async()=>'default',activateTool:noop,activateMode:noop},
 popover:{close:noop,open:async v=>state.popovers.push(v)},
 assets:{uploadScenes:async scenes=>{state.accountScenes=structuredClone(scenes);},downloadScenes:async()=>structuredClone(state.accountScenes)},
};
function builder(type){const data={id:crypto.randomUUID(),type,metadata:{},position:{x:0,y:0},layer:'DRAWING'};const proxy=new Proxy({},{get(_,key){if(key==='build')return()=>structuredClone(data);return value=>{data[key]=value;return proxy;};}});return proxy;}
export const buildText=()=>builder('TEXT');
export const buildPath=()=>builder('PATH');
export const buildImage=()=>builder('IMAGE');
export function buildSceneUpload(){const data={name:'',items:[]};const proxy=new Proxy({},{get(_,key){if(key==='build')return()=>structuredClone(data);return value=>{data[key]=value;return proxy;};}});return proxy;}
export const Command={MOVE:'MOVE',LINE:'LINE',CLOSE:'CLOSE',QUAD:'QUAD',CUBIC:'CUBIC'};
export const isImage=item=>item?.type==='IMAGE';
window.odysseyTest={state,setRole(role,id='owner'){state.role=role;state.id=id;emit(state.playerCallbacks,player());},select(id){state.selection=[id];emit(state.playerCallbacks,player());},refresh(){emit(state.itemCallbacks,state.items);},data(id='hero'){return state.items.find(i=>i.id===id).metadata[metadataKey];}};
export default OBR;
