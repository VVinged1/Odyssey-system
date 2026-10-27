import {armorPoolTier,poolCategoryPaths,groupPoolEntries} from '../poolCategories.js';
import {distanceLabel,createTargetDistance} from '../targetDistance.js';
import assert from 'node:assert/strict';
import {captureBlueprint,blueprintId,copyBlueprint,poolEntries,exportPool,parsePoolJSON,ITEM_POOL_PREFIX} from '../itemPool.js';
import {saveInventoryEntry,normalizeInventory,equipItem,unequipItem,setMagazineRounds,swapMagazine} from '../inventory.js';
import {setCombinedPartHp} from '../characterRules.js';
const character=()=>({body:{Torso:{current:3,max:3,armor:0,slot:'torso'}},odyssey:{skills:{sourceMedicine:2},skillLabels:{sourceMedicine:'Medicine'},attributes:{},inventory:{version:4,items:[],magazines:[],baseArmor:{},baseMax:{},armorHp:{}},ammunition:[{id:'ball',name:'Ball',damage:12,penetration:5,quantity:50}],weapons:{melee:[],ranged:[]}}});
let count=0;const test=async(name,run)=>{await run();count++;console.log('PASS',name);};
await test('One equipped armor separates from spare quantity, keeping independent damage',()=>{
 const t=character();saveInventoryEntry(t,null,{kind:'armor',name:'Vest',quantity:3,armor:12,criticalMax:2,slots:['Torso']},'vest');equipItem(t,'vest','Torso');
 assert.equal(t.odyssey.inventory.items.find(i=>i.id==='vest').quantity,1);const spare=t.odyssey.inventory.items.find(i=>i.id!=='vest');assert.equal(spare.quantity,2);assert.equal(spare.equippedPart,'');assert.equal(t.body.Torso.current,5);assert.equal(t.body.Torso.max,5);
 setCombinedPartHp(t,'Torso',3);assert.equal(t.odyssey.inventory.items.find(i=>i.id==='vest').criticalDamage,2);assert.equal(spare.criticalDamage,0);unequipItem(t,'vest');equipItem(t,spare.id,'Torso');assert.equal(t.body.Torso.current,5);assert.equal(t.odyssey.inventory.items.reduce((sum,i)=>sum+i.quantity,0),3);
});
await test('Previously equipped stacks migrate deterministically without multiplying HP',()=>{
 const raw={version:4,items:[{id:'vest',name:'Vest',kind:'armor',quantity:3,armor:12,criticalMax:2,equippedPart:'Torso'}],magazines:[]};
 const a=normalizeInventory(raw),b=normalizeInventory(a);assert.equal(a.items.length,2);assert.equal(a.items[0].quantity,1);assert.equal(a.items[1].quantity,2);assert.equal(a.items[1].equippedPart,'');assert.deepEqual(a,b);
});
await test('Three magazines are three independent physical instances',()=>{
 const t=character();saveInventoryEntry(t,null,{kind:'magazine',name:'Magazine',quantity:3,capacity:6,loaded:2,loadedAmmoId:'ball'},'mag');assert.equal(t.odyssey.inventory.magazines.length,3);assert.equal(new Set(t.odyssey.inventory.magazines.map(m=>m.id)).size,3);
 saveInventoryEntry(t,null,{kind:'weapon',weaponKind:'ranged',name:'Rifle',acceptsMagazines:true,acceptedMagazineIds:['mag']},'rifle');swapMagazine(t.odyssey,0,'mag');setMagazineRounds(t.odyssey,'mag',1,'ball');assert.equal(t.odyssey.inventory.magazines.filter(m=>m.loaded===2).length,2);swapMagazine(t.odyssey,0,'mag-copy-1');assert.equal(t.odyssey.weapons.ranged[0].loaded,2);
});
await test('Pool identity ignores token IDs, wear, damage and quantity',async()=>{
 const t=character();saveInventoryEntry(t,null,{kind:'armor',name:'Vest',quantity:3,armor:12,criticalMax:2,slots:['Torso']},'source-private-id');const before=captureBlueprint(t,'item','source-private-id');equipItem(t,'source-private-id','Torso');setCombinedPartHp(t,'Torso',4);const after=captureBlueprint(t,'item','source-private-id');
 assert.equal(await blueprintId(before),await blueprintId(after));assert.ok(!JSON.stringify(after).includes('source-private-id'));assert.equal(after.values.criticalDamage,0);assert.equal(after.values.equippedPart,undefined);
});
await test('Pool weapon copies remap portable dependencies and never clone loaded rounds',()=>{
 const source=character();saveInventoryEntry(source,null,{kind:'magazine',name:'Combat mag',capacity:12,loaded:8,loadedAmmoId:'ball'},'secret-mag');saveInventoryEntry(source,null,{kind:'weapon',weaponKind:'ranged',name:'Carbine',weaponCategory:'assault',rangeReduction:{clinch:12,long:3},acceptsMagazines:true,acceptedMagazineIds:['secret-mag']},'secret-gun');
 const blueprint=captureBlueprint(source,'weapon-ranged','secret-gun');const target=character();target.odyssey.ammunition=[];copyBlueprint(target,blueprint,1,'copy');assert.equal(target.odyssey.weapons.ranged[0].loaded,0);assert.equal(target.odyssey.weapons.ranged[0].weaponCategory,'assault');assert.deepEqual(target.odyssey.weapons.ranged[0].rangeReduction,{clinch:12,long:3});assert.equal(target.odyssey.inventory.magazines.length,1);assert.equal(target.odyssey.inventory.magazines[0].loaded,0);assert.ok(!JSON.stringify(blueprint).includes('secret-mag'));
 assert.equal(target.odyssey.weapons.ranged[0].acceptedMagazineIds[0],target.odyssey.inventory.magazines[0].id);
 saveInventoryEntry(source,null,{kind:'weapon',weaponKind:'ranged',name:'Internal',acceptsMagazines:false,capacity:5,ammoIds:['ball'],loaded:3,loadedAmmoId:'ball'},'internal');copyBlueprint(target,captureBlueprint(source,'weapon-ranged','internal'),1,'builtin');assert.equal(target.odyssey.ammunition[0].quantity,0);assert.equal(target.odyssey.weapons.ranged[1].loaded,0);assert.equal(target.odyssey.inventory.magazines.length,1);
});
await test('Pool bonuses bind to destination skills; missing skills remain portable',()=>{
 const source=character();saveInventoryEntry(source,null,{kind:'resource',name:'Manual',quantity:1,bonuses:{skills:{sourceMedicine:2,Piloting:3}}},'manual');const blueprint=captureBlueprint(source,'item','manual');const target=character();target.odyssey.skills={destinationMedicine:4};target.odyssey.skillLabels={destinationMedicine:'Medicine'};copyBlueprint(target,blueprint,2,'copy');assert.deepEqual(target.odyssey.inventory.items[0].bonuses.skills,{destinationMedicine:2,Piloting:3});assert.equal(target.odyssey.inventory.items[0].quantity,2);assert.equal(target.odyssey.skills.destinationMedicine,4);
});
await test('Room catalog survives missing scene items; JSON and tombstones preserve definitions',async()=>{
 const source=character();saveInventoryEntry(source,null,{kind:'magazine',name:'Mag',capacity:5,loaded:0},'mag');const blueprint=captureBlueprint(source,'magazine','mag'),id=await blueprintId(blueprint),metadata={[ITEM_POOL_PREFIX+id]:{blueprint}};
 const entries=poolEntries(metadata).entries;assert.equal(entries.length,1);const imported=parsePoolJSON(JSON.parse(JSON.stringify(exportPool(entries))));const target=character();copyBlueprint(target,imported[0],3,'copies');assert.equal(target.odyssey.inventory.magazines.length,3);assert.ok(target.odyssey.inventory.magazines.every(m=>m.loaded===0));
 metadata[ITEM_POOL_PREFIX+id]={deleted:true};assert.equal(poolEntries(metadata).entries.length,0);assert.throws(()=>parsePoolJSON({format:'other',version:1,entries:[]}),/Choose/);
});

await test('Armor pool categories use per-part boundaries and keep unique parent counts',()=>{
 for(const [value,expected]of [[20,'Light Armor'],[21,'Medium Armor'],[40,'Medium Armor'],[41,'Heavy Armor'],[60,'Heavy Armor'],[61,'Superheavy Armor'],[100,'Superheavy Armor'],[101,'Outside listed ranges']])assert.equal(armorPoolTier(value,true),expected);
 for(const [value,expected]of [[10,'Light Armor'],[11,'Medium Armor'],[20,'Medium Armor'],[21,'Heavy Armor'],[40,'Heavy Armor'],[41,'Superheavy Armor'],[60,'Superheavy Armor'],[61,'Outside listed ranges']])assert.equal(armorPoolTier(value,false),expected);
 const blueprint={values:{kind:'armor',armor:20,slots:['Torso','L.Arm']}};assert.deepEqual(poolCategoryPaths(blueprint),[['Armor','Light Armor','Torso'],['Armor','Medium Armor','Left Arm']]);assert.equal(groupPoolEntries([{id:'one',blueprint}])[0].count,1);
 assert.deepEqual(poolCategoryPaths({values:{kind:'implant',slots:['Head','R.Arm']}}),[['Implants','Head'],['Implants','Right Arm']]);
});
await test('Distance uses SDK grid output and scene precision, including local ruler cleanup',async()=>{
 const items=[],calls=[],reads=[];const points={source:{x:100,y:100},target:{x:400,y:300}};
 const OBR={scene:{items:{getItemBounds:async ids=>({center:points[ids[0]]})},grid:{getScale:async()=>({parsed:{multiplier:5,unit:'ft',digits:1}}),getDistance:async(from,to)=>{calls.push([from,to]);return 2.5;}},local:{getItems:async()=>items,addItems:async rows=>items.push(...rows),updateItems:async(ids,fn)=>fn(items.filter(i=>ids.includes(i.id))),deleteItems:async ids=>{for(let i=items.length-1;i>=0;i--)if(ids.includes(items[i].id))items.splice(i,1);}}}};
 const buildRuler=()=>{const data={type:'RULER'},proxy=new Proxy({},{get:(_,key)=>key==='build'?()=>({...data}):value=>{data[key]=value;return proxy;}});return proxy;};
 const ruler=createTargetDistance({OBR,buildRuler,onRead:value=>reads.push(value)}),source={id:'source',position:points.source},target={id:'target',position:points.target};
 const reading=await ruler.measure(source,target);assert.equal(reading.label,'2.5 cells · 12.5ft');assert.deepEqual(calls[0],[points.source,points.target]);await ruler.measure(source,target);assert.equal(calls.length,1);
 ruler.preview(source,target);await new Promise(resolve=>setTimeout(resolve,120));assert.equal(items.length,1);assert.equal(items[0].measurement,reading.label);assert.equal(items[0].disableHit,true);
 await ruler.clear();assert.equal(items.length,0);ruler.invalidate();await ruler.measure(source,target);assert.equal(calls.length,2);assert.equal(distanceLabel(2,{parsed:{multiplier:5,unit:'m',digits:2}}),'2 cells · 10.00m');
});
await test('Cancelling an in-flight ruler calculation cannot leave a local artifact',async()=>{
 const items=[];let release;const delayed=new Promise(resolve=>release=resolve);
 const OBR={scene:{items:{getItemBounds:async()=>({center:{x:0,y:0}})},grid:{getScale:async()=>({parsed:{multiplier:1,unit:'m'}}),getDistance:async()=>delayed},local:{getItems:async()=>[],addItems:async rows=>items.push(...rows),deleteItems:async()=>{items.length=0;}}}};
 const ruler=createTargetDistance({OBR,buildRuler:()=>{throw new Error('Cancelled ruler must not be built.');},onRead:()=>{}});ruler.preview({id:'source'},{x:100,y:100});await new Promise(resolve=>setTimeout(resolve,90));await ruler.clear();release(1);await new Promise(resolve=>setTimeout(resolve,0));assert.equal(items.length,0);
});
console.log(count+' item pool, category, instance and distance checks passed.');
