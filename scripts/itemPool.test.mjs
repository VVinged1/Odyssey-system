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
 const source=character();saveInventoryEntry(source,null,{kind:'magazine',name:'Combat mag',capacity:12,loaded:8,loadedAmmoId:'ball'},'secret-mag');saveInventoryEntry(source,null,{kind:'weapon',weaponKind:'ranged',name:'Carbine',acceptsMagazines:true,acceptedMagazineIds:['secret-mag']},'secret-gun');
 const blueprint=captureBlueprint(source,'weapon-ranged','secret-gun');const target=character();target.odyssey.ammunition=[];copyBlueprint(target,blueprint,1,'copy');assert.equal(target.odyssey.weapons.ranged[0].loaded,0);assert.equal(target.odyssey.inventory.magazines.length,1);assert.equal(target.odyssey.inventory.magazines[0].loaded,0);assert.ok(!JSON.stringify(blueprint).includes('secret-mag'));
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
console.log(count+' item pool and instance checks passed.');
