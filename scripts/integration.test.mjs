import {normalizeEffects,effectiveOdyssey,reconcileEquipment,ARMOR_TRAINING,useItemOnTracker,absorbEquipmentCritical,setCombinedPartHp,setCombinedPartMax} from '../characterRules.js';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { normalizeInventory, swapMagazine, fillMagazine, emptyMagazine, setMagazineRounds, equipItem, unequipItem, consumeItem, inventoryAmmoTotal,saveInventoryEntry,deleteInventoryEntry } from '../inventory.js';
import { sanitizeAmmunition, sanitizeMagazine, reloadMagazine, getAmmoShot, consumeAmmo } from '../ammunition.js';
import { makeTemplate, instantiateTemplate, profileData, normalizeLibrary, emptyLibrary, mergeLibraries, readLibrary, writeLibrary, libraryFromScenes, LIBRARY_ITEM_KEY, LIBRARY_STORAGE_KEY } from '../profiles.js';
import { attributeCards, skillCards, inventoryView } from '../ui/characterView.js';
const source=readFileSync(new URL('../shared.js',import.meta.url),'utf8').replace(/^import .*;\r?$/gm,'').replace(/^export \{ OBR \};\r?$/gm,'').replace(/^export /gm,'');
const ctx=vm.createContext({structuredClone,normalizeEffects,effectiveOdyssey,reconcileEquipment,setCombinedPartHp,setCombinedPartMax,ARMOR_TRAINING,normalizeInventory,sanitizeAmmunition,sanitizeMagazine,OBR:{},console});
vm.runInContext(source+'\nthis.sanitize=sanitizeTrackerData; this.canControl=canPlayerControlToken;',ctx);
const sanitize=ctx.sanitize;
const legacy=()=>({enabled:true,body:{Torso:{current:2,max:3,armor:6,minor:3,serious:1},'Tail & claw':{current:1,max:4,armor:2,slot:'other',attackPenalty:7,hidden:true,label:'Tail'}},identity:{playerId:'discord-id',characterId:'external-id'},history:[{summary:'Past roll'}],odyssey:{owner:{playerId:'owner',playerName:'Owner'},skills:{Shoot:9,Melee:4,Parry:3},skillCategories:{Shoot:'combat'},skillLabels:{Shoot:'Custom Shot'},skillPerks:{Shoot:['Accurate']},skillStrengthBonuses:{Shoot:true},attributes:{Strength:15,Magic:11},weapons:{melee:[],ranged:[{id:'gun',name:'Rifle',capacity:6,loaded:4,loadedAmmoId:'ball',ammoIds:['ball','ap']}]},ammunition:[{id:'ball',name:'Ball',quantity:10,damage:12,penetration:5},{id:'ap',name:'AP',quantity:8,damage:8,penetration:20}]}});
let count=0;function test(name,run){run();console.log('PASS',name);count++;}
test('Legacy migration keeps custom body, skills, identity and ammo; deterministic read',()=>{
 const raw=legacy(), before=JSON.stringify(raw), a=sanitize(raw), b=sanitize(raw);
 assert.equal(JSON.stringify(raw),before);assert.equal(a.odyssey.skills.Shoot,9);assert.equal(a.body['Tail & claw'].hidden,true);assert.equal(a.identity.characterId,'external-id');
 assert.equal(a.odyssey.inventory.magazines.length,1);assert.equal(a.odyssey.inventory.magazines[0].loaded,4);
 assert.equal(a.odyssey.weapons.ranged[0].magazineId,b.odyssey.weapons.ranged[0].magazineId);assert.equal(inventoryAmmoTotal(a.odyssey),22);
 assert.deepEqual(JSON.parse(JSON.stringify(sanitize(a))),JSON.parse(JSON.stringify(a)));
});
test('V1 shots/reloads update mounted magazine mirror, no ammo duplication',()=>{
 let t=sanitize(legacy()),o=t.odyssey;consumeAmmo(o,0,getAmmoShot(o,0,'ball',3));t=sanitize(t);o=t.odyssey;
 assert.equal(o.inventory.magazines[0].loaded,1);assert.equal(inventoryAmmoTotal(o),19);
 reloadMagazine(o,0,'ball');t=sanitize(t);o=t.odyssey;assert.equal(inventoryAmmoTotal(o),19);assert.equal(o.inventory.magazines[0].loaded,6);
 const id=o.weapons.ranged[0].magazineId;swapMagazine(o,0,'');t=sanitize(t);o=t.odyssey;assert.equal(o.weapons.ranged[0].loaded,0);assert.equal(inventoryAmmoTotal(o),19);
 emptyMagazine(o,id);assert.equal(inventoryAmmoTotal(o),19);fillMagazine(o,id,'ap');assert.equal(inventoryAmmoTotal(o),19);
 swapMagazine(o,0,id);assert.equal(o.weapons.ranged[0].loadedAmmoId,'ap');assert.equal(o.weapons.ranged[0].loaded,6);
});
test('Magazine exclusivity, compatibility and failed operations preserve ammo',()=>{
 const o=sanitize(legacy()).odyssey,id=o.weapons.ranged[0].magazineId;
 o.weapons.ranged.push({id:'gun2',name:'Second',capacity:6,loaded:0,loadedAmmoId:'',ammoIds:['ball'],magazineId:''});
 assert.throws(()=>swapMagazine(o,1,id),/already mounted/);assert.throws(()=>fillMagazine(o,id,'ap'),/Detach/);
 swapMagazine(o,0,'');o.weapons.ranged[1].capacity=5;assert.throws(()=>swapMagazine(o,1,id),/capacity/);assert.equal(inventoryAmmoTotal(o),22);
});
test('GM correction changes mounted and reserve magazines without spending reserve',()=>{
 let t=sanitize(legacy()),o=t.odyssey,id=o.weapons.ranged[0].magazineId;
 const reserve=JSON.stringify(o.ammunition);
 setMagazineRounds(o,id,2,'ball');t=sanitize(t);o=t.odyssey;
 assert.equal(o.weapons.ranged[0].loaded,2);assert.equal(o.inventory.magazines[0].loaded,2);
 assert.equal(JSON.stringify(o.ammunition),reserve);
 setMagazineRounds(o,id,6,'ap');t=sanitize(t);o=t.odyssey;
 assert.equal(o.weapons.ranged[0].loadedAmmoId,'ap');assert.equal(o.inventory.magazines[0].loaded,6);
 for(const invalid of [-1,7,1.5,NaN,Infinity,'2',null]) {
   const before=JSON.stringify(o);assert.throws(()=>setMagazineRounds(o,id,invalid,'ap'));assert.equal(JSON.stringify(o),before);
 }
 assert.throws(()=>setMagazineRounds(o,id,2,'missing'),/valid ammunition/);
 o.ammunition.push({id:'incompatible',name:'Other',quantity:0,damage:1,penetration:0});
 assert.throws(()=>setMagazineRounds(o,id,2,'incompatible'),/incompatible/);
 setMagazineRounds(o,id,0,'');assert.equal(o.weapons.ranged[0].loaded,0);assert.equal(o.weapons.ranged[0].loadedAmmoId,'');
 swapMagazine(o,0,'');setMagazineRounds(o,id,3,'incompatible');t=sanitize(t);o=t.odyssey;
 assert.equal(o.inventory.magazines[0].loaded,3);assert.equal(o.weapons.ranged[0].loaded,0);
 assert.equal(o.ammunition.find(a=>a.id==='incompatible').quantity,0);
});
test('Equipment uses one armor value and restores manual baseline; use consumes only one',()=>{
 const t=sanitize(legacy());t.odyssey.inventory.items=[{id:'armor',name:'Vest',kind:'armor',quantity:1,armor:12,equippedPart:'',previousArmor:0},{id:'med',kind:'medical',quantity:2}];
 equipItem(t,'armor','Torso');assert.equal(t.body.Torso.armor,18);unequipItem(t,'armor');assert.equal(t.body.Torso.armor,6);
 equipItem(t,'armor','Torso');t.body.Torso.armor=14;unequipItem(t,'armor');assert.equal(t.body.Torso.armor,14);
 const hp=t.body.Torso.current;consumeItem(t,'med');assert.equal(t.odyssey.inventory.items[1].quantity,1);assert.equal(t.body.Torso.current,hp);
});
test('Template strips origin ownership/events, restores health, rekeys all item links',()=>{
 const original=sanitize(legacy()),tpl=makeTemplate('Guard',original),json=JSON.stringify(tpl);
 assert.ok(!json.includes('discord-id'));assert.ok(!json.includes('external-id'));assert.ok(!json.includes('Past roll'));
 assert.equal(tpl.data.body.Torso.current,3);assert.equal(tpl.data.body.Torso.minor,0);assert.equal(tpl.data.body['Tail & claw'].current,4);
 const dest=sanitize({...legacy(),identity:{characterId:'destination'},odyssey:{...legacy().odyssey,owner:{playerId:'destination-owner'}}});
 const a=instantiateTemplate(dest,tpl,sanitize),b=instantiateTemplate(dest,tpl,sanitize);
 assert.equal(a.odyssey.owner.playerId,'destination-owner');assert.equal(a.identity.characterId,'destination');assert.equal(a.history[0].summary,'Past roll');
 assert.notEqual(a.odyssey.weapons.ranged[0].id,b.odyssey.weapons.ranged[0].id);
 assert.equal(a.odyssey.weapons.ranged[0].magazineId,a.odyssey.inventory.magazines[0].id);
 assert.equal(a.odyssey.weapons.ranged[0].loadedAmmoId,a.odyssey.ammunition[0].id);
 assert.equal(inventoryAmmoTotal(a.odyssey),22);assert.equal(a.odyssey.skills.Shoot,9);
});
test('JSON, browser storage and account scene payload round-trip; storage errors surface',()=>{
 const library={...emptyLibrary(),templates:[makeTemplate('Guard',sanitize(legacy()))]};
 const storage={value:null,getItem(){return this.value;},setItem(k,v){this.value=v;}};
 writeLibrary(storage,library);assert.equal(readLibrary(storage,sanitize).templates[0].name,'Guard');
 const scene=[{items:[{metadata:{[LIBRARY_ITEM_KEY]:library}}]}];assert.equal(libraryFromScenes(scene,sanitize).templates.length,1);
 assert.equal(mergeLibraries(library,library).templates.length,1);assert.throws(()=>libraryFromScenes([],sanitize),/no Odyssey/);
 assert.throws(()=>normalizeLibrary({...library,version:2},sanitize),/version 1/);
 assert.throws(()=>writeLibrary({setItem(){throw new Error('quota');}},library),/quota/);
 storage.value='{broken';assert.throws(()=>readLibrary(storage,sanitize));assert.equal(storage.value,'{broken');
});
test('Owner access is based on OBR owner, never external identity',()=>{
 const token={id:'token',type:'IMAGE',layer:'CHARACTER',metadata:{'com.codex.body-hp/data':legacy()}};
 // shared.js uses SDK isImage only as a type predicate; inject it at this boundary.
 ctx.isImage=x=>x.type==='IMAGE';assert.equal(ctx.canControl('PLAYER','owner',token),true);assert.equal(ctx.canControl('PLAYER','discord-id',token),false);assert.equal(ctx.canControl('GM','other',token),true);
});
test('V2 view uses full 0-10 scale and escapes names without exposing GM editors',()=>{
 const t=sanitize(legacy());t.odyssey.skillLabels.Shoot='<img src=x onerror=alert(1)>';
 const html=skillCards(t.odyssey,[['Combat',[['Shoot',9]]]],false);assert.ok(html.includes('9/10'));assert.ok(!html.includes('<img'));assert.ok(html.includes('&lt;img'));assert.ok(!html.includes('set-odyssey-skill'));
 assert.ok(!attributeCards(t.odyssey,[['Strength','Strength']],false).includes('set-odyssey-attribute'));
 assert.ok(attributeCards(t.odyssey,[['Strength','Strength']],true).includes('set-odyssey-attribute'));
 assert.ok(inventoryView(t,false).includes('Magazine'));
});
// Compare actual V1 characteristic/skill math against its unchanged base copy.
const main=readFileSync(new URL('../main.js',import.meta.url),'utf8');
const checks=main.slice(main.indexOf('function rollCharacterCheck('),main.indexOf('function renderOwnerFields('));
test('V1 roll boundaries and modifiers remain intact',()=>{
 let sequence=[];const math=Object.create(Math);math.random=()=>sequence.shift();const c=vm.createContext({Math:math});vm.runInContext(checks,c);
 sequence=[0];assert.equal(c.rollCharacterCheck(0,0).outcome,'critical-success');
 sequence=[.999];assert.equal(c.rollCharacterCheck(20,0).outcome,'critical-failure');
 sequence=[.45];assert.equal(c.rollCharacterCheck(9,1).finalAttribute,10);
 sequence=[.94,.4];assert.equal(c.rollSkillCheck(9,2).outcome,'critical-success');
 sequence=[.049,.4];assert.equal(c.rollSkillCheck(10,0).outcome,'critical-failure');
 sequence=[.49,.49];assert.equal(c.rollSkillCheck(9,-5).totalPrimary,135);
});


test('Empty skill groups vanish; pencils replace the expanded editor',()=>{
 const html=skillCards(sanitize(legacy()).odyssey,[['Abilities',[]],['Applied',[]]],true);
 assert.ok(!html.includes('Abilities'));assert.ok(!html.includes('Applied'));assert.ok(!html.includes('No skills.'));assert.ok(html.includes('profile-random-skill'));
 const skills=skillCards(sanitize(legacy()).odyssey,[['Combat',[['Shoot',9]]]],true);assert.ok(skills.includes('edit-skill'));assert.ok(!skills.includes('<details'));
});
test('Equipment, resource and stim bonuses derive skill levels without rewriting base values',()=>{
 let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'resource',name:'Guide',quantity:1,bonuses:{attributes:{Intelligence:2},skills:{Shoot:1}}},'guide');
 saveInventoryEntry(t,null,{kind:'implant',name:'Eye',quantity:1,armor:3,criticalMax:2,slot:'any',bonuses:{attributes:{Perception:3},skills:{Parry:1}},abilities:[{id:'vision',name:'Night vision',level:2}]},'eye');t=sanitize(t);
 let o=effectiveOdyssey(t.odyssey);assert.equal(o.skills.Shoot,10);assert.equal(o.skills.Parry,3);assert.equal(o.attributes.Perception,0);
 equipItem(t,'eye','Torso');o=effectiveOdyssey(t.odyssey);assert.equal(o.skills.Parry,4);assert.equal(o.attributes.Perception,3);assert.equal(o.skills['implant:eye:vision'],2);assert.equal(t.odyssey.skills.Shoot,9);
 useItemOnTracker(t,{subtype:'stim',name:'Focus',bonuses:{attributes:{Perception:2,Willpower:1}}},'','focus');assert.equal(effectiveOdyssey(t.odyssey).attributes.Perception,5);
 t.odyssey.effects=[];assert.equal(effectiveOdyssey(t.odyssey).attributes.Perception,3);unequipItem(t,'eye',true);assert.equal(effectiveOdyssey(t.odyssey).skills['implant:eye:vision'],undefined);
});
test('Implant critical capacity belongs to the body part and cannot heal by toggling',()=>{
 let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'implant',name:'Arm frame',quantity:1,armor:50,criticalMax:4,slot:'any'},'frame');t=sanitize(t);
 const before=t.body.Torso.current;equipItem(t,'frame','Torso');assert.equal(t.body.Torso.max,7);assert.equal(t.body.Torso.current,before);
 t=sanitize(sanitize(t));assert.equal(t.body.Torso.max,7);assert.equal(effectiveOdyssey(t.odyssey).armorSummary.total,0);
 const damage=absorbEquipmentCritical(t,'Torso',2);assert.equal(damage.remaining,2);assert.equal(t.odyssey.inventory.items[0].criticalDamage,0);
 assert.throws(()=>unequipItem(t,'frame',false),/Only the GM/);unequipItem(t,'frame',true);assert.equal(t.body.Torso.max,3);equipItem(t,'frame','Torso');assert.equal(t.body.Torso.current,before);
 t.odyssey.inventory.items[0].playerRemovable=true;unequipItem(t,'frame',false);assert.equal(t.body.Torso.max,3);
});
test('Armor thresholds count only worn armor; training level one removes penalties',()=>{
 for(const [armor,key,agility] of [[70,'light_armor',-1],[71,'medium_armor',-2],[140,'medium_armor',-2],[141,'heavy_armor',-4],[260,'heavy_armor',-4],[261,'superheavy_armor',-4],[350,'superheavy_armor',-4]]) {
   let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'armor',name:'Vest',quantity:1,armor,slot:'any'},'vest');t=sanitize(t);equipItem(t,'vest','Torso');let o=effectiveOdyssey(t.odyssey);assert.equal(o.armorSummary.key,key);assert.equal(o.modifiers.attributes.Agility,agility);
   t.odyssey.skills[key]=1;o=effectiveOdyssey(t.odyssey);assert.equal(o.modifiers.attributes.Agility,undefined);
 }
});
test('Armor absorbs critical wounds; medkits heal body, repair kits heal armor',()=>{
 let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'armor',name:'Vest',quantity:1,armor:12,criticalMax:2,slot:'any'},'vest');t=sanitize(t);equipItem(t,'vest','Torso');
 assert.deepEqual(absorbEquipmentCritical(t,'Torso',3),{remaining:1,absorbed:2});assert.equal(t.body.Torso.armor,6);
 useItemOnTracker(t,{subtype:'medkit',healCritical:1},'Torso');assert.equal(t.body.Torso.current,3);assert.equal(t.odyssey.inventory.items[0].criticalDamage,2);assert.equal(t.body.Torso.minor,3);
 useItemOnTracker(t,{subtype:'repairkit',healCritical:1},'vest');assert.equal(t.body.Torso.current,4);assert.equal(t.odyssey.inventory.items[0].criticalDamage,1);assert.equal(t.body.Torso.armor,18);
 assert.throws(()=>useItemOnTracker(t,{subtype:'grenade'},'Torso'),/placeholder/);
});
test('Type conversion restores body stats; loaded ammo cannot disappear',()=>{
 let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'implant',name:'Frame',quantity:1,armor:3,criticalMax:4,slot:'any'},'frame');t=sanitize(t);equipItem(t,'frame','Torso');
 saveInventoryEntry(t,{type:'item',id:'frame'},{kind:'resource',name:'Scrap',quantity:1},'frame');assert.equal(t.body.Torso.max,3);assert.equal(t.body.Torso.armor,6);assert.equal(t.odyssey.inventory.items[0].equippedPart,'');
 assert.throws(()=>deleteInventoryEntry(t,'ammo','ball'),/Empty all magazines/);
});



test('Passive item can grant a missing skill without changing the base character',()=>{
 let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'resource',name:'Pilot manual',quantity:1,bonuses:{skills:{Piloting:3}}},'pilot');t=sanitize(t);
 assert.equal(t.odyssey.skills.Piloting,undefined);assert.equal(effectiveOdyssey(t.odyssey).skills.Piloting,0);
 t.odyssey.inventory.items[0].quantity=0;assert.equal(effectiveOdyssey(t.odyssey).skills.Piloting,undefined);
});
test('Built-in magazines preserve rounds without creating a physical inventory magazine',()=>{
 let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'weapon',weaponKind:'ranged',name:'Internal rifle',acceptsMagazines:false,capacity:8,ammoIds:['ball'],loaded:2,loadedAmmoId:'ball'},'internal');t=sanitize(t);
 let o=t.odyssey,index=o.weapons.ranged.findIndex(w=>w.id==='internal');assert.equal(o.inventory.magazines.length,1);assert.equal(o.weapons.ranged[index].loaded,2);const total=inventoryAmmoTotal(o);
 reloadMagazine(o,index,'ball');t=sanitize(t);o=t.odyssey;assert.equal(inventoryAmmoTotal(o),total);assert.equal(o.weapons.ranged[index].loaded,8);assert.equal(o.inventory.magazines.length,1);
 consumeAmmo(o,index,getAmmoShot(o,index,'ball',2));t=sanitize(t);o=t.odyssey;assert.equal(o.weapons.ranged[index].loaded,6);assert.equal(inventoryAmmoTotal(o),total-2);
 assert.throws(()=>swapMagazine(o,index,o.inventory.magazines[0].id),/built-in/);deleteInventoryEntry(t,'weapon-ranged','internal');assert.equal(inventoryAmmoTotal(t.odyssey),total-2);
});
test('Detachable weapons accept chosen magazines with their own capacity and ammo',()=>{
 let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'magazine',name:'Large AP',capacity:12,loaded:7,loadedAmmoId:'ap'},'large');
 saveInventoryEntry(t,null,{kind:'weapon',weaponKind:'ranged',name:'Modular',acceptsMagazines:true,acceptedMagazineIds:['large']},'modular');t=sanitize(t);
 const o=t.odyssey,index=o.weapons.ranged.findIndex(w=>w.id==='modular'),total=inventoryAmmoTotal(o);
 assert.throws(()=>swapMagazine(o,index,o.inventory.magazines[0].id),/already mounted|incompatible/);
 swapMagazine(o,index,'large');assert.equal(o.weapons.ranged[index].capacity,12);assert.equal(getAmmoShot(o,index,'ap',3).damage,24);consumeAmmo(o,index,getAmmoShot(o,index,'ap',3));t=sanitize(t);
 assert.equal(inventoryAmmoTotal(t.odyssey),total-3);assert.equal(t.odyssey.inventory.magazines.find(m=>m.id==='large').loaded,4);
 assert.throws(()=>saveInventoryEntry(t,{type:'weapon-ranged',id:'modular'},{kind:'weapon',weaponKind:'ranged',name:'Modular',acceptsMagazines:false,capacity:12,loaded:0,ammoIds:['ap']},'modular'),/Empty or detach/);
 swapMagazine(t.odyssey,index,'');assert.equal(t.odyssey.weapons.ranged[index].capacity,0);assert.equal(t.odyssey.inventory.magazines.find(m=>m.id==='large').loaded,4);
});



test('Players see armor training only at level one; templates are GM only',()=>{
 const o=sanitize(legacy()).odyssey;assert.ok(!skillCards(o,[],false).includes('Armor training'));assert.ok(skillCards(o,[],true).includes('Light Armor 0'));
 o.skills.light_armor=1;const html=skillCards(o,[],false);assert.ok(html.includes('Light Armor 1'));assert.ok(!html.includes('Medium Armor'));
});
test('Several allowed slots, one armor per part and one hundred implants per part',()=>{
 let t=sanitize(legacy());t.body.Arm={...t.body.Torso,slot:'left-arm',current:3,max:3};
 saveInventoryEntry(t,null,{kind:'armor',name:'Vest',quantity:1,slots:['Torso','Arm'],criticalMax:2,armor:5},'armor1');saveInventoryEntry(t,null,{kind:'armor',name:'Second',quantity:1,slots:['Torso','Arm'],criticalMax:2,armor:5},'armor2');t=sanitize(t);
 assert.throws(()=>equipItem(t,'armor1','Tail & claw'),/does not fit/);equipItem(t,'armor1','Arm');assert.throws(()=>equipItem(t,'armor2','Arm'),/already equipped/);
 for(let i=0;i<100;i++){saveInventoryEntry(t,null,{kind:'implant',name:'Implant '+i,quantity:1,slots:['Arm'],criticalMax:1,armor:0,playerRemovable:true},'impl'+i);equipItem(t,'impl'+i,'Arm',false);}
 t=sanitize(sanitize(t));assert.equal(t.body.Arm.max,105);assert.equal(t.body.Arm.current,5);assert.equal(t.odyssey.inventory.items.filter(i=>i.equippedPart==='Arm').length,101);
 t.odyssey.inventory.items.find(i=>i.id==='impl0').playerRemovable=false;unequipItem(t,'impl0',true);assert.throws(()=>equipItem(t,'impl0','Arm',false),/Only the GM/);
});
test('Armor adds HP and Max, retains damage, and keeps healing and repair separate',()=>{
 let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'armor',name:'Vest',quantity:1,slots:['Torso'],criticalMax:4,armor:12},'vest');t=sanitize(t);equipItem(t,'vest','Torso');
 assert.equal(t.body.Torso.current,6);assert.equal(t.body.Torso.max,7);t=sanitize(sanitize(t));assert.equal(t.body.Torso.current,6);
 setCombinedPartHp(t,'Torso',2);assert.equal(t.odyssey.inventory.items[0].criticalDamage,4);assert.equal(t.body.Torso.current,2);
 useItemOnTracker(t,{subtype:'medkit',healCritical:5},'Torso');assert.equal(t.body.Torso.current,3);assert.throws(()=>useItemOnTracker(t,{subtype:'medkit',healCritical:1},'Torso'),/No critical wounds/);
 useItemOnTracker(t,{subtype:'repairkit',healCritical:2},'vest');assert.equal(t.body.Torso.current,5);unequipItem(t,'vest');assert.equal(t.body.Torso.current,3);assert.equal(t.body.Torso.max,3);
 equipItem(t,'vest','Torso');assert.equal(t.body.Torso.current,5);assert.equal(t.body.Torso.max,7);assert.equal(t.odyssey.inventory.items[0].criticalDamage,2);
 setCombinedPartMax(t,'Torso',10);t=sanitize(t);assert.equal(t.body.Torso.max,10);unequipItem(t,'vest');assert.equal(t.body.Torso.max,6);
});
test('Already equipped V3 armor migrates once; healthy templates preserve the combined pool',()=>{
 const raw=legacy();raw.odyssey.inventory={version:3,items:[{id:'vest',kind:'armor',name:'Vest',quantity:1,armor:12,criticalMax:3,criticalDamage:1,equippedPart:'Torso',previousArmor:6,legacyEquipment:false}],baseArmor:{Torso:6},baseMax:{},magazines:[]};
 let t=sanitize(raw);assert.equal(t.body.Torso.current,4);assert.equal(t.body.Torso.max,6);t=sanitize(t);assert.equal(t.body.Torso.current,4);assert.equal(t.body.Torso.max,6);
 const template=makeTemplate('Armored',t),fresh=instantiateTemplate(t,template,sanitize);assert.equal(fresh.body.Torso.current,6);assert.equal(fresh.body.Torso.max,6);assert.equal(fresh.odyssey.inventory.items[0].criticalDamage,0);
});

test('Range mitigation survives sanitization, templates and equipped/passive rules',()=>{
 const raw=legacy();raw.odyssey.weapons.ranged[0].weaponCategory='pistol';raw.odyssey.weapons.ranged[0].rangeReduction={clinch:5};raw.odyssey.skillRangeReduction={Shoot:{long:10},Unknown:{long:999}};
 let t=sanitize(raw);assert.equal(t.odyssey.weapons.ranged[0].weaponCategory,'pistol');assert.equal(t.odyssey.weapons.ranged[0].rangeReduction.clinch,5);assert.equal(t.odyssey.skillRangeReduction.Shoot.long,10);assert.equal(t.odyssey.skillRangeReduction.Unknown,undefined);
 saveInventoryEntry(t,null,{kind:'resource',name:'Scope training',quantity:1,bonuses:{rangeReduction:{long:3}}},'resource');
 saveInventoryEntry(t,null,{kind:'implant',name:'Eye',quantity:1,slots:['Torso'],bonuses:{rangeReduction:{long:4}}},'eye');
 assert.equal(effectiveOdyssey(t.odyssey).rangeReduction.long,0);equipItem(t,'eye','Torso');assert.equal(effectiveOdyssey(t.odyssey).rangeReduction.long,0);unequipItem(t,'eye','Torso');assert.equal(effectiveOdyssey(t.odyssey).rangeReduction.long,0);
 const template=makeTemplate('Shooter',t),copy=instantiateTemplate(sanitize(legacy()),template,sanitize);assert.equal(copy.odyssey.skillRangeReduction.Shoot.long,10);assert.equal(copy.odyssey.weapons.ranged[0].rangeReduction.clinch,5);assert.equal(copy.odyssey.weapons.ranged[0].weaponCategory,'pistol');
});
test('Inventory repair targets a single stored unit and never heals the body',()=>{
 let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'armor',name:'Spare',quantity:2,criticalMax:3,criticalDamage:2,slots:['Torso']},'spares');const hp=t.body.Torso.current;
 assert.throws(()=>useItemOnTracker(t,{subtype:'repairkit',healCritical:1},'Torso'),/inventory/);
 useItemOnTracker(t,{subtype:'repairkit',healCritical:1},'spares','repaired');assert.equal(t.body.Torso.current,hp);assert.equal(t.odyssey.inventory.items.find(i=>i.id==='spares').quantity,1);assert.equal(t.odyssey.inventory.items.find(i=>i.id==='spares').criticalDamage,2);assert.equal(t.odyssey.inventory.items.find(i=>i.id==='repaired').criticalDamage,1);
 useItemOnTracker(t,{subtype:'repairkit',healCritical:99},'repaired');assert.equal(t.odyssey.inventory.items.find(i=>i.id==='repaired').criticalDamage,0);assert.throws(()=>useItemOnTracker(t,{subtype:'repairkit',healCritical:1},'repaired'),/damaged/);
});
test('Item-granted unknown skill stays zero; a learned skill receives bonuses by key or label',()=>{
 let t=sanitize(legacy());saveInventoryEntry(t,null,{kind:'resource',name:'Manual',quantity:1,bonuses:{skills:{Piloting:3}}},'manual');let o=effectiveOdyssey(t.odyssey);assert.equal(o.skills.Piloting,0);assert.equal(o.modifiers.skills.Piloting,undefined);
 t.odyssey.skills.learnedPilot=0;t.odyssey.skillLabels.learnedPilot='Piloting';o=effectiveOdyssey(t.odyssey);assert.equal(o.skills.learnedPilot,3);assert.equal(o.skills.Piloting,undefined);assert.equal(t.odyssey.skills.learnedPilot,0);
});
test('Descriptions survive templates, inventory and legacy string perks',()=>{
 const raw=legacy();raw.odyssey.skillDescriptions={Shoot:'Aim & fire <safely>'};raw.odyssey.skillPerkDescriptions={Shoot:{Accurate:'Keep steady.'}};raw.odyssey.ammunition[0].description='Ammo description';raw.odyssey.weapons.ranged[0].description='Rifle description';let t=sanitize(raw);t.odyssey.inventory.magazines[0].description='Magazine description';
 const template=makeTemplate('Descriptions',t),copy=instantiateTemplate(sanitize(legacy()),template,sanitize);assert.equal(copy.odyssey.skillDescriptions.Shoot,'Aim & fire <safely>');assert.equal(copy.odyssey.skillPerkDescriptions.Shoot.Accurate,'Keep steady.');assert.equal(copy.odyssey.ammunition[0].description,'Ammo description');assert.equal(copy.odyssey.weapons.ranged[0].description,'Rifle description');assert.equal(copy.odyssey.inventory.magazines[0].description,'Magazine description');
});
console.log(`${count} integration checks passed.`);
