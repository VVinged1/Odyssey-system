import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { normalizeInventory, swapMagazine, fillMagazine, emptyMagazine, setMagazineRounds, equipItem, unequipItem, consumeItem, inventoryAmmoTotal } from '../inventory.js';
import { sanitizeAmmunition, sanitizeMagazine, reloadMagazine, getAmmoShot, consumeAmmo } from '../ammunition.js';
import { makeTemplate, instantiateTemplate, profileData, normalizeLibrary, emptyLibrary, mergeLibraries, readLibrary, writeLibrary, libraryFromScenes, LIBRARY_ITEM_KEY, LIBRARY_STORAGE_KEY } from '../profiles.js';
import { attributeCards, skillCards, inventoryView } from '../ui/characterView.js';
const source=readFileSync(new URL('../shared.js',import.meta.url),'utf8').replace(/^import .*;\r?$/gm,'').replace(/^export \{ OBR \};\r?$/gm,'').replace(/^export /gm,'');
const ctx=vm.createContext({structuredClone,normalizeInventory,sanitizeAmmunition,sanitizeMagazine,OBR:{},console});
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
 equipItem(t,'armor','Torso');assert.equal(t.body.Torso.armor,12);unequipItem(t,'armor');assert.equal(t.body.Torso.armor,6);
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
console.log(`${count} integration checks passed.`);
