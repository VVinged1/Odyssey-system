import {findInventoryEntry,saveInventoryEntry,newId,magazineCompatible} from './inventory.js';
export const ITEM_POOL_PREFIX='com.codex.body-hp/item-pool/';
export const ITEM_POOL_FORMAT='odyssey-item-pool';
const TYPES=new Set(['item','ammo','magazine','weapon-melee','weapon-ranged']);
const stable=value=>Array.isArray(value)?value.map(stable):value&&typeof value==='object'?Object.fromEntries(Object.keys(value).sort().map(key=>[key,stable(value[key])])):value;
export const blueprintText=blueprint=>JSON.stringify(stable(blueprint));
const ammoDefinition=a=>({name:a.name,damage:a.damage,penetration:a.penetration});
const sameAmmo=(a,b)=>a.name===b.name&&a.damage===b.damage&&a.penetration===b.penetration;
const sameMagazine=(a,b)=>a.name===b.name&&a.capacity===b.capacity;
function uniqueDefinitions(rows,project) {
  return [...new Map(rows.map(row=>{const value=project(row);return [blueprintText(value),value];})).entries()].sort(([a],[b])=>a.localeCompare(b)).map(([,value])=>value);
}
// Catalog definitions exclude token IDs, equipped state, damage and loaded rounds.
export function captureBlueprint(tracker,type,id) {
  const o=tracker.odyssey,entry=findInventoryEntry(tracker,type,id);if(!entry)throw new Error('Item no longer exists.');
  let values;const dependencies={ammunition:[],magazines:[]};
  if(type==='ammo')values={kind:'ammo',...ammoDefinition(entry),quantity:1};
  else if(type==='magazine')values={kind:'magazine',name:entry.name,quantity:1,capacity:entry.capacity,loaded:0,loadedAmmoId:''};
  else if(type.startsWith('weapon-')) {
    const kind=type.slice(7);values={kind:'weapon',weaponKind:kind,name:entry.name,accuracy:entry.accuracy,damage:entry.damage,...(kind==='ranged'&&entry.weaponCategory?{weaponCategory:entry.weaponCategory}:{}),...(kind==='ranged'&&Object.values(entry.rangeReduction||{}).some(Boolean)?{rangeReduction:structuredClone(entry.rangeReduction)}:{})};
    if(kind==='ranged') {
      values.acceptsMagazines=entry.acceptsMagazines!==false;values.loaded=0;values.loadedAmmoId='';values.magazineId='';
      if(values.acceptsMagazines) {
        let magazines=o.inventory.magazines.filter(m=>m.id===entry.magazineId||magazineCompatible(o,entry,m));
        if(!magazines.length&&entry.capacity)magazines=[{name:entry.name+' magazine',capacity:entry.capacity}];
        dependencies.magazines=uniqueDefinitions(magazines,m=>({name:m.name,capacity:m.capacity})).map((m,index)=>({id:'mag-'+index,...m}));values.acceptedMagazineIds=dependencies.magazines.map(m=>m.id);
      } else {
        dependencies.ammunition=uniqueDefinitions(o.ammunition.filter(a=>entry.ammoIds.includes(a.id)),ammoDefinition).map((a,index)=>({id:'ammo-'+index,...a}));
        values.capacity=entry.capacity;values.ammoIds=dependencies.ammunition.map(a=>a.id);
      }
    }
  } else {
    values={};for(const key of ['kind','name','description','armor','criticalMax','playerRemovable','slots','subtype','healCritical','abilities','bonuses'])if(entry[key]!==undefined)values[key]=structuredClone(entry[key]);
    values.quantity=1;values.criticalDamage=0;values.slots??=[entry.slot||'any'];
    if(values.bonuses) {
      const skills={};for(const [key,value]of Object.entries(values.bonuses.skills||{})){const label=o.skillLabels[key]||key;skills[label]=(skills[label]||0)+value;}values.bonuses.skills=skills;if(values.bonuses.skillCategories){values.bonuses.skillCategories=Object.fromEntries(Object.entries(values.bonuses.skillCategories).map(([key,value])=>[o.skillLabels[key]||key,value]));}
    }
    if(values.abilities)values.abilities=values.abilities.map((ability,index)=>({...ability,id:'ability-'+index,...(ability.skill?{skill:o.skillLabels[ability.skill]||ability.skill}:{}),...(ability.bonuses?{bonuses:{...ability.bonuses,skills:Object.fromEntries(Object.entries(ability.bonuses.skills||{}).map(([key,value])=>[o.skillLabels[key]||key,value]))}}:{})}));
  }
  if(entry.useSkill)values.useSkill=String(o.skillLabels[entry.useSkill]||entry.useSkill).trim().slice(0,80);
  if(entry.description)values.description=String(entry.description).slice(0,2000);
  return {format:'odyssey-item-blueprint',version:1,type,values,dependencies};
}
export function validateBlueprint(raw) {
  if(raw?.format!=='odyssey-item-blueprint'||raw.version!==1||!TYPES.has(raw.type)||!raw.values||typeof raw.values.name!=='string'||!raw.values.name.trim()||raw.values.name.length>120)throw new Error('Invalid item pool definition.');
  if(!['item','consumable','medical','tool','resource','armor','shield','implant','ammo','magazine','weapon'].includes(raw.values.kind))throw new Error('Invalid item type.');
  const expected=raw.type.startsWith('weapon-')?'weapon':raw.type==='item'?null:raw.type;if(raw.type==='item'&&['weapon','ammo','magazine'].includes(raw.values.kind))throw new Error('Item type does not match its definition.');
  if(expected&&raw.values.kind!==expected)throw new Error('Item type does not match its definition.');
  for(const key of ['ammunition','magazines'])if(!Array.isArray(raw.dependencies?.[key])||raw.dependencies[key].length>200||raw.dependencies[key].some(d=>!d||typeof d.id!=='string'||typeof d.name!=='string'))throw new Error('Invalid item dependencies.');
  if(blueprintText(raw).length>200000)throw new Error('Item definition is too large.');
  return structuredClone(raw);
}
export async function blueprintId(blueprint) {
  const hash=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(blueprintText(blueprint)));return [...new Uint8Array(hash)].map(n=>n.toString(16).padStart(2,'0')).join('');
}
export function poolEntries(metadata) {
  const entries=[];let errors=0;
  for(const [key,value]of Object.entries(metadata||{}))if(key.startsWith(ITEM_POOL_PREFIX)&&value&&!value.deleted){try{entries.push({id:key.slice(ITEM_POOL_PREFIX.length),blueprint:validateBlueprint(value.blueprint)});}catch{errors++;}}
  return {entries:entries.sort((a,b)=>a.blueprint.values.name.localeCompare(b.blueprint.values.name)),errors};
}
export function parsePoolJSON(raw) {
  if(raw?.format!==ITEM_POOL_FORMAT||raw.version!==1||!Array.isArray(raw.entries)||raw.entries.length>2000)throw new Error('Choose an Odyssey item pool JSON (version 1).');return raw.entries.map(validateBlueprint);
}
export function exportPool(entries) {return {format:ITEM_POOL_FORMAT,version:1,entries:entries.map(e=>e.blueprint)};}
export function copyBlueprint(tracker,raw,quantity=1,id=newId()) {
  const blueprint=validateBlueprint(raw),o=tracker.odyssey,values=structuredClone(blueprint.values),ammoIds=new Map(),magazineIds=new Map();
  if(!Number.isInteger(quantity)||quantity<1||quantity>999999)throw new Error('Invalid item quantity.');
  for(const [index,definition]of blueprint.dependencies.ammunition.entries()) {
    let ammo=o.ammunition.find(a=>sameAmmo(a,definition));if(!ammo){if(o.ammunition.length>=200)throw new Error('Ammunition type limit reached.');ammo={...definition,id:id+'-ammo-'+index,quantity:0};o.ammunition.push(ammo);}ammoIds.set(definition.id,ammo.id);
  }
  for(const [index,definition]of blueprint.dependencies.magazines.entries()) {
    let matches=o.inventory.magazines.filter(m=>sameMagazine(m,definition));
    if(!matches.length){const magazine={id:id+'-mag-'+index,name:definition.name,capacity:definition.capacity,loaded:0,loadedAmmoId:''};if(o.inventory.magazines.length>=1000)throw new Error('Magazine inventory is full.');o.inventory.magazines.push(magazine);matches=[magazine];}magazineIds.set(definition.id,matches.map(m=>m.id));
  }
  values.equippedPart='';values.previousArmor=0;values.legacyEquipment=false;values.criticalDamage=0;
  if(blueprint.type.startsWith('weapon-'))values.weaponKind=blueprint.type.slice(7);
  if(blueprint.type==='magazine'||blueprint.type.startsWith('weapon-')){values.loaded=0;values.loadedAmmoId='';values.magazineId='';}
  if(values.ammoIds)values.ammoIds=values.ammoIds.map(key=>ammoIds.get(key)).filter(Boolean);
  if(values.acceptedMagazineIds)values.acceptedMagazineIds=values.acceptedMagazineIds.flatMap(key=>magazineIds.get(key)||[]);
  if(values.bonuses?.skills) {
    const skills={};for(const [label,value]of Object.entries(values.bonuses.skills)){const key=Object.keys(o.skills).find(key=>key.toLowerCase()===label.toLowerCase()||(o.skillLabels[key]||key).toLowerCase()===label.toLowerCase())||label;skills[key]=(skills[key]||0)+value;}const categories={};for(const [label,value]of Object.entries(values.bonuses.skillCategories||{})){const key=Object.keys(o.skills).find(key=>key.toLowerCase()===label.toLowerCase()||(o.skillLabels[key]||key).toLowerCase()===label.toLowerCase())||label;categories[key]=value;}if(Object.keys(categories).length)values.bonuses.skillCategories=categories;values.bonuses.skills=skills;
  }
  if(blueprint.type==='ammo') {
    const existing=o.ammunition.find(a=>sameAmmo(a,values));if(existing){if(existing.quantity+quantity>999999)throw new Error('Ammo reserve is full.');existing.quantity+=quantity;return existing.id;}
  }
  if(blueprint.type.startsWith('weapon-')) {
    if(quantity>20)throw new Error('Weapon limit is 20 per type.');
    for(let i=0;i<quantity;i++) {
      const base=values.name;let name=base,suffix=2;while([...o.weapons.melee,...o.weapons.ranged].some(w=>w.name===name))name=base+' ('+(suffix++)+')';
      saveInventoryEntry(tracker,null,{...values,name,quantity:1},i?id+'-copy-'+i:id);
    }
  } else saveInventoryEntry(tracker,null,{...values,quantity},id);
  return id;
}

export const ITEM_POOL_ASSET_KEY='com.codex.body-hp/item-pool-library';
export function poolFromScenes(scenes) {
 const entries=new Map();let found=false;
 for(const scene of scenes||[])for(const item of scene.items||[]){const payload=item.metadata?.[ITEM_POOL_ASSET_KEY];if(!payload)continue;found=true;for(const blueprint of parsePoolJSON(payload))entries.set(blueprintText(blueprint),blueprint);}
 if(!found)throw Error('Selected scenes contain no Odyssey item pool library.');
 if(entries.size>2000)throw Error('Item pool library exceeds 2000 definitions.');
 const result=[...entries.values()];if(new TextEncoder().encode(JSON.stringify(result)).length>4*1024*1024)throw Error('Item pool library exceeds 4 MB.');return result;
}
