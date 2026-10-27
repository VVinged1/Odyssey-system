import {weaponCategoryLabel} from './rangeRules.js';
export const ARMOR_POOL_TIERS=['Light Armor','Medium Armor','Heavy Armor','Superheavy Armor','Outside listed ranges'];
const SLOTS={Torso:['Torso',true],torso:['Torso',true],Head:['Head',false],head:['Head',false],'L.Arm':['Left Arm',false],'left-arm':['Left Arm',false],'R.Arm':['Right Arm',false],'right-arm':['Right Arm',false],'L.Leg':['Left Leg',false],'left-leg':['Left Leg',false],'R.Leg':['Right Leg',false],'right-leg':['Right Leg',false],Shield:['Shield',false],Special:['Special',false],other:['Other parts',false]};
export function armorPoolTier(armor,torso=false) {
  const limits=torso?[20,40,60,100]:[10,20,40,60];const index=limits.findIndex(max=>Number(armor||0)<=max);return ARMOR_POOL_TIERS[index<0?4:index];
}
export function poolCategoryPaths(blueprint) {
  const v=blueprint.values;
  if(v.kind==='armor'||v.kind==='shield'||v.kind==='implant') {
    let slots=Array.isArray(v.slots)?v.slots:[v.slot||'any'];if(slots.includes('any'))slots=['Torso','Head','L.Arm','R.Arm','L.Leg','R.Leg','other'];
    const parts=[...new Map(slots.map(key=>{const value=SLOTS[key]||[key,false];return [value[0],value];})).values()];
    return parts.map(([label,torso])=>v.kind==='implant'?['Implants',label]:['Armor',armorPoolTier(v.armor,torso),label]);
  }
  if(v.kind==='consumable'||v.kind==='medical')return [['Usables',({grenade:'Grenades',medkit:'Medkits',repairkit:'Repair kits',stim:'Stims'})[v.subtype]||'Other usables']];
  if(v.kind==='weapon')return [v.weaponKind==='ranged'?['Weapons','Ranged',weaponCategoryLabel(v.weaponCategory)]:['Weapons','Melee']];
  return [[({ammo:'Ammunition',magazine:'Magazines',resource:'Resources',tool:'Tools'})[v.kind]||'Other items']];
}
export function groupPoolEntries(entries) {
  const root={children:new Map(),entries:new Map()};
  for(const entry of entries)for(const path of poolCategoryPaths(entry.blueprint)) {
    let node=root;for(const label of path){if(!node.children.has(label))node.children.set(label,{label,children:new Map(),entries:new Map(),all:new Set()});node=node.children.get(label);node.all.add(entry.id);}node.entries.set(entry.id,entry);
  }
  const order=['Usables','Resources','Ammunition','Magazines','Weapons','Armor','Implants','Tools','Other items',...ARMOR_POOL_TIERS,'Torso','Head','Left Arm','Right Arm','Left Leg','Right Leg','Shield','Special','Other parts'];
  const result=node=>[...node.children.values()].sort((a,b)=>{const ai=order.indexOf(a.label),bi=order.indexOf(b.label);return (ai<0?999:ai)-(bi<0?999:bi)||a.label.localeCompare(b.label);}).map(child=>({label:child.label,count:child.all.size,entries:[...child.entries.values()],children:result(child)}));
  return result(root);
}
