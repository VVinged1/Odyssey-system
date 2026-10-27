// Derived character values. Base attributes/skills are never overwritten by effects.
export const ARMOR_TRAINING = {light_armor:'Light Armor',medium_armor:'Medium Armor',heavy_armor:'Heavy Armor',superheavy_armor:'Superheavy Armor'};
export const ATTRIBUTES = ['Strength','Agility','Reaction','Endurance','Perception','Intelligence','Charisma','Willpower','Magic'];
const n = v => Math.max(-999,Math.min(999,Math.trunc(Number(v)||0)));
export function normalizeBonuses(raw) {
  const clean=(rows,attributes=false)=>Object.fromEntries(Object.entries(rows||{}).filter(([k,v])=>k.length<=80&&!['__proto__','constructor','prototype'].includes(k)&&(!attributes||ATTRIBUTES.includes(k))&&n(v)).slice(0,80).map(([k,v])=>[k,n(v)]));
  return {attributes:clean(raw?.attributes,true),skills:clean(raw?.skills)};
}
export function normalizeEffects(raw) {
  const seen=new Set();return (Array.isArray(raw)?raw:[]).slice(0,100).filter(e=>e&&typeof e.id==='string'&&!seen.has(e.id)&&seen.add(e.id)).map(e=>({id:e.id.slice(0,120),name:String(e.name||'Effect').slice(0,120),bonuses:normalizeBonuses(e.bonuses)}));
}
export function equipmentFits(item,partName,part) {
  const slots=Array.isArray(item.slots)?item.slots:[item.slot||'any'];
  return slots.includes('any')||slots.includes(partName)||slots.includes(part?.slot);
}
export function itemActive(item) { return item.quantity>0 && (item.kind==='implant' || !item.criticalMax || item.criticalDamage<item.criticalMax); }
export function armorSummary(o) {
  const total=(o.inventory?.items||[]).filter(i=>['armor','shield'].includes(i.kind)&&i.equippedPart&&itemActive(i)).reduce((sum,i)=>sum+i.armor,0);
  const key=total<=0?null:total<=70?'light_armor':total<=140?'medium_armor':total<=260?'heavy_armor':'superheavy_armor';
  // SOURCE boundaries: 1–70, 71–140, 141–260, 261+. Trained removes penalties per requested rule.
  return {total,key,name:key?ARMOR_TRAINING[key]:'No armor'};
}
export function effectiveOdyssey(base) {
  const o=structuredClone(base),rows=[];o.baseAttributes=structuredClone(base.attributes);o.baseSkills=structuredClone(base.skills);o.grantedSkills={};
  for(const i of o.inventory?.items||[]) {
    if(!itemActive(i))continue;
    if(i.kind==='resource'||(['armor','shield','implant'].includes(i.kind)&&i.equippedPart)) rows.push({id:i.id,name:i.name,bonuses:i.bonuses,passive:true});
    if(i.kind==='implant'&&i.equippedPart)for(const a of i.abilities||[]) {
      const key=`implant:${i.id}:${a.id}`;o.skills[key]=a.level;o.skillCategories[key]='abilities';o.skillLabels[key]=a.name;o.skillPerks[key]=a.description?[a.description]:[];o.grantedSkills[key]=i.name;
    }
  }
  rows.push(...(base.effects||[]));
  // Items may grant skills absent from the character; these disappear with the item.
  for(const row of rows)for(const key of Object.keys(row.bonuses?.skills||{}))if(!Object.hasOwn(o.skills,key)) {
    o.skills[key]=0;o.baseSkills[key]=0;o.skillLabels[key]=key;o.skillCategories[key]='applied';o.skillStrengthBonuses[key]=false;o.skillPerks[key]=[];o.grantedSkills[key]=row.name;
  }
  const skillMods={};for(const row of rows)for(const [k,v] of Object.entries(row.bonuses?.skills||{}))skillMods[k]=(skillMods[k]||0)+v;
  for(const [k,v] of Object.entries(skillMods))if(Object.hasOwn(o.skills,k))o.skills[k]=Math.max(0,Math.min(Object.hasOwn(ARMOR_TRAINING,k)?1:10,o.skills[k]+v));
  const armor=armorSummary(o),trained=!armor.key||(o.skills[armor.key]||0)>=1;
  if(!trained) {
    const penalties=armor.key==='light_armor'?{Agility:-1,Endurance:-1}:armor.key==='medium_armor'?{Agility:-2,Endurance:-2,Strength:-2}:{Agility:-4,Endurance:-5,Strength:-4};
    rows.push({id:'armor-penalty',name:`${armor.name} · untrained (${armor.total})`,bonuses:{attributes:penalties,skills:{}},passive:true});
  }
  o.modifiers={attributes:{},skills:skillMods};
  for(const row of rows)for(const [k,v] of Object.entries(row.bonuses?.attributes||{}))o.modifiers.attributes[k]=(o.modifiers.attributes[k]||0)+v;
  for(const k of ATTRIBUTES)o.attributes[k]=Math.max(0,(base.attributes[k]||0)+(o.modifiers.attributes[k]||0));
  o.effectRows=rows.filter(r=>Object.keys(r.bonuses?.attributes||{}).length||Object.keys(r.bonuses?.skills||{}).length);o.armorSummary={...armor,trained};return o;
}
export function reconcileEquipment(tracker) {
  const inv=tracker.odyssey.inventory;if(!inv)return;
  inv.baseArmor??={}; inv.baseMax??={};
  for(const [partName,part] of Object.entries(tracker.body)) {
    const equipped=inv.items.filter(i=>['armor','shield','implant'].includes(i.kind)&&i.equippedPart===partName);
    const implantHp=equipped.filter(i=>i.kind==='implant'&&itemActive(i)).reduce((sum,i)=>sum+(Number(i.criticalMax)||0),0);
    const armor=equipped.filter(i=>['armor','shield'].includes(i.kind)&&i.quantity>0);
    const pool={max:armor.reduce((sum,i)=>sum+(Number(i.criticalMax)||0),0),remaining:armor.reduce((sum,i)=>sum+Math.max(0,(Number(i.criticalMax)||0)-(Number(i.criticalDamage)||0)),0)};
    inv.armorHp??={};const previous=inv.armorHp[partName]?.remaining||0;
    const hpBonus=implantHp+pool.max;
    if(hpBonus&&!Object.hasOwn(inv.baseMax,partName))inv.baseMax[partName]=part.max;
    if(Object.hasOwn(inv.baseMax,partName)) {
      part.max=Math.min(999,inv.baseMax[partName]+hpBonus);part.current=Math.max(0,Math.min(part.current+pool.remaining-previous,part.max));
      if(!hpBonus)delete inv.baseMax[partName];
    }
    if(pool.max)inv.armorHp[partName]=pool;else delete inv.armorHp[partName];
    if(!equipped.length) {
      if(Object.hasOwn(inv.baseArmor,partName)){part.armor=inv.baseArmor[partName];delete inv.baseArmor[partName];}
      continue;
    }
    if(!Object.hasOwn(inv.baseArmor,partName)) {
      // V2 token migration: old equipment replaced armor and recorded its baseline.
      const legacy=equipped.find(i=>i.legacyEquipment);inv.baseArmor[partName]=legacy?legacy.previousArmor:part.armor;
    }
    part.armor=inv.baseArmor[partName]+equipped.filter(itemActive).reduce((sum,i)=>sum+i.armor,0);
    for(const i of equipped)i.legacyEquipment=false;
  }
}
// Critical damage is absorbed by equipped durability before reaching body HP.
// Minor/serious promotion still follows V1 rules. Special protection runs first.
export function absorbEquipmentCritical(tracker,partName,critical) {
  let remaining=Math.max(0,Number(critical)||0),absorbed=0;
  const items=tracker.odyssey.inventory?.items||[];
  for(const i of items.filter(i=>['armor','shield'].includes(i.kind)&&i.equippedPart===partName&&i.criticalMax>0&&itemActive(i)).sort((a,b)=>(a.kind==='implant')-(b.kind==='implant'))) {
    const amount=Math.min(remaining,i.criticalMax-i.criticalDamage);i.criticalDamage+=amount;remaining-=amount;absorbed+=amount;if(!remaining)break;
  }
  reconcileEquipment(tracker);return {remaining,absorbed};
}
export function armorHealth(tracker,partName) {
  const gear=tracker.odyssey.inventory.items.filter(i=>['armor','shield'].includes(i.kind)&&i.equippedPart===partName&&i.quantity>0);
  return {max:gear.reduce((sum,i)=>sum+i.criticalMax,0),remaining:gear.reduce((sum,i)=>sum+i.criticalMax-i.criticalDamage,0)};
}
// Manual HP edits and external HP deltas use the same combined pool as attacks.
export function setCombinedPartMax(tracker,partName,value) {
  const part=tracker.body[partName];if(!part)return;
  const inv=tracker.odyssey.inventory,goal=Math.max(0,Math.min(999,Number(value)||0));
  if(Object.hasOwn(inv.baseMax||{},partName))inv.baseMax[partName]=Math.max(0,goal-(part.max-inv.baseMax[partName]));
  part.max=goal;part.current=Math.min(part.current,part.max);reconcileEquipment(tracker);
}
export function setCombinedPartHp(tracker,partName,value) {
  const part=tracker.body[partName];if(!part)return;
  const goal=Math.max(0,Math.min(Number(value)||0,part.max)),delta=goal-part.current;
  if(delta<0)absorbEquipmentCritical(tracker,partName,-delta);
  if(delta>0) {
    const pool=armorHealth(tracker,partName),bodyCurrent=Math.max(0,part.current-pool.remaining),bodyMax=Math.max(0,part.max-pool.max);
    let repair=Math.max(0,delta-Math.max(0,bodyMax-bodyCurrent));
    for(const i of tracker.odyssey.inventory.items.filter(i=>['armor','shield'].includes(i.kind)&&i.equippedPart===partName&&i.quantity>0)) {
      const heal=Math.min(repair,i.criticalDamage);i.criticalDamage-=heal;repair-=heal;
    }
    reconcileEquipment(tracker);
  }
  part.current=goal;
}
export function useItemOnTracker(target,item,partName,effectId) {
  if(item.subtype==='grenade')throw new Error('Grenades are a placeholder; no action is performed.');
  if(item.subtype==='stim'){target.odyssey.effects??=[];if(target.odyssey.effects.length>=100)throw new Error('Too many active effects.');target.odyssey.effects.push({id:effectId,name:item.name,bonuses:normalizeBonuses(item.bonuses)});return;}
  const part=target.body[partName];if(!part)throw new Error('Choose a body part.');
  if(item.subtype==='medkit') {
    const pool=armorHealth(target,partName),missing=Math.max(0,part.max-pool.max-Math.max(0,part.current-pool.remaining));
    if(!item.healCritical||!missing)throw new Error('No critical wounds to heal on this part.');
    part.current+=Math.min(missing,item.healCritical);return;
  }
  if(item.subtype==='repairkit') {
    const gear=target.odyssey.inventory.items.filter(i=>i.equippedPart===partName&&['armor','shield'].includes(i.kind)&&i.quantity>0&&i.criticalDamage>0);
    if(!gear.length)throw new Error('No damaged armor is equipped on this part.');
    if(!item.healCritical)throw new Error('Repair amount must be greater than zero.');
    let remaining=item.healCritical;for(const i of gear){const heal=Math.min(remaining,i.criticalDamage);i.criticalDamage-=heal;remaining-=heal;}
    reconcileEquipment(target);return;
  }
  throw new Error('This item cannot be used.');
}
