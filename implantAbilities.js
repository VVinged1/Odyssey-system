import {normalizeBonuses} from './characterRules.js';
export function activeImplantAbilities(o){return (o.inventory?.items||[]).filter(i=>i.kind==='implant'&&i.equippedPart&&i.quantity>0).flatMap(item=>(item.abilities||[]).filter(a=>['attack','bonus'].includes(a.kind)).map(ability=>({item,ability,key:'implant:'+item.id+':'+ability.id})));}
export function implantAttacks(o){return activeImplantAbilities(o).filter(({ability})=>ability.kind==='attack').map(({item,ability,key})=>({id:key,selectionKey:key,name:item.name+' · '+ability.name,damage:ability.damage,accuracy:ability.accuracy||0,boundSkill:ability.skill,implantId:item.id,abilityId:ability.id}));}
export function resolveAbilitySkill(o,label){return Object.keys(o.skills).find(k=>k.toLocaleLowerCase()===String(label).trim().toLocaleLowerCase()||String(o.skillLabels[k]||k).toLocaleLowerCase()===String(label).trim().toLocaleLowerCase())||'';}
export function activateImplantAbility(tracker,itemId,abilityId){
 const row=activeImplantAbilities(tracker.odyssey).find(r=>r.item.id===itemId&&r.ability.id===abilityId&&r.ability.kind==='bonus');if(!row)throw Error('Equipped implant ability is unavailable.');
 const effect={id:'implant-effect:'+itemId+':'+abilityId,name:row.item.name+' · '+row.ability.name,sourceItemId:itemId,bonuses:normalizeBonuses(row.ability.bonuses)};
 tracker.odyssey.effects??=[];if(tracker.odyssey.effects.some(e=>e.id===effect.id))throw Error('Ability is already active.');if(tracker.odyssey.effects.length>=100)throw Error('Too many active effects.');tracker.odyssey.effects.push(effect);return effect;
}
