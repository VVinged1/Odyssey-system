import {ATTRIBUTES,effectiveOdyssey,useItemOnTracker} from './characterRules.js';
export function itemUseRequirement(odyssey,item) {
 const skill=String(item.useSkill||'').trim();if(!skill)return {mode:'none',skill:''};
 const match=Object.keys(odyssey.skills||{}).find(key=>key.toLowerCase()===skill.toLowerCase()||String(odyssey.skillLabels?.[key]||key).toLowerCase()===skill.toLowerCase());
 const learned=match&&(!Object.hasOwn(odyssey.grantedSkills||{},match)||match.startsWith('implant:'));
 return learned?{mode:'skill',skill:odyssey.skillLabels?.[match]||skill,key:match,level:odyssey.skills[match]}:{mode:'attribute',skill};
}
const passed=result=>['success','critical-success'].includes(result.outcome);
export function resolveItemUseCheck({odyssey,item,attribute,rollCharacter,rollSkill}) {
 const requirement=itemUseRequirement(odyssey,item),stages=[];
 if(requirement.mode==='none')return {...requirement,stages,success:true};
 if(requirement.mode==='attribute'){
  if(!ATTRIBUTES.includes(attribute))throw Error('Choose a characteristic for the unknown skill.');
  const result=rollCharacter(odyssey.attributes[attribute]||0,0);stages.push({kind:'attribute',name:attribute,result});if(!passed(result))return {...requirement,stages,success:false};
 }
 const result=rollSkill(requirement.mode==='skill'?requirement.level:0,0);stages.push({kind:requirement.mode==='skill'?'skill':'random',name:requirement.mode==='skill'?requirement.skill:'d100 vs d100',result});return {...requirement,stages,success:passed(result)};
}
// Validate and project the effect before any roll. Cancellation/invalid targets
// do not spend a unit; failed checks spend exactly one without changing target.
export function applyCheckedItemUse({source,target,itemId,selection,attribute,effectId,rollCharacter,rollSkill}) {
 const item=source.odyssey.inventory.items.find(row=>row.id===itemId);
 if(!item||!['consumable','medical'].includes(item.kind)||item.quantity<1)throw Error('No usable item available.');
 const projected=structuredClone(target);useItemOnTracker(projected,item,selection,effectId);
 const check=resolveItemUseCheck({odyssey:effectiveOdyssey(source.odyssey),item,attribute,rollCharacter,rollSkill});
 const same=source===target,nextSource=same&&check.success?projected:structuredClone(source);
 nextSource.odyssey.inventory.items.find(row=>row.id===itemId).quantity--;
 return {...check,itemName:item.name,source:nextSource,target:same?nextSource:check.success?projected:target};
}
export function itemUseSummary(result) {
 const stages=result.stages.map(stage=>stage.kind==='attribute'?stage.name+': '+stage.result.roll+' vs '+stage.result.finalAttribute+' ('+stage.result.result+')':stage.name+': '+stage.result.totalPrimary+' vs '+stage.result.totalSecondary+' ('+stage.result.result+')');
 return result.itemName+' · '+(result.success?'effect applied':'no effect')+' · 1 item spent'+(stages.length?' · '+stages.join(' → '):'');
}
