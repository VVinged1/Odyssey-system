const stable=value=>Array.isArray(value)?value.map(stable):value&&typeof value==='object'?Object.fromEntries(Object.keys(value).sort().map(k=>[k,stable(value[k])])):value;
export const itemNameKey=item=>String(item.name||'').trim().toLocaleLowerCase();
export function groupNamedItems(rows,getValue=row=>row){
 const groups=new Map();for(const row of rows){const v=getValue(row),key=(v.kind||'')+':'+itemNameKey(v);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(row);}return [...groups.values()];
}
export function mergeInventoryStacks(items,preferredId='',skillLabels={}){
 const result=[],known=new Map();for(const item of items){
  if(item.equippedPart){result.push(item);continue;}
  const {id,quantity,equippedPart,previousArmor,legacyEquipment,...definition}=item;
  if(!['consumable','medical'].includes(item.kind))for(const k of ['subtype','healCritical','useSkill'])delete definition[k];
  if(!['armor','shield','implant'].includes(item.kind))for(const k of ['armor','criticalMax','criticalDamage','slots','slot','playerRemovable'])delete definition[k];
  if(item.kind!=='implant')delete definition.abilities;else definition.abilities=(definition.abilities||[]).map(({id,...ability})=>ability);
  if(definition.bonuses){definition.bonuses=structuredClone(definition.bonuses);for(const field of ['skills','skillCategories'])if(definition.bonuses[field])definition.bonuses[field]=Object.entries(definition.bonuses[field]).reduce((result,[key,value])=>{const label=String(skillLabels[key]||key).trim().toLocaleLowerCase();result[label]=field==='skills'?(result[label]||0)+value:value;return result;},{});}
  const key=JSON.stringify(stable({...definition,name:itemNameKey(item)})),previous=known.get(key);
  if(previous&&previous.quantity+quantity<=999999){if(id===preferredId){item.quantity+=previous.quantity;result[result.indexOf(previous)]=item;known.set(key,item);}else previous.quantity+=quantity;}
  else{result.push(item);known.set(key,item);}
 }return result;
}
