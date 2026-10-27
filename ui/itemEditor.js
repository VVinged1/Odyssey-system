import { esc } from './characterView.js';
import { ATTRIBUTES } from '../characterRules.js';
export const ITEM_TYPES=[['consumable','Usable'],['resource','Resource'],['ammo','Ammunition'],['magazine','Magazine'],['weapon','Weapon'],['armor','Armor'],['shield','Shield'],['implant','Implant']];
export function itemEditor({existing={},type='resource',skills=[],parts=[],ammunition=[],magazines=[]}) {
  const dialog=document.createElement('dialog');dialog.className='od-dialog od-item-editor';
  const field=(name,label,value='',inputType='text',max=999)=>`<label>${esc(label)}<input name="${name}" type="${inputType}" value="${esc(value)}" ${inputType==='number'?`min="${['damage','accuracy','penetration'].includes(name)?-999:0}" max="${max}" step="1"`:'maxlength="120"'}></label>`;
  const select=(name,label,options,value)=>`<label>${esc(label)}<select name="${name}">${options.map(([k,l])=>`<option value="${esc(k)}" ${k===value?'selected':''}>${esc(l)}</option>`).join('')}</select></label>`;
  const checkbox=(name,label,checked)=>`<label class="od-check"><input name="${name}" type="checkbox" ${checked?'checked':''}>${esc(label)}</label>`;
  const values={...existing,kind:type,acceptsMagazines:existing.acceptsMagazines!==false};
  values.slots??=[existing.slot||'any'];
  values.acceptedMagazineIds??=existing.id?magazines.filter(m=>m.id===existing.magazineId||(m.capacity===existing.capacity&&(!m.loaded||(existing.ammoIds||[]).includes(m.loadedAmmoId)))).map(m=>m.id):[];
  const skillChoices=[...skills];
  for(const key of Object.keys(existing.bonuses?.skills||{}))if(!skillChoices.some(([k])=>k===key))skillChoices.push([key,key]);
  dialog.innerHTML=`<form><h2>${existing.id?'Edit item':'New item'}</h2>${select('kind','Type',ITEM_TYPES,type)}${field('name','Name',existing.name||'')}${field('quantity','Quantity / ammo reserve',existing.quantity??1,'number',999999)}<div data-type-fields></div><label>Description<textarea name="description" maxlength="2000">${esc(existing.description||'')}</textarea></label><div class="cp-row"><button type="button" data-cancel>Cancel</button><button type="submit">Save</button></div></form>`;
  const form=dialog.querySelector('form'),host=dialog.querySelector('[data-type-fields]');
  function bonusEditor(bonuses={},allowSkills=true) {
    return `<fieldset class="od-bonuses"><legend>Passive bonuses</legend>${ATTRIBUTES.map(k=>field('attr:'+k,k,bonuses.attributes?.[k]||0,'number')).join('')}${allowSkills?skillChoices.map(([k,label])=>field('skill:'+k,`${label} · skill level`,bonuses.skills?.[k]||0,'number')).join('')+`<div class="od-new-bonus-skill">${field('newBonusSkillName','New skill name')}${field('newBonusSkillValue','Level bonus',1,'number',10)}<button type="button" data-add-bonus-skill>Add skill bonus</button></div>`:''}</fieldset>`;
  }
  function abilityRow(a={}) {return `<div class="od-ability-row">${field('abilityName','Ability',a.name||'')}${field('abilityLevel','Level',a.level||0,'number',10)}${field('abilityDescription','Description',a.description||'')}<button type="button" data-remove-ability aria-label="Remove ability">×</button></div>`;}
  function collect() {
    const data=Object.fromEntries(new FormData(form));Object.assign(values,data);
    if(form.elements.acceptsMagazines)values.acceptsMagazines=form.elements.acceptsMagazines.checked;
    if(form.querySelector('[name="magazineCompatible"]'))values.acceptedMagazineIds=[...form.querySelectorAll('[name="magazineCompatible"]:checked')].map(e=>e.value);
    if(form.querySelector('[data-equipment-slots]'))values.slots=[...form.querySelectorAll('[name="equipmentSlot"]:checked')].map(e=>e.value);
    values.playerRemovable=!!form.elements.playerRemovable?.checked;
    const bonuses={attributes:{},skills:{}};
    for(const input of form.querySelectorAll('[name^="attr:"],[name^="skill:"]')) {
      const [kind,...key]=input.name.split(':');const value=Number(input.value)||0;if(value)bonuses[kind==='attr'?'attributes':'skills'][key.join(':')]=value;
    }
    if(form.querySelector('.od-bonuses'))values.bonuses=bonuses;
    if(form.querySelector('[data-abilities]'))values.abilities=[...form.querySelectorAll('.od-ability-row')].map((row,index)=>({id:existing.abilities?.[index]?.id||crypto.randomUUID(),name:row.querySelector('[name="abilityName"]').value,level:Number(row.querySelector('[name="abilityLevel"]').value)||0,description:row.querySelector('[name="abilityDescription"]').value})).filter(a=>a.name.trim());
    values.ammoIds=[...form.querySelectorAll('[name="ammoCompatible"]:checked')].map(e=>e.value);
  }
  function render() {
    const kind=form.elements.kind.value;values.kind=kind;
    let html='';
    if(kind==='consumable') {
      html+=select('subtype','Use type',[['grenade','Grenade · placeholder'],['medkit','Medkit'],['repairkit','Repair kit'],['stim','Stim']],values.subtype||'grenade');
      if(['medkit','repairkit'].includes(values.subtype))html+=field('healCritical','Critical wounds restored',values.healCritical??1,'number');
      if(values.subtype==='stim')html+=bonusEditor(values.bonuses,false);
      if(!values.subtype||values.subtype==='grenade')html+='<p class="cp-muted">Grenades have no action yet.</p>';
    }
    if(kind==='resource')html+=bonusEditor(values.bonuses);
    if(kind==='ammo')html+=field('damage','Damage / round',values.damage??0,'number')+field('penetration','Armor penetration',values.penetration??0,'number');
    if(kind==='magazine')html+=field('capacity','Capacity',values.capacity??6,'number')+field('loaded','Rounds · GM correction',values.loaded??0,'number')+select('loadedAmmoId','Ammunition', [['','Empty'],...ammunition.map(a=>[a.id,a.name])],values.loadedAmmoId||'');
    if(kind==='weapon') {
      html+=select('weaponKind','Weapon type',[['melee','Melee'],['ranged','Ranged']],values.weaponKind||'melee')+field('accuracy','Accuracy',values.accuracy??0,'number');
      if(values.weaponKind==='ranged') {
        html+=checkbox('acceptsMagazines','Accepts detachable magazines',values.acceptsMagazines);
        if(values.acceptsMagazines)html+='<fieldset><legend>Compatible magazines</legend>'+magazines.map(m=>checkbox('magazineCompatible',m.name+' · '+m.capacity+' rounds',values.acceptedMagazineIds.includes(m.id)).replace('name="magazineCompatible"','name="magazineCompatible" value="'+esc(m.id)+'"')).join('')+(magazines.length?'':'<p class="cp-muted">Create a magazine in Items first.</p>')+'</fieldset>';
        else html+=field('capacity','Built-in magazine capacity',values.capacity||6,'number')+field('loaded','Loaded rounds · GM correction',values.loaded||0,'number')+select('loadedAmmoId','Loaded ammunition',[['','Empty'],...ammunition.map(a=>[a.id,a.name])],values.loadedAmmoId||'')+'<fieldset><legend>Compatible ammunition</legend>'+ammunition.map(a=>checkbox('ammoCompatible',a.name,(values.ammoIds||[]).includes(a.id)).replace('name="ammoCompatible"','name="ammoCompatible" value="'+esc(a.id)+'"')).join('')+'</fieldset>';
      }
      else html+=field('damage','Damage',values.damage??0,'number');
    }
    if(['armor','shield','implant'].includes(kind)) {
      html+=field('armor','Armor',values.armor??0,'number')+field('criticalMax',kind==='implant'?'Body part HP bonus':'Critical durability',values.criticalMax??1,'number')+(kind==='implant'?'':field('criticalDamage','Critical damage taken',values.criticalDamage??0,'number'))+'<fieldset data-equipment-slots class="od-equipment-slots"><legend>Equipment slots · choose allowed parts</legend>'+[['any','Any part'],...parts,...values.slots.filter(key=>key!=='any'&&!parts.some(([k])=>k===key)).map(key=>[key,key])].map(([key,label])=>checkbox('equipmentSlot',label,values.slots.includes(key)).replace('name="equipmentSlot"','name="equipmentSlot" value="'+esc(key)+'"')).join('')+'</fieldset>';
      if(kind==='implant')html+='<p class="cp-muted">Adds maximum body-part HP. Equipping does not heal wounds.</p>';
      if(kind==='implant')html+=checkbox('playerRemovable','Player may equip / remove this implant',values.playerRemovable)+bonusEditor(values.bonuses)+`<fieldset><legend>Granted abilities · available only while equipped</legend><div data-abilities>${(values.abilities||[]).map(abilityRow).join('')}</div><button type="button" data-add-ability>+ Ability</button></fieldset>`;
    }
    host.innerHTML=html;
    // Bonus inputs accept signed modifiers.
    for(const input of host.querySelectorAll('[name^="attr:"],[name^="skill:"]'))input.min='-999';
    form.elements.quantity.closest('label').hidden=['magazine','weapon'].includes(kind);
  }
  form.addEventListener('change',e=>{if(['kind','subtype','weaponKind','acceptsMagazines'].includes(e.target.name)){collect();render();}});
  form.addEventListener('click',e=>{if(e.target.closest('[data-add-bonus-skill]')) {
    const input=form.elements.newBonusSkillName,name=input.value.trim(),bonus=Number(form.elements.newBonusSkillValue.value)||0;
    if(!name||name.length>40||['__proto__','constructor','prototype'].includes(name)){input.setCustomValidity('Enter a skill name (1–40 characters).');input.reportValidity();input.oninput=()=>input.setCustomValidity('');return;}
    collect();let choice=skillChoices.find(([key,label])=>key.toLowerCase()===name.toLowerCase()||label.toLowerCase()===name.toLowerCase());
    if(!choice){choice=[name,name];skillChoices.push(choice);}values.bonuses.skills[choice[0]]=bonus;render();return;
  }if(e.target.closest('[data-add-ability]')){host.querySelector('[data-abilities]').insertAdjacentHTML('beforeend',abilityRow());}e.target.closest('[data-remove-ability]')?.closest('.od-ability-row').remove();});
  document.body.append(dialog);render();
  return new Promise(resolve=>{
    const finish=v=>{dialog.close();dialog.remove();resolve(v);};
    form.querySelector('[data-cancel]').onclick=()=>finish(null);dialog.addEventListener('cancel',e=>{e.preventDefault();finish(null);});
    form.addEventListener('submit',e=>{e.preventDefault();collect();if(['armor','shield','implant'].includes(values.kind)&&!values.slots.length){host.querySelector('[data-equipment-slots]').insertAdjacentHTML('beforeend','<p class="cp-muted">Choose at least one allowed slot.</p>');return;}if(!values.name.trim()){form.elements.name.setCustomValidity('Enter a name.');form.elements.name.reportValidity();form.elements.name.oninput=()=>form.elements.name.setCustomValidity('');return;}finish(values);});dialog.showModal();
  });
}
export function skillEditor(o,key) {
  const existing=key&&Object.hasOwn(o.skills,key),core=['Melee','Parry'].includes(key);
  const dialog=document.createElement('dialog');dialog.className='od-dialog';
  dialog.innerHTML=`<form><h2>${existing?'Edit skill':'New skill'}</h2><label>Name<input name="name" required maxlength="40" value="${esc(existing?o.skillLabels[key]||key:'')}"></label><label>Base level<input name="level" type="number" min="0" max="10" value="${existing?o.skills[key]:0}"></label><label>Category<select name="category">${['combat','abilities','applied'].map(c=>`<option ${o.skillCategories[key]===c?'selected':''}>${c}</option>`).join('')}</select></label><label class="od-check"><input name="strength" type="checkbox" ${o.skillStrengthBonuses[key]?'checked':''}>Strength bonus for attack damage</label><fieldset><legend>Perks</legend><div data-perks></div><button type="button" data-add-perk>+ Perk</button></fieldset><div class="cp-row"><button type="button" data-cancel>Cancel</button>${existing&&!core?'<button type="button" class="danger" data-delete>Delete skill</button>':''}<button type="submit">Save</button></div></form>`;
  const list=dialog.querySelector('[data-perks]');
  const add=value=>list.insertAdjacentHTML('beforeend',`<div class="cp-row"><input name="perk" maxlength="40" value="${esc(value)}" aria-label="Perk"><button type="button" data-remove-perk aria-label="Remove perk">×</button></div>`);
  for(const p of o.skillPerks[key]||[])add(p);
  document.body.append(dialog);
  return new Promise(resolve=>{const finish=v=>{dialog.close();dialog.remove();resolve(v);};dialog.querySelector('[data-cancel]').onclick=()=>finish(null);dialog.querySelector('[data-delete]')?.addEventListener('click',()=>finish({delete:true}));dialog.addEventListener('cancel',e=>{e.preventDefault();finish(null);});dialog.addEventListener('click',e=>{if(e.target.closest('[data-add-perk]')&&list.children.length<12)add('');e.target.closest('[data-remove-perk]')?.parentElement.remove();});dialog.querySelector('form').onsubmit=e=>{e.preventDefault();const f=e.target;finish({name:f.elements.name.value.trim(),level:Number(f.elements.level.value),category:f.elements.category.value,strength:f.elements.strength.checked,perks:[...list.querySelectorAll('input')].map(e=>e.value.trim()).filter(Boolean)});};dialog.showModal();});
}
