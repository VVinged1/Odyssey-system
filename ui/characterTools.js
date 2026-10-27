import {installDescriptions} from './descriptions.js';
import {createPoolMetadata} from '../poolMetadata.js';
import {captureBlueprint,blueprintText,blueprintId,validateBlueprint,poolEntries,parsePoolJSON,exportPool,copyBlueprint,ITEM_POOL_PREFIX} from '../itemPool.js';
import {showItemPool} from './itemPoolView.js';
import {itemEditor,skillEditor} from './itemEditor.js';
import {ARMOR_TRAINING} from '../characterRules.js';
import {findInventoryEntry,saveInventoryEntry,deleteInventoryEntry} from '../inventory.js';
import {reloadMagazine} from '../ammunition.js';
import { buildText, buildSceneUpload } from '@owlbear-rodeo/sdk';
import { newId, swapMagazine, fillMagazine, emptyMagazine, setMagazineRounds, equipItem, unequipItem, consumeItem } from '../inventory.js';
import { makeTemplate, instantiateTemplate, readLibrary, writeLibrary, emptyLibrary, mergeLibraries, normalizeLibrary, libraryFromScenes, LIBRARY_ITEM_KEY, LIBRARY_STORAGE_KEY } from '../profiles.js';
import { inventoryView, templatesView, esc } from './characterView.js';

// A dialog lives outside the character panel so scene refreshes cannot destroy its draft.
export function formDialog(title, fields, {danger = false, message = '', onChange} = {}) {
  const dialog = document.createElement('dialog'); dialog.className='od-dialog';
  dialog.innerHTML=`<form><h2>${esc(title)}</h2>${message?`<p>${esc(message)}</p>`:''}${fields.map(f=>`<label>${esc(f.label)}${f.options?`<select name="${esc(f.name)}">${f.options.map(([v,l])=>`<option value="${esc(v)}" ${v===f.value?'selected':''}>${esc(l)}</option>`).join('')}</select>`:f.type==='textarea'?`<textarea name="${esc(f.name)}" maxlength="2000">${esc(f.value)}</textarea>`:`<input name="${esc(f.name)}" type="${esc(f.type||'text')}" value="${esc(f.value ?? '')}" ${f.required?'required':''} ${f.type==='number'?`min="${f.min ?? 0}" max="${f.max ?? 999999}" step="1"`:'maxlength="120"'}>`}</label>`).join('')}<div class="cp-row"><button type="button" data-cancel>Cancel</button><button type="submit" class="${danger?'danger':''}">${danger?'Apply':'Save'}</button></div></form>`;
  document.body.append(dialog);
  if(onChange)dialog.querySelector('form').addEventListener('change',event=>onChange(event.target.form,event));
  return new Promise(resolve=>{
    let done=false; const finish=value=>{if(done)return;done=true;dialog.close();dialog.remove();resolve(value);};
    dialog.querySelector('[data-cancel]').onclick=()=>finish(null);
    dialog.addEventListener('cancel',e=>{e.preventDefault();finish(null);});
    dialog.querySelector('form').addEventListener('submit',e=>{e.preventDefault();finish(Object.fromEntries(new FormData(e.target)));});
    dialog.showModal();
  });
}
export function installCharacterTools(host) {
  installDescriptions();
  const { OBR, sanitize, getToken, getRole, canControl, update, refreshOverlay, render, status, download } = host;
  let library, storageError='';
  try { library=readLibrary(window.localStorage,sanitize); } catch(e) { library=emptyLibrary();storageError=e.message; }
  let poolMetadata={},catalog={entries:[],errors:0},poolDialog=null,scanTimer=null,scanPromise=null,scanFingerprint='';
  const knownBlueprintIds=new Map();
  const refreshPool=metadata=>{poolMetadata=metadata||{};catalog=poolEntries(poolMetadata);poolDialog?.refresh();render();};
  const readPool=createPoolMetadata({getMetadata:()=>OBR.room.getMetadata(),onMetadataChange:fn=>OBR.room.onMetadataChange(fn),onChange:refreshPool});
  const requirePoolGM=async()=>{if(getRole()!=='GM'||await OBR.player.getRole()!=='GM')throw new Error('Only the GM can edit the campaign item pool.');};
  const putBlueprints=async(blueprints,automatic=false)=>{
    await requirePoolGM();const metadata=await OBR.room.getMetadata(),patch={};let count=Object.entries(metadata).filter(([key,value])=>key.startsWith(ITEM_POOL_PREFIX)&&value&&!value.deleted).length;
    for(const blueprint of blueprints) {
      validateBlueprint(blueprint);const text=blueprintText(blueprint);let id=knownBlueprintIds.get(text);if(!id){id=await blueprintId(blueprint);knownBlueprintIds.set(text,id);}
      const key=ITEM_POOL_PREFIX+id;if(metadata[key]&&(automatic||!metadata[key].deleted))continue;if(Object.hasOwn(patch,key))continue;
      if((!metadata[key]||metadata[key].deleted)&&++count>2000)throw new Error('Campaign item pool is full. Export a backup and remove unused definitions.');patch[key]={blueprint};
    }
    await requirePoolGM();if(Object.keys(patch).length)await OBR.room.setMetadata(patch);await readPool();
  };
  const scanPool=()=>{
    if(getRole()!=='GM')return Promise.resolve();if(scanPromise)return scanPromise;
    scanPromise=(async()=>{
    try {
      const candidates=new Map();for(const token of host.getTokens()) {
        const t=host.tracker(token),o=t.odyssey;
        const entries=[...o.inventory.items.map(i=>['item',i.id]),...o.inventory.magazines.map(i=>['magazine',i.id]),...o.ammunition.map(i=>['ammo',i.id]),...['melee','ranged'].flatMap(kind=>o.weapons[kind].map(i=>['weapon-'+kind,i.id]))];
        for(const [type,id]of entries){const blueprint=captureBlueprint(t,type,id);candidates.set(blueprintText(blueprint),blueprint);}
      }
      const fingerprint=[...candidates.keys()].sort().join('\n');if(fingerprint===scanFingerprint)return;
      await putBlueprints([...candidates.values()],true);scanFingerprint=fingerprint;
    } catch(e){scanFingerprint='';status('Item pool: '+e.message,'error');}
    })().finally(()=>{scanPromise=null;});return scanPromise;
  };
  const schedulePoolScan=()=>{if(getRole()!=='GM'||scanTimer)return;scanTimer=setTimeout(()=>{scanTimer=null;void scanPool();},400);};
  void readPool().catch(e=>status('Item pool: '+e.message,'error'));
  let busy=false;
  const saveLibrary=next=>{if(getRole()!=='GM')throw new Error('Only the GM can edit templates.');writeLibrary(window.localStorage,next);library=next;storageError='';render();};
  const check=(id,gm=false)=>{
    const token=getToken(id);
    if(!token||!canControl(token)||(gm&&getRole()!=='GM')) throw new Error(gm?'Only the GM can perform this action.':'You cannot control this character.');
    return token;
  };
  const mutate=async(id,fn,gm=false)=>{
    check(id,gm);
    await update(id,current=>{check(id,gm);const next=structuredClone(current);fn(next);return next;});
    await refreshOverlay(id);render();
  };
  const readSelect=(root,attribute,value)=>Array.from(root.querySelectorAll(`[${attribute}]`)).find(e=>e.getAttribute(attribute)===String(value))?.value || '';
  async function handle(action,node) {
    const token=check(); const tokenId=token.id, panel=node.closest('#selectedTokenPanel'), id=node.dataset.id;
    const gmActions=['add-item','edit-item','delete-item','add-mag','fill-mag','set-mag-rounds','empty-mag','delete-mag','mount-mag','apply-template'];
    if(action==='open-item-pool')check(tokenId,true);
    if(gmActions.includes(action)||action.includes('template')||action.includes('library')||action.startsWith('account-')) check(tokenId,true);
    if(action==='add-item'||action==='edit-entry') {
      check(tokenId,true);const t=host.tracker(token),o=t.odyssey,type=node.dataset.entryType||'item';
      const existing=action==='edit-entry'?findInventoryEntry(t,type,id):null;
      if(action==='edit-entry'&&!existing)throw new Error('Item no longer exists.');
      const kind=type.startsWith('weapon-')?'weapon':type==='item'?(['item','tool'].includes(existing?.kind)?'resource':existing?.kind==='medical'?'consumable':existing?.kind||'resource'):type;
      const values=await itemEditor({existing:{...existing,...(type.startsWith('weapon-')?{weaponKind:type.slice(7)}:{})},type:kind,skills:Object.entries(o.skills).map(([k])=>[k,o.skillLabels[k]||k]),parts:Object.entries(t.body).map(([k,p])=>[k,p.label||k]),ammunition:o.ammunition,magazines:o.inventory.magazines});if(!values)return;
      const entryId=existing?.id||newId(),entryType=values.kind==='weapon'?'weapon-'+values.weaponKind:['ammo','magazine'].includes(values.kind)?values.kind:'item';let blueprint;
      await mutate(tokenId,t=>{saveInventoryEntry(t,existing?{type,id}:null,values,entryId);blueprint=captureBlueprint(sanitize(t),entryType,entryId);},true);
      try{await putBlueprints([blueprint]);}catch(e){throw new Error('Item saved on character; campaign pool was not updated: '+e.message);}
    } else if(action==='open-item-pool') {
      await requirePoolGM();await readPool();await scanPool();await readPool();
      poolDialog=showItemPool({getCatalog:()=>catalog,character:token.name||'Character',
        onGive:async entryId=>{
          const entry=catalog.entries.find(e=>e.id===entryId);if(!entry)throw new Error('Item definition no longer exists.');
          const value=await formDialog('Give '+entry.blueprint.values.name,[{name:'quantity',label:'Quantity',type:'number',min:1,max:entry.blueprint.type==='ammo'?999999:entry.blueprint.type.startsWith('weapon-')?20:1000,value:1}]);if(!value)return false;
          const copyId=newId();await mutate(tokenId,t=>copyBlueprint(t,entry.blueprint,Number(value.quantity),copyId),true);return true;
        },
        onDelete:async entryId=>{if(!await formDialog('Remove from campaign pool?',[],{danger:true,message:'Existing items on characters are kept.'}))return false;await requirePoolGM();await OBR.room.setMetadata({[ITEM_POOL_PREFIX+entryId]:{deleted:true}});await readPool();return true;},
        onExport:async()=>{await requirePoolGM();download('odyssey-item-pool.json',JSON.stringify(exportPool(catalog.entries),null,2));},
        onImport:()=>new Promise((resolve,reject)=>{
          const input=document.createElement('input');input.type='file';input.accept='.json,application/json';input.addEventListener('cancel',()=>resolve(false));
          input.onchange=async()=>{try{const file=input.files?.[0];if(!file){resolve(false);return;}if(file.size>4*1024*1024)throw new Error('Item pool JSON must be under 4 MB.');await requirePoolGM();const blueprints=parsePoolJSON(JSON.parse(await file.text())).map(raw=>{const t=sanitize({}),id=newId();copyBlueprint(t,raw,1,id);return captureBlueprint(sanitize(t),raw.type,id);});await putBlueprints(blueprints);resolve(true);}catch(e){reject(e);}};input.click();
        }),
      });await poolDialog.closed;poolDialog=null;
    } else if(action==='delete-entry') {
      check(tokenId,true);if(!await formDialog('Remove item?',[],{danger:true,message:'Remove this item from this token?'}))return;
      await mutate(tokenId,t=>deleteInventoryEntry(t,node.dataset.entryType,id),true);
    } else if(action==='add-skill'||action==='edit-skill') {
      check(tokenId,true);const key=node.dataset.skill,o=host.tracker(token).odyssey;
      const value=await skillEditor(o,key);if(!value)return;
      const skillKey=key||newId();await mutate(tokenId,t=>{
        const d=t.odyssey;if(key&&!Object.hasOwn(d.skills,key))throw new Error('Skill no longer exists.');
        if(value.delete){if(['Melee','Parry'].includes(key))throw new Error('Core combat skill cannot be removed.');for(const field of ['skills','skillLabels','skillCategories','skillStrengthBonuses','skillPerks','skillDescriptions','skillPerkDescriptions','skillRangeReduction'])delete d[field][key];return;}
        d.skills[skillKey]=value.level;d.skillLabels[skillKey]=value.name;d.skillCategories[skillKey]=value.category;d.skillStrengthBonuses[skillKey]=value.strength;d.skillPerks[skillKey]=value.perks;d.skillDescriptions[skillKey]=value.description;d.skillPerkDescriptions[skillKey]=value.perkDescriptions;d.skillRangeReduction??={};d.skillRangeReduction[skillKey]=value.rangeReduction;
      },true);
    } else if(action==='edit-training') {
      check(tokenId,true);const o=host.tracker(token).odyssey;
      const value=await formDialog('Armor training',Object.entries(ARMOR_TRAINING).map(([key,label])=>({name:key,label,type:'number',min:0,max:1,value:o.skills[key]||0})));if(!value)return;
      await mutate(tokenId,t=>{for(const key of Object.keys(ARMOR_TRAINING))t.odyssey.skills[key]=Number(value[key])>=1?1:0;},true);
    } else if(action==='remove-effect') {
      await mutate(tokenId,t=>{t.odyssey.effects=t.odyssey.effects.filter(e=>e.id!==id);},true);
    } else if(action==='use-item') {
      const item=host.tracker(token).odyssey.inventory.items.find(i=>i.id===id);if(!item)throw new Error('Item no longer exists.');
      const targets=host.getTokens().filter(t=>getRole()==='GM'||t.visible!==false);
      const partsFor=target=>Object.entries(host.tracker(target).body).filter(([,p])=>getRole()==='GM'||!p.hidden).map(([k,p])=>[k,p.label||k]);
      const fields=[{name:'target',label:'Target',value:'',options:[['','Self'],...targets.filter(t=>t.id!==tokenId).map(t=>[t.id,t.name||t.id])]}];
      const repairOptions=target=>host.tracker(target).odyssey.inventory.items.filter(i=>['armor','shield'].includes(i.kind)&&i.quantity>0&&i.criticalDamage>0&&(getRole()==='GM'||!i.equippedPart||!host.tracker(target).body[i.equippedPart]?.hidden)).map(i=>[i.id,i.name+' · durability '+(i.criticalMax-i.criticalDamage)+'/'+i.criticalMax+(i.equippedPart?' · equipped':' · inventory')]);
      if(item.subtype==='repairkit')fields.push({name:'gear',label:'Armor in inventory',options:[['','Choose damaged armor'],...repairOptions(token)]});
      else if(item.subtype!=='stim')fields.push({name:'part',label:'Body part',options:partsFor(token)});
      const value=await formDialog('Use '+item.name,fields,{onChange:(form,event)=>{
        if(event.target.name!=='target')return;
        if(form.elements.gear){const target=host.getToken(form.elements.target.value||tokenId);form.elements.gear.innerHTML='<option value="">Choose damaged armor</option>'+(target?repairOptions(target).map(([key,label])=>'<option value="'+esc(key)+'">'+esc(label)+'</option>').join(''):'');return;}
        if(!form.elements.part)return;
        const target=host.getToken(form.elements.target.value||tokenId);form.elements.part.innerHTML=target?partsFor(target).map(([key,label])=>'<option value="'+esc(key)+'">'+esc(label)+'</option>').join(''):'';
      }});if(!value)return;
      check(tokenId);const target=host.getToken(value.target||tokenId);if(!target)throw new Error('Target no longer exists.');
      await host.useItem(tokenId,id,target.id,item.subtype==='repairkit'?value.gear||'':value.part||'');
    } else if(action==='reload-weapon') {
      const index=Number(node.dataset.index),ammoId=readSelect(panel,'data-reload-kind',index);await mutate(tokenId,t=>reloadMagazine(t.odyssey,index,ammoId),true);
    } else if(action==='delete-item') {
      const yes=await formDialog('Remove item?',[],{danger:true,message:'This removes the item from this token.'});if(!yes)return;
      await mutate(tokenId,t=>{unequipItem(t,id);t.odyssey.inventory.items=t.odyssey.inventory.items.filter(i=>i.id!==id);},true);
    } else if(action==='consume') await mutate(tokenId,t=>consumeItem(t,id));
    else if(action==='equip') {const part=readSelect(panel,'data-equip-part',id);await mutate(tokenId,t=>equipItem(t,id,part,getRole()==='GM'));}
    else if(action==='unequip') await mutate(tokenId,t=>unequipItem(t,id,getRole()==='GM'));
    else if(action==='add-mag') {
      const value=await formDialog('New magazine',[{name:'name',label:'Name',required:true,value:'Magazine'},{name:'capacity',label:'Capacity',type:'number',min:1,max:999,value:6}]);if(!value)return;
      const magazine={id:newId(),name:value.name,capacity:Number(value.capacity),loaded:0,loadedAmmoId:''};
      await mutate(tokenId,t=>{if(t.odyssey.inventory.magazines.length>=200)throw new Error('Magazine limit is 200.');t.odyssey.inventory.magazines.push(magazine);},true);
    } else if(action==='set-mag-rounds') {
      const tracker=host.tracker(check(tokenId,true)), o=tracker.odyssey;
      const mag=o.inventory.magazines.find(m=>m.id===id);
      if(!mag)throw new Error('Magazine no longer exists.');
      const weapon=o.weapons.ranged.find(w=>w.magazineId===id);
      const ammo=o.ammunition.filter(a=>!weapon||Array.isArray(weapon.acceptedMagazineIds)||weapon.ammoIds.includes(a.id));
      const value=await formDialog('GM · Set magazine rounds',[
        {name:'rounds',label: 'Rounds (0–'+mag.capacity+')',type:'number',min:0,max:mag.capacity,value:mag.loaded,required:true},
        {name:'ammoId',label:'Ammunition type',value:mag.loadedAmmoId,options:[['','Empty / no ammunition'],...ammo.map(a=>[a.id,a.name])]},
      ],{message:'Direct GM correction. Ammunition reserve is not changed; this is not a reload.'});
      if(!value)return;
      await mutate(tokenId,t=>setMagazineRounds(t.odyssey,id,Number(value.rounds),value.ammoId),true);
    } else if(action==='fill-mag') {const ammoId=readSelect(panel,'data-mag-ammo',id);await mutate(tokenId,t=>fillMagazine(t.odyssey,id,ammoId),true);}
    else if(action==='empty-mag') await mutate(tokenId,t=>emptyMagazine(t.odyssey,id),true);
    else if(action==='delete-mag') await mutate(tokenId,t=>{emptyMagazine(t.odyssey,id);t.odyssey.inventory.magazines=t.odyssey.inventory.magazines.filter(m=>m.id!==id);},true);
    else if(action==='mount-mag') {const index=Number(node.dataset.index),magId=readSelect(panel,'data-weapon-mag',index);await mutate(tokenId,t=>swapMagazine(t.odyssey,index,magId),true);}
    else if(action==='save-template') {
      const value=await formDialog('Save character template',[{name:'name',label:'Template name',required:true,value:token.name||'Character'}],{message:'Saves a healthy character. Owner, target, history and bridge events are excluded.'});if(!value)return;
      if(storageError)throw new Error('Export the unreadable local library before importing a replacement. '+storageError);
      const current=check(tokenId);const template=makeTemplate(value.name,host.tracker(current));saveLibrary(mergeLibraries(library,{...emptyLibrary(),templates:[template]}));
    } else if(action==='export-library') {
      const json=storageError?window.localStorage.getItem(LIBRARY_STORAGE_KEY):JSON.stringify(library,null,2);
      download('odyssey-templates.json',json||'{}');return;
    } else if(action==='import-library') {
      const input=document.createElement('input');input.type='file';input.accept='.json,application/json';
      input.onchange=async()=>{try {const file=input.files?.[0];if(!file)return;if(file.size>4*1024*1024)throw new Error('Library must be under 4 MB.');const incoming=normalizeLibrary(JSON.parse(await file.text()),sanitize);if(storageError){const yes=await formDialog('Replace unreadable library?',[],{danger:true,message:'Export your existing JSON first. This replaces the unreadable local copy.'});if(!yes)return;saveLibrary(incoming);}else saveLibrary(mergeLibraries(library,incoming));status('Template library imported.','success');}catch(e){status(e.message,'error');}};
      input.click();return;
    } else if(action==='apply-template') {
      const template=library.templates.find(t=>t.id===id);if(!template)throw new Error('Template no longer exists.');
      const yes=await formDialog(`Apply ${template.name}?`,[],{danger:true,message:'Replaces body, attributes, skills, weapons and inventory on this token. Destination owner and history are preserved. Saved attack target is cleared.'});if(!yes)return;
      const source=check(tokenId,true), next=instantiateTemplate(host.tracker(source),template,sanitize);
      await mutate(tokenId,t=>{t.enabled=next.enabled;t.minor=next.minor;t.serious=next.serious;t.body=next.body;t.odyssey={...next.odyssey,owner:t.odyssey.owner};},true);
      host.resetDraft?.(tokenId);
    } else if(action==='rename-template'||action==='duplicate-template') {
      const t=library.templates.find(t=>t.id===id);if(!t)throw new Error('Template no longer exists.');
      const value=await formDialog(action==='rename-template'?'Rename template':'Duplicate template',[{name:'name',label:'Name',required:true,value:t.name+(action==='duplicate-template'?' copy':'')}]);if(!value)return;
      const copy={...structuredClone(t),name:value.name.trim(),updatedAt:new Date().toISOString()};
      const next=structuredClone(library);if(action==='duplicate-template'){copy.id=newId();next.templates.push(copy);}else next.templates=next.templates.map(x=>x.id===id?copy:x);
      saveLibrary(normalizeLibrary(next,sanitize));
    } else if(action==='delete-template') {
      const yes=await formDialog('Remove template?',[],{danger:true,message:'Existing characters are unaffected.'});if(!yes)return;
      saveLibrary({...library,templates:library.templates.filter(t=>t.id!==id)});
    } else if(action==='account-save') {
      if(!library.templates.length)throw new Error('Save at least one template first.');
      const items=library.templates.map(t=>buildText().name(t.name).plainText(t.name).position({x:0,y:0}).metadata({[LIBRARY_ITEM_KEY]:{...emptyLibrary(),templates:[t]}}).build());
      await OBR.assets.uploadScenes([buildSceneUpload().name('Odyssey Templates').items(items).build()],true);
      status('Assets folder picker closed. Verify the library snapshot in your Owlbear account.','info');return;
    } else if(action==='account-load') {
      const scenes=await OBR.assets.downloadScenes(true,'Odyssey Templates');if(!scenes?.length)return;
      const incoming=libraryFromScenes(scenes,sanitize);
      if(storageError){const yes=await formDialog('Replace unreadable library?',[],{danger:true,message:'Export your existing JSON before replacing it with the account snapshot.'});if(!yes)return;saveLibrary(incoming);}else saveLibrary(mergeLibraries(library,incoming));
    }
    status('Saved.','success');
  }
  document.addEventListener('click',event=>{
    const node=event.target.closest?.('[data-od-action]');if(!node||node.disabled||busy)return;
    busy=true;document.querySelectorAll('[data-od-action]').forEach(e=>{if(e instanceof HTMLButtonElement)e.disabled=true;});
    handle(node.dataset.odAction,node).catch(e=>status(e.message,'error')).finally(()=>{busy=false;render();if(node.isConnected)node.disabled=false;});
  });
  const reload=()=>{try {library=readLibrary(window.localStorage,sanitize);storageError='';render();}catch(e){storageError=e.message;render();}};
  window.addEventListener('storage',event=>{if(event.key===LIBRARY_STORAGE_KEY)reload();});
  const disabledWhileBusy=html=>busy?html.replaceAll('<button ', '<button disabled ').replaceAll('<select ', '<select disabled ').replaceAll('<input ', '<input disabled '):html;
  return {syncPool:schedulePoolScan,isBusy:()=>busy,renderInventory:tracker=>disabledWhileBusy(inventoryView(tracker,getRole()==='GM')),renderTemplates:()=>getRole()!=='GM'?'':disabledWhileBusy(`${storageError?`<p class="cp-banner err">Local library could not be read: ${esc(storageError)}. Export JSON before replacing it.</p>`:''}${templatesView(library,getRole()==='GM')}`)};
}
