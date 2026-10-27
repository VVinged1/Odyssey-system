import { buildText, buildSceneUpload } from '@owlbear-rodeo/sdk';
import { newId, swapMagazine, fillMagazine, emptyMagazine, equipItem, unequipItem, consumeItem } from '../inventory.js';
import { makeTemplate, instantiateTemplate, readLibrary, writeLibrary, emptyLibrary, mergeLibraries, normalizeLibrary, libraryFromScenes, LIBRARY_ITEM_KEY, LIBRARY_STORAGE_KEY } from '../profiles.js';
import { inventoryView, templatesView, esc } from './characterView.js';

// A dialog lives outside the character panel so scene refreshes cannot destroy its draft.
export function formDialog(title, fields, {danger = false, message = ''} = {}) {
  const dialog = document.createElement('dialog'); dialog.className='od-dialog';
  dialog.innerHTML=`<form><h2>${esc(title)}</h2>${message?`<p>${esc(message)}</p>`:''}${fields.map(f=>`<label>${esc(f.label)}${f.options?`<select name="${esc(f.name)}">${f.options.map(([v,l])=>`<option value="${esc(v)}" ${v===f.value?'selected':''}>${esc(l)}</option>`).join('')}</select>`:f.type==='textarea'?`<textarea name="${esc(f.name)}" maxlength="2000">${esc(f.value)}</textarea>`:`<input name="${esc(f.name)}" type="${esc(f.type||'text')}" value="${esc(f.value ?? '')}" ${f.required?'required':''} ${f.type==='number'?`min="${f.min ?? 0}" max="${f.max ?? 999999}" step="1"`:'maxlength="120"'}>`}</label>`).join('')}<div class="cp-row"><button type="button" data-cancel>Cancel</button><button type="submit" class="${danger?'danger':''}">${danger?'Apply':'Save'}</button></div></form>`;
  document.body.append(dialog);
  return new Promise(resolve=>{
    let done=false; const finish=value=>{if(done)return;done=true;dialog.close();dialog.remove();resolve(value);};
    dialog.querySelector('[data-cancel]').onclick=()=>finish(null);
    dialog.addEventListener('cancel',e=>{e.preventDefault();finish(null);});
    dialog.querySelector('form').addEventListener('submit',e=>{e.preventDefault();finish(Object.fromEntries(new FormData(e.target)));});
    dialog.showModal();
  });
}
export function installCharacterTools(host) {
  const { OBR, sanitize, getToken, getRole, canControl, update, refreshOverlay, render, status, download } = host;
  let library, storageError='';
  try { library=readLibrary(window.localStorage,sanitize); } catch(e) { library=emptyLibrary();storageError=e.message; }
  let busy=false;
  const saveLibrary=next=>{writeLibrary(window.localStorage,next);library=next;storageError='';render();};
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
    const gmActions=['add-item','edit-item','delete-item','add-mag','fill-mag','empty-mag','delete-mag','mount-mag','apply-template'];
    if(gmActions.includes(action)) check(tokenId,true);
    if(action==='add-item'||action==='edit-item') {
      const existing=host.tracker(token).odyssey.inventory.items.find(i=>i.id===id);
      if(action==='edit-item'&&!existing) throw new Error('Item no longer exists.');
      const value=await formDialog(existing?'Edit item':'Add item / armor',[
        {name:'name',label:'Name',required:true,value:existing?.name},
        {name:'kind',label:'Type',value:existing?.kind||'item',options:['item','consumable','medical','tool','resource','armor','shield','implant'].map(k=>[k,k])},
        {name:'quantity',label:'Quantity',type:'number',value:existing?.quantity??1},
        {name:'armor',label:'Armor value (equipment)',type:'number',max:999,value:existing?.armor??0},
        {name:'description',label:'Description',type:'textarea',value:existing?.description},
      ]);if(!value)return;
      const itemId=existing?.id||newId();
      await mutate(tokenId,t=>{
        const rows=t.odyssey.inventory.items, i=rows.find(i=>i.id===itemId);
        if(i?.equippedPart)throw new Error('Unequip the item before editing it.');
        if(!i&&rows.length>=200)throw new Error('Inventory limit is 200 items.');
        const item={id:itemId,...value,quantity:Number(value.quantity),armor:Number(value.armor),equippedPart:'',previousArmor:0};
        if(i)Object.assign(i,item);else rows.push(item);
      },true);
    } else if(action==='delete-item') {
      const yes=await formDialog('Remove item?',[],{danger:true,message:'This removes the item from this token.'});if(!yes)return;
      await mutate(tokenId,t=>{unequipItem(t,id);t.odyssey.inventory.items=t.odyssey.inventory.items.filter(i=>i.id!==id);},true);
    } else if(action==='consume') await mutate(tokenId,t=>consumeItem(t,id));
    else if(action==='equip') {const part=readSelect(panel,'data-equip-part',id);await mutate(tokenId,t=>equipItem(t,id,part));}
    else if(action==='unequip') await mutate(tokenId,t=>unequipItem(t,id));
    else if(action==='add-mag') {
      const value=await formDialog('New magazine',[{name:'name',label:'Name',required:true,value:'Magazine'},{name:'capacity',label:'Capacity',type:'number',min:1,max:999,value:6}]);if(!value)return;
      const magazine={id:newId(),name:value.name,capacity:Number(value.capacity),loaded:0,loadedAmmoId:''};
      await mutate(tokenId,t=>{if(t.odyssey.inventory.magazines.length>=200)throw new Error('Magazine limit is 200.');t.odyssey.inventory.magazines.push(magazine);},true);
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
    busy=true;node.disabled=true;
    handle(node.dataset.odAction,node).catch(e=>status(e.message,'error')).finally(()=>{busy=false;render();if(node.isConnected)node.disabled=false;});
  });
  const reload=()=>{try {library=readLibrary(window.localStorage,sanitize);storageError='';render();}catch(e){storageError=e.message;render();}};
  window.addEventListener('storage',event=>{if(event.key===LIBRARY_STORAGE_KEY)reload();});
  const disabledWhileBusy=html=>busy?html.replaceAll('<button ', '<button disabled '):html;
  return {renderInventory:tracker=>disabledWhileBusy(inventoryView(tracker,getRole()==='GM')),renderTemplates:()=>disabledWhileBusy(`${storageError?`<p class="cp-banner err">Local library could not be read: ${esc(storageError)}. Export JSON before replacing it.</p>`:''}${templatesView(library,getRole()==='GM')}`)};
}
