import {groupNamedItems} from '../itemGroups.js';
import {descriptionButton} from './descriptions.js';
import {groupPoolEntries,poolCategoryPaths} from '../poolCategories.js';
import {esc} from './characterView.js';
export function showItemPool({getCatalog,character,onGive,onDelete,onExport,onImport,onRefresh,onAccountSave,onAccountLoad}) {
  const dialog=document.createElement('dialog');dialog.className='od-dialog od-pool-dialog';
  dialog.innerHTML='<div class="cp-row"><h2>Campaign item pool</h2><button type="button" data-pool-close aria-label="Close item pool">×</button></div><p>Give reusable items to '+esc(character)+'. The pool is shared in this scene and backed up in this browser for this room. Open the extension after changing scenes to restore the pool. Copied equipment is unequipped; magazines are empty.</p><div class="cp-row"><input type="search" data-pool-search placeholder="Search items" aria-label="Search item pool"><button type="button" data-pool-refresh title="Refresh item pool" aria-label="Refresh item pool">↻</button><button type="button" data-pool-export>Export item pool</button><button type="button" data-pool-import>Import item pool</button><button type="button" data-pool-account-save title="Save item pool to account" aria-label="Save item pool to account">Save to account</button><button type="button" data-pool-account-load title="Load item pool from account" aria-label="Load item pool from account">Load from account</button></div><div class="cp-row"><button type="button" data-pool-expand title="Expand categories" aria-label="Expand categories">▾</button><button type="button" data-pool-collapse title="Collapse categories" aria-label="Collapse categories">▸</button></div><p data-pool-warning class="cp-banner err" role="alert" hidden></p><p data-pool-status role="status"></p><div class="cp-list" data-pool-list></div>';
  document.body.append(dialog);let busy=false;const expanded=new Map();
  const message=dialog.querySelector('[data-pool-status]');
  const refresh=()=>{
    const catalog=getCatalog(),search=dialog.querySelector('[data-pool-search]').value.trim().toLowerCase();
    const warning=dialog.querySelector('[data-pool-warning]');warning.hidden=!catalog.warning;warning.textContent=catalog.warning||'';
    const list=dialog.querySelector('[data-pool-list]');
    for(const details of list.querySelectorAll('details[data-pool-category]'))expanded.set(details.dataset.poolCategory,details.open);
    const row=e=>{const b=e.blueprint,v=b.values;return '<article class="cp-card cp-row"><b>'+esc(v.name)+'</b><span class="cp-pill">'+esc(v.kind)+'</span><span class="cp-muted">'+esc([v.armor?'Armor '+v.armor:'',v.criticalMax?'HP '+v.criticalMax:'',v.capacity?'Capacity '+v.capacity:'',v.damage?'DMG '+v.damage:'',v.useSkill?'Use: '+v.useSkill:''].filter(Boolean).join(' · '))+'</span>'+descriptionButton(v.name,v.description)+'<button type="button" data-pool-give="'+esc(e.id)+'" '+(busy?'disabled':'')+'>Give item</button><button type="button" data-pool-delete="'+esc(e.id)+'" class="danger" aria-label="Remove definition from pool" '+(busy?'disabled':'')+'>×</button></article>';};
    const rows=entries=>groupNamedItems(entries,e=>({...e.blueprint.values,kind:e.blueprint.type==='item'?e.blueprint.values.kind:e.id})).map(group=>group.length===1?row(group[0]):'<details class="cp-card od-named-group"><summary>'+esc(group[0].blueprint.values.name)+' · '+group.length+' variants</summary>'+group.map(row).join('')+'</details>').join('');
    if(search) {
      const words=search.split(/\s+/);const matches=catalog.entries.filter(e=>{const text=[e.blueprint.values.name,e.blueprint.values.kind,e.blueprint.values.description||'',...poolCategoryPaths(e.blueprint).flat()].join(' ').toLowerCase();return words.every(word=>text.includes(word));});
      list.innerHTML=rows(matches)||'<p class="cp-muted">No matching items.</p>';
    } else {
      const group=(node,path=[],level=0)=>{const key=JSON.stringify([...path,node.label]),open=expanded.get(key)??level===0;return '<details class="od-pool-category" data-pool-category="'+esc(key)+'" '+(open?'open':'')+'><summary>'+esc(node.label)+' <span class="cp-pill">'+node.count+'</span></summary><div class="od-pool-category-body">'+rows(node.entries)+node.children.map(child=>group(child,[...path,node.label],level+1)).join('')+'</div></details>';};
      list.innerHTML=groupPoolEntries(catalog.entries).map(node=>group(node)).join('')||'<p class="cp-muted">The campaign item pool is empty.</p>';
    }
    if(catalog.errors)message.textContent='Some catalog entries could not be read. Export a backup before importing changes.';
  };
  const run=async fn=>{if(busy)return;busy=true;refresh();try{const result=await fn();if(result!==false)message.textContent=result?.message||'Saved.';}catch(e){message.textContent=e.message;}finally{busy=false;refresh();}};
  dialog.querySelector('[data-pool-search]').oninput=refresh;
  dialog.addEventListener('toggle',event=>{if(event.target instanceof HTMLDetailsElement&&event.target.isConnected&&event.target.dataset.poolCategory)expanded.set(event.target.dataset.poolCategory,event.target.open);},true);
  dialog.addEventListener('click',event=>{
    const node=event.target.closest('button');if(!node)return;
    if(node.hasAttribute('data-pool-expand')||node.hasAttribute('data-pool-collapse')){const open=node.hasAttribute('data-pool-expand');for(const details of dialog.querySelectorAll('details[data-pool-category]')){details.open=open;expanded.set(details.dataset.poolCategory,open);}}
    if(node.dataset.poolGive)void run(()=>onGive(node.dataset.poolGive));
    if(node.dataset.poolDelete)void run(()=>onDelete(node.dataset.poolDelete));
    if(node.hasAttribute('data-pool-account-save')&&onAccountSave)void run(onAccountSave);
    if(node.hasAttribute('data-pool-account-load')&&onAccountLoad)void run(onAccountLoad);
    if(node.hasAttribute('data-pool-refresh')&&onRefresh)void run(onRefresh);
    if(node.hasAttribute('data-pool-export'))void run(onExport);
    if(node.hasAttribute('data-pool-import'))void run(onImport);
  });
  let finish;const closed=new Promise(resolve=>finish=resolve);const close=()=>{if(busy)return;dialog.close();dialog.remove();finish();};
  dialog.querySelector('[data-pool-close]').onclick=close;dialog.addEventListener('cancel',event=>{event.preventDefault();close();});refresh();dialog.showModal();return {closed,refresh};
}
