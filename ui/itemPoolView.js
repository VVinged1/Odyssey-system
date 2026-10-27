import {esc} from './characterView.js';
export function showItemPool({getCatalog,character,onGive,onDelete,onExport,onImport}) {
  const dialog=document.createElement('dialog');dialog.className='od-dialog od-pool-dialog';
  dialog.innerHTML='<div class="cp-row"><h2>Campaign item pool</h2><button type="button" data-pool-close aria-label="Close item pool">×</button></div><p>Give reusable items to '+esc(character)+'. The pool is shared across scenes in this Owlbear room. Copied equipment is unequipped; magazines are empty.</p><div class="cp-row"><input type="search" data-pool-search placeholder="Search items" aria-label="Search item pool"><button type="button" data-pool-export>Export item pool</button><button type="button" data-pool-import>Import item pool</button></div><p data-pool-status role="status"></p><div class="cp-list" data-pool-list></div>';
  document.body.append(dialog);let busy=false;
  const message=dialog.querySelector('[data-pool-status]');
  const refresh=()=>{
    const catalog=getCatalog(),search=dialog.querySelector('[data-pool-search]').value.trim().toLowerCase();
    const entries=catalog.entries.filter(e=>(e.blueprint.values.name+' '+e.blueprint.values.kind).toLowerCase().includes(search));
    dialog.querySelector('[data-pool-list]').innerHTML=entries.map(e=>{const b=e.blueprint,v=b.values;return '<article class="cp-card cp-row"><b>'+esc(v.name)+'</b><span class="cp-pill">'+esc(v.kind)+'</span><span class="cp-muted">'+esc([v.armor?'Armor '+v.armor:'',v.criticalMax?'HP '+v.criticalMax:'',v.capacity?'Capacity '+v.capacity:'',v.damage?'DMG '+v.damage:''].filter(Boolean).join(' · '))+'</span><button type="button" data-pool-give="'+esc(e.id)+'" '+(busy?'disabled':'')+'>Give item</button><button type="button" data-pool-delete="'+esc(e.id)+'" class="danger" aria-label="Remove definition from pool" '+(busy?'disabled':'')+'>×</button></article>';}).join('')||'<p class="cp-muted">No matching items.</p>';
    if(catalog.errors)message.textContent='Some catalog entries could not be read. Export a backup before importing changes.';
  };
  const run=async fn=>{if(busy)return;busy=true;refresh();try{const result=await fn();if(result!==false)message.textContent='Saved.';}catch(e){message.textContent=e.message;}finally{busy=false;refresh();}};
  dialog.querySelector('[data-pool-search]').oninput=refresh;
  dialog.addEventListener('click',event=>{
    const node=event.target.closest('button');if(!node)return;
    if(node.dataset.poolGive)void run(()=>onGive(node.dataset.poolGive));
    if(node.dataset.poolDelete)void run(()=>onDelete(node.dataset.poolDelete));
    if(node.hasAttribute('data-pool-export'))void run(onExport);
    if(node.hasAttribute('data-pool-import'))void run(onImport);
  });
  let finish;const closed=new Promise(resolve=>finish=resolve);const close=()=>{if(busy)return;dialog.close();dialog.remove();finish();};
  dialog.querySelector('[data-pool-close]').onclick=close;dialog.addEventListener('cancel',event=>{event.preventDefault();close();});refresh();dialog.showModal();return {closed,refresh};
}
