export function installTooltips(){
 const tip=document.createElement('div');tip.className='od-tooltip';tip.id='od-hover-tooltip';tip.setAttribute('role','tooltip');tip.setAttribute('popover','manual');document.body.append(tip);let active=null;
 const hide=()=>{tip.classList.remove('visible');active?.removeAttribute('aria-describedby');active=null;};
 const show=node=>{hide();const text=node.dataset.descriptionText||node.dataset.tooltip||node.getAttribute('title')||node.getAttribute('aria-label');if(!text)return;active=node;tip.textContent=text;node.setAttribute('aria-describedby',tip.id);const host=node.closest('dialog[open]')||document.body;if(tip.matches(':popover-open'))tip.hidePopover();if(tip.parentElement!==host)host.append(tip);tip.showPopover();const r=node.getBoundingClientRect();tip.style.left=Math.max(8,Math.min(r.left,innerWidth-328))+'px';tip.style.top=Math.max(8,Math.min(r.bottom+7,innerHeight-tip.offsetHeight-8))+'px';requestAnimationFrame(()=>{if(active===node)tip.classList.add('visible');});};
 document.addEventListener('pointerover',e=>{const node=e.target.closest?.('button,[data-tooltip]');if(node&&node!==active&&!node.contains(e.relatedTarget))show(node);});
 document.addEventListener('pointerout',e=>{if(active&&!active.contains(e.relatedTarget))hide();});
 document.addEventListener('focusin',e=>{const node=e.target.closest?.('button,[data-tooltip]');if(node)show(node);});document.addEventListener('focusout',hide);document.addEventListener('pointerdown',hide);document.addEventListener('scroll',hide,true);document.addEventListener('keydown',e=>{if(e.key==='Escape')hide();});
 new MutationObserver(()=>{if(active&&!active.isConnected)hide();}).observe(document.body,{childList:true,subtree:true});
}
