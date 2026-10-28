import {installTooltips} from './tooltips.js';
// Shared compact action buttons. Character/stat cards retain their full content.
const paths={
 info:'M12 11v6 M12 7h.01 M22 12a10 10 0 1 1-20 0 10 10 0 0 1 20 0',
 plus:'M12 5v14M5 12h14',close:'m6 6 12 12M18 6 6 18',edit:'m4 16-1 5 5-1L20 8l-4-4Z M14 6l4 4',
 save:'M5 3h13l3 3v15H3V3Z M7 3v6h10V3 M7 21v-8h10v8',
 refresh:'M20 7v5h-5M4 17v-5h5 M5 8a8 8 0 0 1 13-3l2 3 M19 16a8 8 0 0 1-13 3l-2-3',
 target:'M12 2v4m0 12v4M2 12h4m12 0h4 M19 12a7 7 0 1 1-14 0 7 7 0 0 1 14 0 M12 10v4m-2-2h4',
 download:'M12 3v12m-5-5 5 5 5-5 M4 16v5h16v-5',upload:'M12 16V4m-5 5 5-5 5 5 M4 16v5h16v-5',
 cloud:'M6 18a4 4 0 0 1-1-8 7 7 0 0 1 13-1 5 5 0 0 1 0 9Z M12 10v8m-3-3 3 3 3-3',
 shield:'m12 3 8 3v6c0 5-8 9-8 9s-8-4-8-9V6Z',
 dice:'M4 4h16v16H4Z M8 8h.01 M16 8h.01 M12 12h.01 M8 16h.01 M16 16h.01',
 attack:'m4 3 16 17m0-17L4 20 M3 15l6 6m6-18 6 6',
 heal:'M9 3h6v6h6v6h-6v6H9v-6H3V9h6Z',
 book:'M3 4h7l2 2 2-2h7v15h-7l-2 2-2-2H3Z M12 6v15',
 play:'m8 4 12 8-12 8Z',box:'m3 7 9-4 9 4v13H3Z M3 7l9 4 9-4 M12 11v9',
 window:'M3 3h18v18H3Z M3 8h18 M8 8v13',copy:'M8 8h13v13H8Z M16 8V3H3v13h5'
};
function iconFor(text) {
 if(/description/i.test(text))return 'info';
 if(/cancel|clear|remove|delete|unequip|detach|close/i.test(text))return 'close';
 if(/edit|rename|set rounds|set magazine/i.test(text))return 'edit';
 if(/duplicate/i.test(text))return 'copy';
 if(/export|empty to|load from/i.test(text))return 'download';
 if(/import/i.test(text))return 'upload';
 if(/account/i.test(text))return 'cloud';
 if(/save/i.test(text))return 'save';
 if(/add|new|^\+/i.test(text))return 'plus';
 if(/reload|refresh|rebuild|fill|sync/i.test(text))return 'refresh';
 if(/select|pick|target|focus/i.test(text))return 'target';
 if(/roll|d100|dice/i.test(text))return 'dice';
 if(/attack|resolve/i.test(text))return 'attack';
 if(/heal|repair/i.test(text))return 'heal';
 if(/equip/i.test(text))return 'shield';
 if(/template/i.test(text))return 'book';
 if(/hud|window|open/i.test(text))return 'window';
 if(/use|apply/i.test(text))return 'play';
 return 'box';
}
export function installCompactButtons() {
 installTooltips();
 const enhance=()=>{
  for(const button of document.querySelectorAll('button')) {
   if(button.matches('.cp-attr,.cp-skill-card,.list-item,.od-wide-action')||button.querySelector('[data-compact-icon]'))continue;
   const text=button.textContent.trim();if(!text||text.length<=2)continue;
   button.title=text;button.setAttribute('aria-label',text);button.classList.add('od-icon-button');
   button.innerHTML=`<svg data-compact-icon viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="${paths[iconFor(text)]}"></path></svg>`;
  }
 };
 let scheduled=false;new MutationObserver(()=>{if(scheduled)return;scheduled=true;queueMicrotask(()=>{scheduled=false;enhance();});}).observe(document.body,{childList:true,subtree:true,characterData:true});enhance();
}
