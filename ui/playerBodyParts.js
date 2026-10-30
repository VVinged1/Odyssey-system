import {partWounds} from '../characterRules.js';
const esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function playerBodyPartsView(data,names){
 const parts=names.filter(key=>{const p=data.body[key];return p&&!p.hidden&&(key!=='Special'||p.max>0||p.current>0||p.minor||p.serious);});
 return '<div class="body-table-wrap" data-player-body-parts><table class="body-table od-body-readonly"><thead><tr><th>Part</th><th>HP</th><th>Armor</th><th>M / S / C</th></tr></thead><tbody>'+parts.map(key=>{const p=data.body[key],w=partWounds(data,key);return '<tr data-body-part="'+esc(key)+'"><td>'+esc(p.label||key)+'</td><td>'+p.current+' / '+p.max+'</td><td>'+p.armor+'</td><td>'+w.minor+' / '+w.serious+' / '+w.critical+'</td></tr>';}).join('')+'</tbody></table></div>';
}
