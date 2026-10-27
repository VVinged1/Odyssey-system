const escape=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export const descriptionButton=(title,text)=>'<button type="button" class="cp-info" data-description-title="'+escape(title)+'" data-description-text="'+escape(text||'No description yet.')+'" aria-label="Description: '+escape(title)+'" title="Show description">i</button>';
export function installDescriptions() {
 document.addEventListener('click',event=>{
  const node=event.target.closest?.('[data-description-title]');if(!node)return;event.preventDefault();event.stopPropagation();
  const dialog=document.createElement('dialog');dialog.className='od-dialog od-description-dialog';const title=document.createElement('h2'),body=document.createElement('p'),close=document.createElement('button');
  title.textContent=node.dataset.descriptionTitle;body.textContent=node.dataset.descriptionText||'No description yet.';close.textContent='Close';close.type='button';close.setAttribute('aria-label','Close description');close.dataset.descriptionClose='';
  const finish=()=>{dialog.close();dialog.remove();};close.onclick=finish;dialog.addEventListener('cancel',event=>{event.preventDefault();finish();});dialog.append(title,body,close);document.body.append(dialog);dialog.showModal();
 });
}
