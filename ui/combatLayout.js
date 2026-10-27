// Move existing controls into HUD-style cards; event selectors and IDs stay intact.
export function arrangeCombatPanels(root=document) {
  for(const panel of root.querySelectorAll('[data-combat-layout]:not([data-combat-arranged])')) {
    panel.dataset.combatArranged='true';
    const kind=panel.dataset.combatLayout;
    const grid=document.createElement('div');grid.className='od-combat-grid';
    const cards=['Weapon & skill',kind==='manual'?'Manual defense':'Target','Action'].map(title=>{
      const card=document.createElement('section');card.className='od-combat-card';
      const heading=document.createElement('div');heading.className='od-combat-title';heading.textContent=title;card.append(heading);grid.append(card);return card;
    });
    for(const field of [...panel.querySelectorAll('label.field-stack')]) {
      const control=field.querySelector('[data-attack-field],[data-manual-attack-field],input,select,button,[id],.combat-value');
      const key=control?.dataset.attackField||control?.dataset.manualAttackField||control?.id||control?.dataset.action||'';
      const isWeapon=['skill','weaponName','weaponDamage','ammoId','rounds','attackSourceName','attackSkill','weaponAccuracy'].includes(key)||field.querySelector('[data-weapon-accuracy-display]');
      const isTarget=['distanceMeters','targetPart','targetName','targetTokenId','pick-attack-target','pickTargetBtn','clearTargetBtn','manualArmor','manualParry'].includes(key)||field.querySelector('.combat-value:not([data-weapon-accuracy-display])');
      cards[isWeapon?0:isTarget?1:2].append(field);
    }
    for(const hidden of [...panel.querySelectorAll('input[type="hidden"]')])cards[1].append(hidden);
    for(const note of [...panel.querySelectorAll('.combat-notes')])cards[2].append(note);
    const notes=[...panel.querySelectorAll(':scope > .muted')];
    if(notes.length) {
      const details=document.createElement('details');details.className='combat-notes';
      const summary=document.createElement('summary');summary.textContent='Attack rules';details.append(summary,...notes);cards[2].append(details);
    }
    const footer=document.createElement('div');footer.className='od-combat-footer';
    const attack=panel.querySelector('[data-action="perform-attack"],[data-action="perform-manual-attack"],#environmentAttackBtn');
    const pick=cards[1].querySelector('[data-action="pick-attack-target"],#pickTargetBtn');
    const wide=(button,label)=>{button.classList.remove('od-icon-button');button.classList.add('od-wide-action');button.textContent=label;button.title=label;button.setAttribute('aria-label',label);footer.append(button);};
    if(attack)wide(attack,kind==='manual'?'No Target Attack':'Attack');
    if(pick){const field=pick.closest('label');wide(pick,/cancel/i.test(pick.textContent||pick.title)?'Cancel Target Pick':'Pick Target On Map');field?.remove();}
    for(const obsolete of [...panel.querySelectorAll(':scope > .combat-action,:scope > .row')])obsolete.remove();
    for(const obsolete of [...panel.querySelectorAll(':scope > .combat-form-grid,:scope > .combat-modifiers')])obsolete.remove();
    panel.append(grid,footer);panel.classList.add('od-combat-panel');
  }
}
