import {normalizeWeaponCategory,normalizeRangeReduction} from './rangeRules.js';
import { normalizeBonuses, reconcileEquipment, equipmentFits } from "./characterRules.js";
// Token-owned inventory. Equipped magazines mirror the V1 weapon fields so
// attacks/reloads keep using ammunition.js without a second source of damage math.
const LIMIT = 2000;
const int = (v, max = 999999) => Math.min(max, Math.max(0, Math.floor(Number(v) || 0)));
const text = (v, max = 120) => String(v ?? '').trim().slice(0, max);
const kinds = new Set(['item', 'consumable', 'medical', 'tool', 'resource', 'armor', 'shield', 'implant']);
export function newId() { return globalThis.crypto.randomUUID(); }
function rows(raw, map) {
  const seen = new Set();
  return (Array.isArray(raw) ? raw : []).slice(0, LIMIT).filter(r => r && typeof r === 'object').map(map)
    .filter(r => r.id && !seen.has(r.id) && seen.add(r.id));
}
export function normalizeInventory(raw, weapons = []) {
  const items = rows(raw?.items, (r, i) => ({
    id: text(r.id || `legacy-item-${i}`), name: text(r.name) || 'Item',
    kind: kinds.has(r.kind) ? r.kind : 'item', quantity: int(r.quantity ?? 1),
    description: text(r.description, 2000), armor: int(r.armor, 999),
    subtype: ['grenade','medkit','repairkit','stim'].includes(r.subtype)?r.subtype:r.kind==='medical'?'medkit':'grenade',
    healCritical: int(r.healCritical ?? 1,999), slots:[...new Set((Array.isArray(r.slots)?r.slots:[r.slot||'any']).map(s=>text(s,40)).filter(Boolean))].slice(0,100), slot: text(r.slot||'any',40),
    criticalMax:int(r.criticalMax,999), criticalDamage:int(r.criticalDamage,int(r.criticalMax,999)),
    playerRemovable:r.playerRemovable===true, bonuses:normalizeBonuses(['resource','implant'].includes(r.kind)?{...r.bonuses,rangeReduction:{}}:r.bonuses),
    abilities:(Array.isArray(r.abilities)?r.abilities:[]).slice(0,20).map((a,j)=>({id:text(a.id||'ability-'+j),name:text(a.name)||'Ability',level:int(a.level,10),description:text(a.description,2000)})),
    legacyEquipment: r.legacyEquipment ?? (raw?.version<3 || !raw?.version),
    equippedPart: text(r.equippedPart, 40), previousArmor: int(r.previousArmor, 999),
  }));
  // A worn unit owns its durability and bonuses independently of spare copies.
  for(const item of [...items])if(item.equippedPart&&['armor','shield','implant'].includes(item.kind)&&item.quantity>1) {
    let id=(item.id.slice(0,90)+'-reserve'),suffix=1;while(items.some(i=>i.id===id))id=item.id.slice(0,90)+'-reserve-'+suffix++;
    items.push({...structuredClone(item),id,quantity:item.quantity-1,equippedPart:'',previousArmor:0,legacyEquipment:false});item.quantity=1;
  }
  const magazines = rows(raw?.magazines, (r, i) => ({
    id: text(r.id || `legacy-magazine-${i}`), name: text(r.name) || 'Magazine',
    description:text(r.description,2000), capacity: int(r.capacity, 999), loadedAmmoId: text(r.loadedAmmoId),
    loaded: int(r.loaded, int(r.capacity, 999)),
  }));
  const assigned = new Set();
  for (const [i, weapon] of weapons.entries()) {
    if(weapon.acceptsMagazines===false){weapon.magazineId='';continue;}
    const requested = text(weapon.magazineId);
    // A legacy weapon's magazine gets a stable ID; reading does not spend ammo.
    let id = requested || ((!Object.hasOwn(weapon, 'magazineId') || weapon.loaded > 0) && weapon.capacity > 0 ? `legacy-mag-${weapon.id || i}` : '');
    if (!id) { weapon.magazineId = ''; weapon.loaded = 0; weapon.loadedAmmoId = ''; continue; }
    if (assigned.has(id)) id = `recovered-mag-${weapon.id || i}-${i}`;
    let mag = magazines.find(m => m.id === id);
    if (!mag) { mag = {id, name: `${weapon.name} magazine`, description:'', capacity: weapon.capacity, loaded: weapon.loaded, loadedAmmoId: weapon.loadedAmmoId}; magazines.push(mag); }
    weapon.magazineId = id;
    assigned.add(id);
    // V1 remains authoritative while mounted. Normalization persists this mirror
    // after consumeAmmo/reloadMagazine, including external snapshot imports.
    Object.assign(mag, {capacity: weapon.capacity, loaded: weapon.loaded, loadedAmmoId: weapon.loadedAmmoId});
  }
  return {version:5,items,magazines,armorHp:Object.fromEntries(Object.entries(raw?.armorHp||{}).filter(([k])=>!['__proto__','constructor','prototype'].includes(k)).map(([k,v])=>[text(k,40),{max:int(v?.max,99999),remaining:int(v?.remaining,99999)}])),baseMax:Object.fromEntries(Object.entries(raw?.baseMax||{}).filter(([k])=>!['__proto__','constructor','prototype'].includes(k)).map(([k,v])=>[text(k,40),int(v,99999)])),baseArmor:Object.fromEntries(Object.entries(raw?.baseArmor||{}).filter(([k])=>!['__proto__','constructor','prototype'].includes(k)).map(([k,v])=>[text(k,40),int(v,999)]))};
}
export function magazineWeapon(odyssey, id) {
  return odyssey.weapons.ranged.find(w => w.magazineId === id);
}
export function magazineCompatible(odyssey,weapon,magazine) {
  if(weapon.acceptsMagazines===false)return false;
  if(!Array.isArray(weapon.acceptedMagazineIds))return magazine.capacity===weapon.capacity&&(!magazine.loaded||weapon.ammoIds.includes(magazine.loadedAmmoId));
  return weapon.acceptedMagazineIds.some(id=>id===magazine.id||odyssey.inventory.magazines.some(reference=>reference.id===id&&reference.name===magazine.name&&reference.capacity===magazine.capacity));
}
export function swapMagazine(odyssey, weaponIndex, magazineId) {
  const weapon = odyssey.weapons.ranged[weaponIndex];
  if (!weapon) throw new Error('Weapon no longer exists.');
  if(weapon.acceptsMagazines===false)throw new Error('This weapon has a built-in magazine.');
  const current = odyssey.inventory.magazines.find(m => m.id === weapon.magazineId);
  if (current) Object.assign(current, {capacity: weapon.capacity, loaded: weapon.loaded, loadedAmmoId: weapon.loadedAmmoId});
  const next = magazineId ? odyssey.inventory.magazines.find(m => m.id === magazineId) : null;
  if (magazineId && !next) throw new Error('Magazine no longer exists.');
  if (next && magazineWeapon(odyssey, next.id) && weapon.magazineId !== next.id) throw new Error('Magazine is already mounted.');
  const selectedMagazines=Array.isArray(weapon.acceptedMagazineIds);
  if(next&&selectedMagazines&&!magazineCompatible(odyssey,weapon,next))throw new Error('This magazine is incompatible with the weapon.');
  if (next && !selectedMagazines && next.capacity !== weapon.capacity) throw new Error('Magazine capacity must match this weapon.');
  if (next?.loaded && !selectedMagazines && !weapon.ammoIds.includes(next.loadedAmmoId)) throw new Error('Magazine ammunition is incompatible.');
  if(selectedMagazines){weapon.capacity=next?.capacity||0;weapon.ammoIds=next?.loadedAmmoId?[next.loadedAmmoId]:[];}
  weapon.magazineId = next?.id || '';
  weapon.loaded = next?.loaded || 0;
  weapon.loadedAmmoId = next?.loadedAmmoId || '';
}
export function fillMagazine(odyssey, id, ammoId) {
  const mag = odyssey.inventory.magazines.find(m => m.id === id);
  const ammo = odyssey.ammunition.find(a => a.id === ammoId);
  if (!mag || !ammo) throw new Error('Choose a magazine and ammunition.');
  if (magazineWeapon(odyssey, id)) throw new Error('Detach the magazine before filling it, or use weapon reload.');
  if (mag.loaded && mag.loadedAmmoId !== ammoId) throw new Error('Empty the magazine before changing ammunition.');
  const count = Math.min(mag.capacity - mag.loaded, ammo.quantity);
  if (count < 1) throw new Error('Magazine is full or reserve is empty.');
  mag.loadedAmmoId = ammoId; mag.loaded += count; ammo.quantity -= count;
}
// GM correction: adjusts the magazine count without transferring reserve ammo.
// Permissions are checked again at the controller's mutation boundary.
export function setMagazineRounds(odyssey, id, rounds, ammoId) {
  const mag = odyssey.inventory.magazines.find(m => m.id === id);
  if (!mag) throw new Error('Magazine no longer exists.');
  if (!Number.isInteger(rounds) || rounds < 0 || rounds > mag.capacity) {
    throw new Error(`Rounds must be a whole number from 0 to ${mag.capacity}.`);
  }
  const selectedAmmoId = String(ammoId ?? mag.loadedAmmoId ?? '');
  const ammo = odyssey.ammunition.find(a => a.id === selectedAmmoId);
  const weapon = magazineWeapon(odyssey, id);
  if ((rounds > 0 || selectedAmmoId) && !ammo) throw new Error('Choose a valid ammunition type.');
  if (weapon && !Array.isArray(weapon.acceptedMagazineIds) && selectedAmmoId && !weapon.ammoIds.includes(selectedAmmoId)) {
    throw new Error('This ammunition is incompatible with the mounted weapon.');
  }
  mag.loaded = rounds;
  mag.loadedAmmoId = selectedAmmoId;
  if (weapon) {
    if(Array.isArray(weapon.acceptedMagazineIds))weapon.ammoIds=selectedAmmoId?[selectedAmmoId]:[];
    weapon.loaded = rounds;
    weapon.loadedAmmoId = selectedAmmoId;
  }
}
export function emptyMagazine(odyssey, id) {
  const mag = odyssey.inventory.magazines.find(m => m.id === id);
  if (!mag) throw new Error('Magazine no longer exists.');
  if (magazineWeapon(odyssey, id)) throw new Error('Detach the magazine first.');
  const ammo = odyssey.ammunition.find(a => a.id === mag.loadedAmmoId);
  if (mag.loaded && !ammo) throw new Error('The loaded ammunition is missing.');
  if (ammo && ammo.quantity + mag.loaded > 999999) throw new Error('Ammo reserve is full.');
  if (ammo) ammo.quantity += mag.loaded;
  mag.loaded = 0; mag.loadedAmmoId = '';
}
export function equipItem(tracker,id,partName,gm=true) {
  const item=tracker.odyssey.inventory.items.find(i=>i.id===id);
  if(!item||!['armor','shield','implant'].includes(item.kind))throw new Error('Choose equipment.');
  if(item.kind==='implant'&&!gm&&!item.playerRemovable)throw new Error('Only the GM can equip this implant.');
  if(item.quantity<1||item.equippedPart)throw new Error('Item is unavailable or already equipped.');
  const part=tracker.body[partName];if(!part)throw new Error('Body part no longer exists.');
  if(!equipmentFits(item,partName,part))throw new Error('This item does not fit the selected slot.');
  if(item.kind!=='implant'&&tracker.odyssey.inventory.items.some(i=>['armor','shield'].includes(i.kind)&&i.equippedPart===partName))throw new Error('Armor is already equipped on this part.');
  if(item.quantity>1) {
    if(tracker.odyssey.inventory.items.length>=1000)throw new Error('Inventory is full.');
    tracker.odyssey.inventory.items.push({...structuredClone(item),id:newId(),quantity:item.quantity-1,equippedPart:'',previousArmor:0,legacyEquipment:false});item.quantity=1;
  }
  tracker.odyssey.inventory.baseArmor??={};
  if(!Object.hasOwn(tracker.odyssey.inventory.baseArmor,partName))tracker.odyssey.inventory.baseArmor[partName]=part.armor;
  item.equippedPart=partName;item.legacyEquipment=false;reconcileEquipment(tracker);
}
export function unequipItem(tracker,id,gm=true) {
  const item=tracker.odyssey.inventory.items.find(i=>i.id===id);if(!item)throw new Error('Item no longer exists.');
  if(item.kind==='implant'&&!gm&&!item.playerRemovable)throw new Error('Only the GM can remove this implant.');
  const part=tracker.body[item.equippedPart],inv=tracker.odyssey.inventory;
  if(part&&Object.hasOwn(inv.baseArmor||{},item.equippedPart)) {
    const expected=inv.baseArmor[item.equippedPart]+inv.items.filter(i=>i.equippedPart===item.equippedPart&&['armor','shield','implant'].includes(i.kind)&&i.quantity>0&&(i.kind==='implant'||!i.criticalMax||i.criticalDamage<i.criticalMax)).reduce((sum,i)=>sum+i.armor,0);
    if(part.armor!==expected)inv.baseArmor[item.equippedPart]=part.armor;
  }
  item.equippedPart='';reconcileEquipment(tracker);
}
export function consumeItem(tracker, id) {
  const item = tracker.odyssey.inventory.items.find(i => i.id === id);
  if (!item || !['consumable','medical','resource'].includes(item.kind) || item.quantity < 1) throw new Error('No usable item available.');
  item.quantity -= 1; // No new automatic healing or combat rules.
}
export function inventoryAmmoTotal(odyssey) {
  return odyssey.weapons.ranged.filter(w=>w.acceptsMagazines===false).reduce((sum,w)=>sum+w.loaded,0)+odyssey.ammunition.reduce((n,a) => n + a.quantity, 0) + odyssey.inventory.magazines.reduce((n,m) => n + m.loaded, 0);
}

// Typed editor saves through existing V1 weapon/ammo collections; no duplicate inventory.
export function findInventoryEntry(tracker,type,id) {
  const o=tracker.odyssey;
  if(type==='ammo')return o.ammunition.find(a=>a.id===id);
  if(type==='magazine')return o.inventory.magazines.find(m=>m.id===id);
  if(type?.startsWith('weapon-'))return o.weapons[type.slice(7)]?.find((w,index)=>(w.id||String(index))===id);
  return o.inventory.items.find(i=>i.id===id);
}
export function deleteInventoryEntry(tracker,type,id) {
  const o=tracker.odyssey;
  if(type==='ammo') {
    if(o.inventory.magazines.some(m=>m.loadedAmmoId===id&&m.loaded)||o.weapons.ranged.some(w=>w.acceptsMagazines===false&&w.loadedAmmoId===id&&w.loaded))throw new Error('Empty all magazines using this ammo first.');
    o.ammunition=o.ammunition.filter(a=>a.id!==id);for(const w of o.weapons.ranged)w.ammoIds=w.ammoIds.filter(a=>a!==id);return;
  }
  if(type==='magazine') {
    const removed=o.inventory.magazines.find(m=>m.id===id),replacement=removed&&o.inventory.magazines.find(m=>m.id!==id&&m.name===removed.name&&m.capacity===removed.capacity);const index=o.weapons.ranged.findIndex(w=>w.magazineId===id);if(index>=0)swapMagazine(o,index,'');emptyMagazine(o,id);o.inventory.magazines=o.inventory.magazines.filter(m=>m.id!==id);for(const w of o.weapons.ranged)if(w.acceptedMagazineIds)w.acceptedMagazineIds=[...new Set(w.acceptedMagazineIds.flatMap(m=>m===id?replacement?[replacement.id]:[]:[m]))];return;
  }
  if(type?.startsWith('weapon-')) {
    const kind=type.slice(7),index=o.weapons[kind].findIndex((w,j)=>(w.id||String(j))===id);
    if(index<0)throw new Error('Weapon no longer exists.');if(kind==='ranged'){const w=o.weapons.ranged[index];if(w.acceptsMagazines===false&&w.loaded){const a=o.ammunition.find(a=>a.id===w.loadedAmmoId);if(!a||a.quantity+w.loaded>999999)throw new Error('Cannot return loaded ammunition to reserve.');a.quantity+=w.loaded;}else if(w.magazineId)swapMagazine(o,index,'');}o.weapons[kind].splice(index,1);return;
  }
  const item=o.inventory.items.find(i=>i.id===id);if(!item)throw new Error('Item no longer exists.');if(item.equippedPart)unequipItem(tracker,id,true);o.inventory.items=o.inventory.items.filter(i=>i.id!==id);
}
export function saveInventoryEntry(tracker,origin,values,id) {
  const o=tracker.odyssey,existing=origin?findInventoryEntry(tracker,origin.type,origin.id):null;
  if(origin&&!existing)throw new Error('Item no longer exists.');
  const kind=values.kind,weaponKind=values.weaponKind==='ranged'?'ranged':'melee';
  const dest=kind==='weapon'?'weapon-'+weaponKind:kind==='ammo'?'ammo':kind==='magazine'?'magazine':'item';
  if(origin&&origin.type!==dest)deleteInventoryEntry(tracker,origin.type,origin.id);
  const same=origin?.type===dest?existing:null;
  if(kind==='ammo') {
    const row={id,name:values.name,description:text(values.description,2000),damage:Number(values.damage)||0,penetration:Number(values.penetration)||0,quantity:Number(values.quantity)||0};
    if(same)Object.assign(same,row);else {if(o.ammunition.length>=200)throw new Error('Too many ammo types.');o.ammunition.push(row);}return;
  }
  if(kind==='magazine') {
    const capacity=Number(values.capacity),loaded=Number(values.loaded);
    if(!Number.isInteger(capacity)||capacity<1||capacity>999||!Number.isInteger(loaded)||loaded<0||loaded>capacity)throw new Error('Invalid magazine capacity or rounds.');
    const mounted=same&&magazineWeapon(o,same.id);if(mounted&&capacity!==mounted.capacity)throw new Error('Detach the magazine before changing its capacity.');
    if(same){same.capacity=capacity;same.name=values.name;same.description=text(values.description,2000);setMagazineRounds(o,same.id,loaded,values.loadedAmmoId);}
    else {
      const quantity=Number(values.quantity??1);if(!Number.isInteger(quantity)||quantity<1||quantity>1000||o.inventory.magazines.length+quantity>1000)throw new Error('Invalid magazine quantity or inventory is full.');
      for(let i=0;i<quantity;i++){const magazineId=i?id+'-copy-'+i:id;o.inventory.magazines.push({id:magazineId,name:values.name,description:text(values.description,2000),capacity,loaded:0,loadedAmmoId:''});setMagazineRounds(o,magazineId,loaded,values.loadedAmmoId);}
    }return;
  }
  if(kind==='weapon') {
    if([...o.weapons.melee,...o.weapons.ranged].some(w=>w!==same&&w.name===values.name))throw new Error('A weapon with this name already exists.');
    const row={id,name:values.name,description:text(values.description,2000),accuracy:Number(values.accuracy)||0,damage:Number(values.damage)||0};
    if(weaponKind==='ranged') {
      row.weaponCategory=normalizeWeaponCategory(values.weaponCategory);row.rangeReduction=normalizeRangeReduction(values.rangeReduction);
      const acceptsMagazines=values.acceptsMagazines!==false;
      if(same&&(same.acceptsMagazines!==false)!==acceptsMagazines&&same.loaded>0)throw new Error('Empty or detach the current magazine before changing magazine type.');
      if(acceptsMagazines) {
        const acceptedMagazineIds=[...new Set(values.acceptedMagazineIds||[])];
        if(acceptedMagazineIds.some(id=>!o.inventory.magazines.some(m=>m.id===id)))throw new Error('Selected magazine no longer exists.');
        if(same?.magazineId&&!magazineCompatible(o,{...same,acceptedMagazineIds},o.inventory.magazines.find(m=>m.id===same.magazineId)))throw new Error('Detach the current magazine before removing its compatibility.');
        Object.assign(row,{acceptsMagazines:true,acceptedMagazineIds,capacity:same?.magazineId?same.capacity:0,ammoIds:same?.magazineId?(same.ammoIds||[]):[],loaded:same?.loaded||0,loadedAmmoId:same?.loadedAmmoId||'',magazineId:same?.magazineId||''});
      } else {
        const capacity=Number(values.capacity),ammoIds=values.ammoIds||[],loaded=Number(values.loaded??same?.loaded??0),loadedAmmoId=String(values.loadedAmmoId??same?.loadedAmmoId??'');
        if(!Number.isInteger(capacity)||capacity<1||capacity>999||!Number.isInteger(loaded)||loaded<0||loaded>capacity)throw new Error('Invalid built-in magazine capacity or rounds.');
        if(loaded&&(!ammoIds.includes(loadedAmmoId)||!o.ammunition.some(a=>a.id===loadedAmmoId)))throw new Error('Choose compatible ammunition for loaded rounds.');
        Object.assign(row,{acceptsMagazines:false,acceptedMagazineIds:[],capacity,ammoIds,loaded,loadedAmmoId:loaded?loadedAmmoId:'',magazineId:''});
      }
    }
    if(same)Object.assign(same,row);else{if(o.weapons[weaponKind].length>=20)throw new Error('Weapon limit is 20 per type.');o.weapons[weaponKind].push(row);}return;
  }
  if(same?.equippedPart&& !['armor','shield','implant'].includes(kind))unequipItem(tracker,id,true);
  const row={...values,id,quantity:Number(values.quantity)||0,armor:Number(values.armor)||0,healCritical:Number(values.healCritical)||0,criticalMax:Number(values.criticalMax)||0,criticalDamage:Number(values.criticalDamage)||0,playerRemovable:values.playerRemovable===true,legacyEquipment:false,equippedPart:same?.equippedPart||'',previousArmor:same?.previousArmor||0};
  if(!['armor','shield','implant'].includes(kind)){row.armor=0;row.criticalMax=0;row.criticalDamage=0;}
  if(kind==='implant')row.criticalDamage=0;
  if(row.equippedPart&&row.quantity>1&&o.inventory.items.length>=1000)throw new Error('Inventory is full; cannot separate equipped and spare items.');
  if(row.criticalDamage>row.criticalMax)throw new Error('Critical damage cannot exceed durability.');
  if(same?.equippedPart&&!equipmentFits(row,same.equippedPart,tracker.body[same.equippedPart]))throw new Error('Unequip before changing the equipment slot.');
  if(same)Object.assign(same,row);else{if(o.inventory.items.length>=1000)throw new Error('Inventory limit is 1000 items.');o.inventory.items.push(row);}reconcileEquipment(tracker);
}
