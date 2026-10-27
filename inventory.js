// Token-owned inventory. Equipped magazines mirror the V1 weapon fields so
// attacks/reloads keep using ammunition.js without a second source of damage math.
const LIMIT = 200;
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
    equippedPart: text(r.equippedPart, 40), previousArmor: int(r.previousArmor, 999),
  }));
  const magazines = rows(raw?.magazines, (r, i) => ({
    id: text(r.id || `legacy-magazine-${i}`), name: text(r.name) || 'Magazine',
    capacity: int(r.capacity, 999), loadedAmmoId: text(r.loadedAmmoId),
    loaded: int(r.loaded, int(r.capacity, 999)),
  }));
  const assigned = new Set();
  for (const [i, weapon] of weapons.entries()) {
    const requested = text(weapon.magazineId);
    // A legacy weapon's magazine gets a stable ID; reading does not spend ammo.
    let id = requested || ((!Object.hasOwn(weapon, 'magazineId') || weapon.loaded > 0) && weapon.capacity > 0 ? `legacy-mag-${weapon.id || i}` : '');
    if (!id) { weapon.magazineId = ''; weapon.loaded = 0; weapon.loadedAmmoId = ''; continue; }
    if (assigned.has(id)) id = `recovered-mag-${weapon.id || i}-${i}`;
    let mag = magazines.find(m => m.id === id);
    if (!mag) { mag = {id, name: `${weapon.name} magazine`, capacity: weapon.capacity, loaded: weapon.loaded, loadedAmmoId: weapon.loadedAmmoId}; magazines.push(mag); }
    weapon.magazineId = id;
    assigned.add(id);
    // V1 remains authoritative while mounted. Normalization persists this mirror
    // after consumeAmmo/reloadMagazine, including external snapshot imports.
    Object.assign(mag, {capacity: weapon.capacity, loaded: weapon.loaded, loadedAmmoId: weapon.loadedAmmoId});
  }
  return {version: 2, items, magazines};
}
export function magazineWeapon(odyssey, id) {
  return odyssey.weapons.ranged.find(w => w.magazineId === id);
}
export function swapMagazine(odyssey, weaponIndex, magazineId) {
  const weapon = odyssey.weapons.ranged[weaponIndex];
  if (!weapon) throw new Error('Weapon no longer exists.');
  const current = odyssey.inventory.magazines.find(m => m.id === weapon.magazineId);
  if (current) Object.assign(current, {capacity: weapon.capacity, loaded: weapon.loaded, loadedAmmoId: weapon.loadedAmmoId});
  const next = magazineId ? odyssey.inventory.magazines.find(m => m.id === magazineId) : null;
  if (magazineId && !next) throw new Error('Magazine no longer exists.');
  if (next && magazineWeapon(odyssey, next.id) && weapon.magazineId !== next.id) throw new Error('Magazine is already mounted.');
  if (next && next.capacity !== weapon.capacity) throw new Error('Magazine capacity must match this weapon.');
  if (next?.loaded && !weapon.ammoIds.includes(next.loadedAmmoId)) throw new Error('Magazine ammunition is incompatible.');
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
export function equipItem(tracker, id, partName) {
  const item = tracker.odyssey.inventory.items.find(i => i.id === id);
  if (!item || !['armor','shield','implant'].includes(item.kind)) throw new Error('Choose equipment.');
  if (item.quantity < 1) throw new Error('Item is not available.');
  if (item.equippedPart) throw new Error('Unequip this item first.');
  const part = tracker.body[partName];
  if (!part) throw new Error('Body part no longer exists.');
  if (tracker.odyssey.inventory.items.some(i => i.equippedPart === partName)) throw new Error('This body part already has equipment.');
  item.previousArmor = part.armor; item.equippedPart = partName;
  part.armor = item.armor;
}
export function unequipItem(tracker, id) {
  const item = tracker.odyssey.inventory.items.find(i => i.id === id);
  if (!item) throw new Error('Item no longer exists.');
  const part = tracker.body[item.equippedPart];
  // Preserve a subsequent manual GM adjustment rather than overwriting it.
  if (part && part.armor === item.armor) part.armor = item.previousArmor;
  item.equippedPart = ''; item.previousArmor = 0;
}
export function consumeItem(tracker, id) {
  const item = tracker.odyssey.inventory.items.find(i => i.id === id);
  if (!item || !['consumable','medical','resource'].includes(item.kind) || item.quantity < 1) throw new Error('No usable item available.');
  item.quantity -= 1; // No new automatic healing or combat rules.
}
export function inventoryAmmoTotal(odyssey) {
  return odyssey.ammunition.reduce((n,a) => n + a.quantity, 0) + odyssey.inventory.magazines.reduce((n,m) => n + m.loaded, 0);
}
