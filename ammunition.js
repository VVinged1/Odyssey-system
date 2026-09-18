const integer = (value, max = 999999) => Math.min(max, Math.max(0, Math.floor(Number(value) || 0)));
const modifier = (value) => Math.min(999, Math.max(-999, Math.trunc(Number(value) || 0)));

export function sanitizeAmmunition(raw) {
  const seen = new Set();
  return (Array.isArray(raw) ? raw : []).filter((item) => {
    if (!item || typeof item.id !== "string" || !item.id || seen.has(item.id)) return false;
    seen.add(item.id);
    return true;
  }).map((item) => ({
    id: item.id,
    name: String(item.name || "Ammo").trim() || "Ammo",
    damage: integer(item.damage, 999),
    penetration: modifier(item.penetration),
    quantity: integer(item.quantity),
  }));
}

export function sanitizeMagazine(weapon) {
  const capacity = integer(weapon.capacity, 999);
  return {
    ammoIds: [...new Set(Array.isArray(weapon.ammoIds) ? weapon.ammoIds.filter((id) => typeof id === "string") : [])],
    capacity,
    loadedAmmoId: String(weapon.loadedAmmoId || ""),
    loaded: integer(weapon.loaded, capacity),
  };
}

export function reloadMagazine(odyssey, weaponIndex, ammoId) {
  const weapon = odyssey.weapons.ranged[weaponIndex];
  const ammo = odyssey.ammunition.find((item) => item.id === ammoId);
  if (!weapon || !ammo || !weapon.ammoIds.includes(ammoId)) throw new Error("Choose compatible ammunition.");
  if (weapon.capacity < 1) throw new Error("Set magazine capacity first.");
  if (ammo.quantity < 1) throw new Error("Ammo reserve is empty.");
  if (weapon.loadedAmmoId !== ammoId && weapon.loaded > 0) {
    const previous = odyssey.ammunition.find((item) => item.id === weapon.loadedAmmoId);
    if (!previous) throw new Error("Loaded ammunition is missing from the inventory.");
    if (previous.quantity + weapon.loaded > 999999) throw new Error("Ammo reserve is full.");
    previous.quantity += weapon.loaded;
    weapon.loaded = 0;
  }
  const count = Math.min(weapon.capacity - weapon.loaded, ammo.quantity);
  if (count <= 0) throw new Error("Magazine is full or ammo reserve is empty.");
  weapon.loadedAmmoId = ammoId;
  weapon.loaded += count;
  ammo.quantity -= count;
}

export function getAmmoShot(odyssey, weaponIndex, ammoId, count) {
  const weapon = odyssey.weapons.ranged[weaponIndex];
  const ammo = odyssey.ammunition.find((item) => item.id === ammoId);
  if (!weapon || !ammo || !weapon.ammoIds.includes(ammoId) || weapon.loadedAmmoId !== ammoId) {
    throw new Error("The selected ammunition must be loaded by the GM first.");
  }
  if (!Number.isInteger(count) || count < 1 || count > weapon.loaded || count > weapon.capacity) {
    throw new Error("Not enough rounds in the magazine.");
  }
  return { ammoId, count, damage: ammo.damage * count, penetration: ammo.penetration, name: ammo.name };
}

export function consumeAmmo(odyssey, weaponIndex, shot) {
  getAmmoShot(odyssey, weaponIndex, shot.ammoId, shot.count);
  odyssey.weapons.ranged[weaponIndex].loaded -= shot.count;
}

export function effectiveArmor(armor, penetration) {
  return Math.max(0, armor - penetration);
}
