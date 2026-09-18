import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { reloadMagazine, getAmmoShot, consumeAmmo, effectiveArmor, sanitizeAmmunition, sanitizeMagazine } from "./ammunition.js";

const inventory = () => ({
  ammunition: [
    { id: "ball", name: "Ball", damage: 12, penetration: 5, quantity: 20 },
    { id: "ap", name: "AP", damage: 8, penetration: 20, quantity: 4 },
  ],
  weapons: { melee: [], ranged: [{ id: "rifle", name: "Rifle", ammoIds: ["ball", "ap"], capacity: 6, loaded: 0, loadedAmmoId: "" }] },
});
const data = inventory();
reloadMagazine(data, 0, "ball");
assert.equal(data.ammunition[0].quantity, 14);
assert.equal(data.weapons.ranged[0].loaded, 6);
assert.throws(() => reloadMagazine(data, 0, "ball"));
assert.throws(() => getAmmoShot(data, 0, "ap", 1));
for (const count of [0, -1, 1.5, 7, NaN, Infinity]) assert.throws(() => getAmmoShot(data, 0, "ball", count));
const shot = getAmmoShot(data, 0, "ball", 3);
assert.equal(shot.damage, 36);
assert.equal(effectiveArmor(10, shot.penetration), 5);
assert.equal(effectiveArmor(10, 20), 0);
assert.equal(effectiveArmor(10, -5), 15);
consumeAmmo(data, 0, shot);
assert.equal(data.weapons.ranged[0].loaded, 3);
reloadMagazine(data, 0, "ap");
assert.equal(data.ammunition[0].quantity, 17);
assert.equal(data.ammunition[1].quantity, 0);
assert.equal(data.weapons.ranged[0].loaded, 4);
consumeAmmo(data, 0, getAmmoShot(data, 0, "ap", 4));
assert.throws(() => getAmmoShot(data, 0, "ap", 1));
assert.deepEqual(sanitizeAmmunition(undefined), []);
assert.equal(sanitizeMagazine({}).loaded, 0);

// Exercise the real attack handler with the scene persistence and UI boundaries mocked.
const main = readFileSync(new URL("./main.js", import.meta.url), "utf8");
const handler = main.slice(main.indexOf("let attackPending = false;"), main.indexOf("async function performRollDice"));
assert.ok(handler.includes("async function performAttackOnce"));
for (const manualDefense of [false, true]) {
  const state = { odyssey: inventory(), history: [] };
  state.odyssey.skills = { Shoot: 5 };
  state.odyssey.attributes = { Strength: 20 };
  reloadMagazine(state.odyssey, 0, "ball");
  const token = { id: "attacker" };
  const target = { id: "target" };
  let resolved;
  let calls = 0;
  const context = vm.createContext({
    structuredClone, activeTokenId: token.id, playerName: "Player", SPECIAL_PART_NAME: "Special", PARRY_SKILL_NAME: "Parry",
    getCharacterById: (id) => id === token.id ? token : target,
    canUseToken: () => true,
    getOdysseyData: () => state.odyssey,
    getTrackerData: () => ({ body: { Torso: { armor: 10, current: 5 } } }),
    getActionFieldValue: (selector) => {
      const field = selector.match(/="([^"]+)"/)[1];
      return ({targetTokenId: "target", skill: "Shoot", weaponName: "Rifle", ammoId: "ball", rounds: "3", targetPart: "Torso", manualArmor: "10", weaponDamage: "999", parryMode: "off"})[field] ?? "0";
    },
    getWeaponByName: () => ({ ...state.odyssey.weapons.ranged[0], rangedIndex: 0 }),
    getAmmoShot, consumeAmmo, effectiveArmor,
    getTargetableBodyParts: () => ["Torso"], getAutomaticTargetPenalty: () => 0,
    getParryDivisor: () => 0, saveAttackDraftValue: () => {}, persistAttackTargetToken: async () => {},
    clamp: (n, min, max) => Math.min(max, Math.max(min, n)), hasConfiguredSpecial: () => false,
    getSkillStrengthBonusFlag: () => true,
    resolveAttack: (args) => { calls++; resolved = args; return { hit: false, damage: null, outcome: "failure", targetPart: "Torso", attackTotal: 50, summary: "Miss" }; },
    getNormalizedPartState: (part) => part,
    getCharacterName: (item) => item.id,
    updateTrackerData: async (id, updater) => { if (id === token.id) Object.assign(state, updater(state)); },
    ensureOverlayForToken: async () => {}, pushDebugEntry: async () => {},
    getAttackOutcomeIcon: () => "", formatAttackDebug: () => "", setStatus: () => {}, scheduleRender: () => {},
  });
  vm.runInContext(handler, context);
  await context.performAttack({ manualDefense });
  assert.equal(calls, 1);
  assert.equal(resolved.weaponDamage, 36);
  assert.equal(resolved.targetArmor, 5);
  assert.equal(state.odyssey.weapons.ranged[0].loaded, 3, "A miss must consume rounds");
  await context.performAttack({ manualDefense });
  assert.equal(state.odyssey.weapons.ranged[0].loaded, 0);
  await assert.rejects(context.performAttack({ manualDefense }));
  assert.equal(calls, 2, "An empty magazine must not roll");
}
console.log("Ammo checks passed: reload, conservation, type switching, limits, penetration, both attack modes and miss consumption.");
