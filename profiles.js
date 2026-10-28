import {reconcileEquipment} from './characterRules.js';
import { newId } from './inventory.js';
export const LIBRARY_FORMAT = 'odyssey-template-library';
export const TEMPLATE_FORMAT = 'odyssey-character-template';
export const LIBRARY_ITEM_KEY = 'com.codex.body-hp/template-library';
export const LIBRARY_STORAGE_KEY = 'odyssey.templates.v1';
const profileFields = ['skills','skillCategories','skillStrengthBonuses','skillLabels','skillPerks','skillDescriptions','skillPerkDescriptions','skillRangeReduction','attributes','weapons','ammunition','inventory'];
export function profileData(tracker) {
  return {enabled:true, minor:tracker.minor || 0, serious:tracker.serious || 0,
    body:structuredClone(tracker.body), odyssey:{...Object.fromEntries(profileFields.map(k => [k, structuredClone(tracker.odyssey[k])])),...(tracker.odyssey.vehicle?.enabled?{vehicle:{enabled:true,pilotTokenId:''}}:{})}};
}
export function makeTemplate(name, tracker) {
  const data = profileData(tracker);
  data.minor = 0; data.serious = 0;
  for(const item of data.odyssey.inventory.items)item.criticalDamage=0;
  reconcileEquipment(data);
  for (const part of Object.values(data.body)) { part.current = part.max; part.minor = 0; part.serious = 0; }
  if(tracker.odyssey.vehicle?.enabled)data.odyssey.vehicle={enabled:true,pilotTokenId:''};
  return {format:TEMPLATE_FORMAT, version:1, id:newId(), name:String(name || 'Character').trim().slice(0,120), updatedAt:new Date().toISOString(), data};
}
export function normalizeLibrary(raw, sanitize) {
  if (raw?.format !== LIBRARY_FORMAT || raw.version !== 1 || !Array.isArray(raw.templates)) throw new Error('Choose an Odyssey template library JSON (version 1).');
  if (raw.templates.length > 100) throw new Error('A library can contain at most 100 templates.');
  const seen = new Set();
  const templates = raw.templates.map(t => {
    if (t?.format !== TEMPLATE_FORMAT || t.version !== 1 || !t.data?.body || !t.data?.odyssey) throw new Error('Invalid character template.');
    const id = String(t.id || newId()).slice(0,120);
    if (seen.has(id)) throw new Error('Duplicate template IDs in library.');
    seen.add(id);
    return {format:TEMPLATE_FORMAT, version:1, id, name:String(t.name || 'Character').trim().slice(0,120), updatedAt:String(t.updatedAt || ''), data:profileData(sanitize(t.data))};
  });
  return {format:LIBRARY_FORMAT, version:1, templates};
}
export function emptyLibrary() { return {format:LIBRARY_FORMAT,version:1,templates:[]}; }
export function mergeLibraries(local, imported) {
  const entries = new Map(local.templates.map(t => [t.id,t]));
  for (const template of imported.templates) entries.set(template.id,template);
  if (entries.size > 100) throw new Error('Library limit is 100 templates. Export or remove templates first.');
  return {...emptyLibrary(),templates:[...entries.values()]};
}
export function instantiateTemplate(current, template, sanitize) {
  const validated = normalizeLibrary({...emptyLibrary(),templates:[template]},sanitize).templates[0];
  const data = structuredClone(validated.data);
  const ammoIds = new Map(data.odyssey.ammunition.map(a => [a.id,newId()]));
  const magIds = new Map(data.odyssey.inventory.magazines.map(m => [m.id,newId()]));
  for (const a of data.odyssey.ammunition) a.id = ammoIds.get(a.id);
  for (const mag of data.odyssey.inventory.magazines) { mag.id = magIds.get(mag.id); mag.loadedAmmoId = ammoIds.get(mag.loadedAmmoId) || ''; }
  for (const w of data.odyssey.weapons.ranged) {
    if(w.acceptedMagazineIds)w.acceptedMagazineIds=w.acceptedMagazineIds.map(id=>magIds.get(id)).filter(Boolean);
    w.id = newId(); w.ammoIds = w.ammoIds.map(id => ammoIds.get(id)).filter(Boolean);
    w.loadedAmmoId = ammoIds.get(w.loadedAmmoId) || ''; w.magazineId = magIds.get(w.magazineId) || '';
  }
  for (const w of data.odyssey.weapons.melee) w.id=newId();
  for (const item of data.odyssey.inventory.items) item.id = newId();
  // Preserve destination identity, permissions and history; never inherit source identity.
  const next = sanitize({...current, ...data, odyssey:{...current.odyssey,...data.odyssey,owner:current.odyssey.owner,attackDraft:{targetTokenId:'',targetTokenName:''}}});
  return next;
}
export function readLibrary(storage, sanitize) {
  const raw = storage.getItem(LIBRARY_STORAGE_KEY);
  return raw ? normalizeLibrary(JSON.parse(raw),sanitize) : emptyLibrary();
}
export function writeLibrary(storage, library) {
  const json = JSON.stringify(library);
  if (new TextEncoder().encode(json).length > 4 * 1024 * 1024) throw new Error('Library exceeds 4 MB. Export a smaller library.');
  // Failure must be visible; do not claim success if storage is blocked/full.
  storage.setItem(LIBRARY_STORAGE_KEY,json);
}
export function libraryFromScenes(scenes, sanitize) {
  let result = emptyLibrary(); let found = false;
  for (const scene of scenes || []) for (const item of scene.items || []) {
    const payload = item.metadata?.[LIBRARY_ITEM_KEY];
    if (payload) { result = mergeLibraries(result,normalizeLibrary(payload,sanitize)); found = true; }
  }
  if (!found) throw new Error('Selected scenes contain no Odyssey template library.');
  return result;
}
