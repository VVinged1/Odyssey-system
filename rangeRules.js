// Categories are explicit: legacy weapons remain unclassified until edited by the GM.
export const WEAPON_CATEGORIES=[['','Unclassified'],['pistol','Pistol'],['smg','SMG'],['shotgun','Shotgun'],['assault','Assault rifle / machine gun'],['sniper','Sniper rifle']];
export const normalizeWeaponCategory=value=>WEAPON_CATEGORIES.some(([key])=>key===value)?value:'';
export const weaponCategoryLabel=value=>WEAPON_CATEGORIES.find(([key])=>key===value)?.[1]||'Unclassified';

export function distanceMeters(cells,scale) {
  const value=Number(scale?.parsed?.multiplier),unit=String(scale?.parsed?.unit||'').trim().toLowerCase();
  const factor=({m:1,meter:1,meters:1,metre:1,metres:1,'м':1,cm:0.01,'см':0.01,km:1000,'км':1000,ft:0.3048,foot:0.3048,feet:0.3048,'фут':0.3048,'футы':0.3048,in:0.0254,inch:0.0254,inches:0.0254,yd:0.9144,yard:0.9144,yards:0.9144,mi:1609.344,mile:1609.344,miles:1609.344})[unit];
  return Number.isFinite(cells)&&cells>=0&&Number.isFinite(value)&&value>0&&factor?cells*value*factor:null;
}
export const RANGE_BANDS=[['clinch','Clinch · 0–2 m'],['short','Short · >2–15 m'],['medium','Medium · >15–50 m'],['long','Long · >50 m']];
export function distanceBand(meters) {return !Number.isFinite(meters)||meters<0?null:meters<=2?'clinch':meters<=15?'short':meters<=50?'medium':'long';}

export function normalizeRangeReduction(raw) {
  return Object.fromEntries(RANGE_BANDS.map(([key])=>[key,Math.max(0,Math.min(999,Math.trunc(Number(raw?.[key])||0)))]).filter(([,v])=>v));
}
const MATRIX={pistol:[-15,0,-5,-10],smg:[-25,15,0,-10],shotgun:[-25,30,-15,-50],assault:[-30,0,0,0],sniper:[-50,-30,0,10]};
export function rangeAdjustment(weapon,meters,odyssey={},skill='') {
  const category=normalizeWeaponCategory(weapon?.weaponCategory),band=distanceBand(meters);
  if(!category)return {category,band,raw:0,reduction:0,modifier:0,required:false};
  if(!band)return {category,band,raw:0,reduction:0,modifier:0,required:true};
  let raw=MATRIX[category][RANGE_BANDS.findIndex(([key])=>key===band)];
  // (50,75] => 10; (75,100] => 20; …; >150 => 50.
  if(band==='long'&&['pistol','smg','sniper'].includes(category))raw=Math.sign(raw)*Math.min(50,Math.ceil((meters-50)/25)*10);
  const reduction=[weapon.rangeReduction,odyssey.skillRangeReduction?.[skill],odyssey.rangeReduction].reduce((sum,row)=>sum+(normalizeRangeReduction(row)[band]||0),0);
  return {category,band,raw,reduction,modifier:raw<0?Math.min(0,raw+reduction):raw,required:false};
}
export function rangeAdjustmentLabel(result) {
  if(!result.category)return 'Range: unclassified';
  if(result.required)return 'Range: distance in metres required';
  const band=RANGE_BANDS.find(([key])=>key===result.band)[1].split(' · ')[0],signed=n=>(n>=0?'+':'')+n;
  return band+' · '+signed(result.modifier)+(result.raw<0&&result.reduction?' (base '+result.raw+', relief '+Math.min(result.reduction,-result.raw)+')':'');
}
