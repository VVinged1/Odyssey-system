import OBR, { Command, buildImage, buildPath, isImage } from "@owlbear-rodeo/sdk";

export { OBR };

export const EXTENSION_ID = "com.codex.body-hp";
export const META_KEY = `${EXTENSION_ID}/data`;
export const OVERLAY_KEY = `${EXTENSION_ID}/overlayFor`;
export const SHIELD_PART_NAME = "Shield";
export const SPECIAL_PART_NAME = "Special";
const BODY_TOTAL_ORDER = ["Head", "L.Arm", "R.Arm", "Torso", "L.Leg", "R.Leg"];
export const BODY_ORDER = [...BODY_TOTAL_ORDER, SHIELD_PART_NAME, SPECIAL_PART_NAME];
export const ROLL_HISTORY_LIMIT = 12;
export const COMBAT_SKILL_CATEGORY = "combat";
export const APPLIED_SKILL_CATEGORY = "applied";
export const ABILITIES_SKILL_CATEGORY = "abilities";
export const MELEE_SKILL_NAME = "Melee";
export const PARRY_SKILL_NAME = "Parry";
const LEGACY_MELEE_SKILL_NAMES = new Set(["Hand", "Cold", "\u0420\u0443\u043A\u043E\u043F\u0430\u0448\u043D\u044B\u0439"]);
const LEGACY_REMOVED_SKILLS = new Set(["Hand", "Cold", "Throwing", "Rifle", "Turrets"]);
const VISUAL_VERSION = 13;
const OVERLAY_RENDER_MODE = "image";
const OVERLAY_IMAGE_KIND = "overlay-image";
const OVERLAY_STROKE_WIDTH = 0.75;
const OVERLAY_RUNTIME_CACHE = `${EXTENSION_ID}/overlay-runtime`;
const OVERLAY_RUNTIME_SW_PATH = "./overlay-runtime-sw.js";
const OVERLAY_RUNTIME_PATH_SEGMENT = "__overlay_runtime__";
const SPECIAL_RING_COLOR = "#57D8FF";
const HP_COLOR_STOPS = [
  { ratio: 1, color: "#73FF5A" },
  { ratio: 0.75, color: "#FFF243" },
  { ratio: 0.5, color: "#FFAF22" },
  { ratio: 0.25, color: "#AC0004" },
  { ratio: 0, color: "#000000" },
];
const RING_COLORS = {
  base: "#000000",
  border: "#050505",
};
const OUTER_SEGMENTS = [
  { part: "Head", angle: -90, span: 30 },
  { part: "R.Arm", angle: -18, span: 30 },
  { part: "R.Leg", angle: 54, span: 30 },
  { part: "L.Leg", angle: 126, span: 30 },
  { part: "L.Arm", angle: 198, span: 30 },
];
const FIXED_OVERLAY_KINDS = [
  "outer-base",
  ...OUTER_SEGMENTS.map((segment) => `segment-${segment.part}`),
  "torso-ring",
  "special-ring",
  "shield-ring",
];
const overlayEnsureQueue = new Map();
const overlayRuntimeUrlByTokenId = new Map();
let cachedGridDpi = null;
let overlayRuntimeReadyPromise = null;
export const DEFAULT_ODYSSEY_SKILLS = {
  [MELEE_SKILL_NAME]: 0,
  [PARRY_SKILL_NAME]: 0,
};
export const DEFAULT_ODYSSEY_SKILL_CATEGORIES = {
  [MELEE_SKILL_NAME]: COMBAT_SKILL_CATEGORY,
  [PARRY_SKILL_NAME]: COMBAT_SKILL_CATEGORY,
};
export const DEFAULT_ODYSSEY_SKILL_STRENGTH_BONUSES = {
  [MELEE_SKILL_NAME]: true,
  [PARRY_SKILL_NAME]: false,
};

export const BODY_DEFAULTS = {
  Head: { current: 1, max: 1, armor: 0, minor: 0, serious: 0 },
  "L.Arm": { current: 2, max: 2, armor: 2, minor: 0, serious: 0 },
  "R.Arm": { current: 2, max: 2, armor: 2, minor: 0, serious: 0 },
  Torso: { current: 3, max: 3, armor: 6, minor: 0, serious: 0 },
  "L.Leg": { current: 2, max: 2, armor: 2, minor: 0, serious: 0 },
  "R.Leg": { current: 2, max: 2, armor: 2, minor: 0, serious: 0 },
  [SHIELD_PART_NAME]: { current: 0, max: 0, armor: 0, minor: 0, serious: 0 },
  [SPECIAL_PART_NAME]: { current: 0, max: 0, armor: 0, minor: 0, serious: 0 },
};

export const DEFAULT_TRACKER_DATA = {
  enabled: true,
  minor: 0,
  serious: 0,
  body: structuredClone(BODY_DEFAULTS),
  identity: {
    playerId: "",
    characterId: "",
  },
  lastRoll: null,
  history: [],
  sync: {
    lastEventId: 0,
    lastSyncedAt: null,
  },
  odyssey: {
    owner: {
      playerId: "",
      playerName: "",
    },
    attackDraft: {
      targetTokenId: "",
      targetTokenName: "",
    },
    skills: structuredClone(DEFAULT_ODYSSEY_SKILLS),
    skillCategories: structuredClone(DEFAULT_ODYSSEY_SKILL_CATEGORIES),
    skillStrengthBonuses: structuredClone(DEFAULT_ODYSSEY_SKILL_STRENGTH_BONUSES),
    attributes: {
      Strength: 0,
      Agility: 0,
      Reaction: 0,
      Endurance: 0,
      Perception: 0,
      Intelligence: 0,
      Charisma: 0,
      Willpower: 0,
      Magic: 0,
    },
    weapons: {
      melee: [],
      ranged: [],
    },
  },
};

export function deepClone(value) {
  return structuredClone(value);
}

export function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function numberOrFallback(value, fallback) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : fallback;
}

export function sanitizeTrackerData(raw) {
  const next = deepClone(DEFAULT_TRACKER_DATA);
  if (!raw || typeof raw !== "object") return next;

  next.enabled = raw.enabled !== false;
  next.minor = clamp(Number(raw.minor ?? 0) || 0, 0, 4);
  next.serious = clamp(Number(raw.serious ?? 0) || 0, 0, 2);
  next.identity.playerId = String(raw.identity?.playerId ?? "").trim();
  next.identity.characterId = String(raw.identity?.characterId ?? "").trim();
  next.lastRoll = sanitizeRollSummary(raw.lastRoll);
  next.history = Array.isArray(raw.history)
    ? raw.history.map(sanitizeRollSummary).filter(Boolean).slice(0, ROLL_HISTORY_LIMIT)
    : [];
  next.sync.lastEventId = Math.max(0, Number(raw.sync?.lastEventId ?? 0) || 0);
  next.sync.lastSyncedAt = raw.sync?.lastSyncedAt
    ? String(raw.sync.lastSyncedAt)
    : null;
  next.odyssey = sanitizeOdysseyData(raw.odyssey);

  for (const partName of BODY_ORDER) {
    const source = raw.body?.[partName] ?? {};
    const part = next.body[partName];
    part.max = clamp(numberOrFallback(source.max, part.max), 0, 99);
    part.current = clamp(
      numberOrFallback(source.current, part.current),
      0,
      part.max,
    );
    part.armor = clamp(numberOrFallback(source.armor, part.armor), 0, 99);
    part.minor = clamp(numberOrFallback(source.minor, part.minor), 0, 3);
    part.serious = clamp(numberOrFallback(source.serious, part.serious), 0, 1);
  }

  return next;
}

export function sanitizeOdysseyData(raw) {
  const next = deepClone(DEFAULT_TRACKER_DATA.odyssey);
  if (!raw || typeof raw !== "object") return next;

  next.owner.playerId = String(raw.owner?.playerId ?? "").trim();
  next.owner.playerName = String(raw.owner?.playerName ?? "").trim();
  next.attackDraft.targetTokenId = String(raw.attackDraft?.targetTokenId ?? "").trim();
  next.attackDraft.targetTokenName = String(raw.attackDraft?.targetTokenName ?? "").trim();

  const rawSkills = raw.skills && typeof raw.skills === "object" ? raw.skills : {};
  const rawSkillCategories =
    raw.skillCategories && typeof raw.skillCategories === "object"
      ? raw.skillCategories
      : {};
  const rawSkillStrengthBonuses =
    raw.skillStrengthBonuses && typeof raw.skillStrengthBonuses === "object"
      ? raw.skillStrengthBonuses
      : {};

  const migratedMeleeValue = Math.max(
    Number(rawSkills[MELEE_SKILL_NAME] ?? 0) || 0,
    ...Array.from(LEGACY_MELEE_SKILL_NAMES).map((skillName) => Number(rawSkills[skillName] ?? 0) || 0),
    Number(DEFAULT_ODYSSEY_SKILLS[MELEE_SKILL_NAME] ?? 0) || 0,
  );
  const migratedParryValue = Math.max(
    Number(rawSkills[PARRY_SKILL_NAME] ?? 0) || 0,
    Number(raw.attributes?.Parry ?? 0) || 0,
    Number(DEFAULT_ODYSSEY_SKILLS[PARRY_SKILL_NAME] ?? 0) || 0,
  );

  next.skills[MELEE_SKILL_NAME] = clamp(migratedMeleeValue, 0, 10);
  next.skillCategories[MELEE_SKILL_NAME] = COMBAT_SKILL_CATEGORY;
  next.skillStrengthBonuses[MELEE_SKILL_NAME] = true;
  next.skills[PARRY_SKILL_NAME] = clamp(migratedParryValue, 0, 10);
  next.skillCategories[PARRY_SKILL_NAME] = COMBAT_SKILL_CATEGORY;
  next.skillStrengthBonuses[PARRY_SKILL_NAME] = false;

  for (const [key, value] of Object.entries(rawSkills)) {
    const normalizedKey = String(key).trim();
    if (!normalizedKey) continue;
    if (
      normalizedKey === MELEE_SKILL_NAME ||
      normalizedKey === PARRY_SKILL_NAME ||
      LEGACY_MELEE_SKILL_NAMES.has(normalizedKey) ||
      LEGACY_REMOVED_SKILLS.has(normalizedKey)
    ) {
      continue;
    }

    next.skills[normalizedKey] = clamp(Number(value) || 0, 0, 10);
    const categoryValue = String(
      rawSkillCategories[normalizedKey] ?? rawSkillCategories[key] ?? "",
    ).toLowerCase();
    next.skillCategories[normalizedKey] =
      categoryValue === COMBAT_SKILL_CATEGORY
        ? COMBAT_SKILL_CATEGORY
        : categoryValue === ABILITIES_SKILL_CATEGORY
          ? ABILITIES_SKILL_CATEGORY
          : APPLIED_SKILL_CATEGORY;
    next.skillStrengthBonuses[normalizedKey] = Boolean(
      rawSkillStrengthBonuses[normalizedKey] ?? rawSkillStrengthBonuses[key] ?? false,
    );
  }

  for (const key of Object.keys(next.attributes)) {
    const fallbackValue =
      key === "Magic"
        ? raw.attributes?.[key] ?? raw.attributes?.Psionics ?? 0
        : raw.attributes?.[key] ?? 0;
    next.attributes[key] = clamp(Number(fallbackValue) || 0, 0, 20);
  }

  next.weapons.melee = sanitizeWeapons(raw.weapons?.melee);
  next.weapons.ranged = sanitizeWeapons(raw.weapons?.ranged);

  return next;
}

function sanitizeWeapons(raw) {
  if (!Array.isArray(raw)) return [];
  return raw
    .filter((item) => item && typeof item === "object")
    .map((item) => ({
      name: String(item.name ?? "").trim() || "Weapon",
      damage: clamp(Number(item.damage ?? 0) || 0, -99, 99),
    }))
    .slice(0, 20);
}

export function sanitizeRollSummary(raw) {
  if (!raw || typeof raw !== "object") return null;
  const eventId = Math.max(0, Number(raw.eventId ?? 0) || 0);
  const summary = String(raw.summary ?? "").trim();
  const actorName = String(raw.actorName ?? "").trim();
  const outcome = String(raw.outcome ?? "").trim();
  const total = raw.total == null ? null : Number(raw.total) || 0;
  const targetPart = String(raw.targetPart ?? "").trim();
  const timestamp = raw.timestamp ? String(raw.timestamp) : null;
  const source = String(raw.source ?? "bridge").trim();

  if (!summary && !actorName && total == null && !outcome) {
    return null;
  }

  return {
    eventId,
    summary,
    actorName,
    outcome,
    total,
    targetPart,
    timestamp,
    source,
  };
}

export function getTrackerData(item) {
  return sanitizeTrackerData(item?.metadata?.[META_KEY]);
}

export function isCharacterToken(item) {
  return Boolean(item) && isImage(item) && item.layer === "CHARACTER";
}

export function isTrackedCharacter(item) {
  return isCharacterToken(item) && item.metadata?.[META_KEY]?.enabled === true;
}

export function isOverlayItem(item) {
  return Boolean(item?.metadata?.[OVERLAY_KEY]);
}

export function getCharacterName(item) {
  if (!item) return "Unnamed character";
  const byName = typeof item.name === "string" ? item.name.trim() : "";
  if (byName) return byName;
  return `Character ${item.id.slice(0, 6)}`;
}

export function sortCharacters(items) {
  return [...items].sort((left, right) =>
    getCharacterName(left).localeCompare(getCharacterName(right)),
  );
}

export function formatOverlayText(data) {
  const body = data.body;
  const lines = [];

  if (hasConfiguredShield(data)) {
    lines.push(
      `${SHIELD_PART_NAME} ${body[SHIELD_PART_NAME].current}/${body[SHIELD_PART_NAME].max}(${body[SHIELD_PART_NAME].armor})`,
    );
  }

  if (hasConfiguredSpecial(data)) {
    lines.push(
      `${SPECIAL_PART_NAME} ${body[SPECIAL_PART_NAME].current}/${body[SPECIAL_PART_NAME].max}(${body[SPECIAL_PART_NAME].armor})`,
    );
  }

  lines.push(
    `Head ${body["Head"].current}/${body["Head"].max}(${body["Head"].armor}) | L.Arm ${body["L.Arm"].current}/${body["L.Arm"].max}(${body["L.Arm"].armor}) | R.Arm ${body["R.Arm"].current}/${body["R.Arm"].max}(${body["R.Arm"].armor})`,
  );
  lines.push(
    `Torso ${body["Torso"].current}/${body["Torso"].max}(${body["Torso"].armor}) | L.Leg ${body["L.Leg"].current}/${body["L.Leg"].max}(${body["L.Leg"].armor}) | R.Leg ${body["R.Leg"].current}/${body["R.Leg"].max}(${body["R.Leg"].armor})`,
  );

  return lines.join("\n");
}

export function formatLastRoll(lastRoll) {
  const parts = [];
  if (lastRoll.actorName) parts.push(lastRoll.actorName);
  if (lastRoll.total != null) parts.push(`roll ${lastRoll.total}`);
  if (lastRoll.outcome) parts.push(lastRoll.outcome);
  if (lastRoll.targetPart) parts.push(`target ${lastRoll.targetPart}`);
  if (lastRoll.summary) parts.push(lastRoll.summary);
  return parts.join(" | ");
}

export function getOdysseyData(item) {
  return sanitizeOdysseyData(getTrackerData(item).odyssey);
}

export function canPlayerControlToken(playerRole, playerId, token) {
  if (!token || !isCharacterToken(token)) return false;
  if (playerRole === "GM") return true;
  const odyssey = getOdysseyData(token);
  return Boolean(playerId) && odyssey.owner.playerId === playerId;
}

export function getAvailableWeapons(token, mode = "melee") {
  const odyssey = getOdysseyData(token);
  const list = mode === "ranged" ? odyssey.weapons.ranged : odyssey.weapons.melee;
  return list.length ? list : [{ name: "Default", damage: 0 }];
}

export function getBodyTotals(data) {
  return BODY_TOTAL_ORDER.reduce(
    (accumulator, partName) => {
      accumulator.current += data.body[partName].current;
      accumulator.max += data.body[partName].max;
      return accumulator;
    },
    { current: 0, max: 0 },
  );
}

export function getArmorTotal(dataOrBody) {
  const body = dataOrBody?.body ?? dataOrBody ?? {};
  return BODY_ORDER.reduce(
    (total, partName) => total + (Number(body?.[partName]?.armor) || 0),
    0,
  );
}

export function hasConfiguredShield(dataOrBody) {
  const body = dataOrBody?.body ?? dataOrBody;
  const shield = body?.[SHIELD_PART_NAME];
  if (!shield || typeof shield !== "object") return false;

  return (
    (Number(shield.max) || 0) > 0 ||
    (Number(shield.current) || 0) > 0 ||
    (Number(shield.armor) || 0) > 0 ||
    (Number(shield.minor) || 0) > 0 ||
    (Number(shield.serious) || 0) > 0
  );
}

export function hasConfiguredSpecial(dataOrBody) {
  const body = dataOrBody?.body ?? dataOrBody;
  const special = body?.[SPECIAL_PART_NAME];
  if (!special || typeof special !== "object") return false;

  return (
    (Number(special.max) || 0) > 0 ||
    (Number(special.current) || 0) > 0 ||
    (Number(special.armor) || 0) > 0 ||
    (Number(special.minor) || 0) > 0 ||
    (Number(special.serious) || 0) > 0
  );
}

export function getTargetableBodyParts(dataOrBody) {
  return BODY_ORDER.filter(
    (partName) =>
      partName !== SPECIAL_PART_NAME &&
      (partName !== SHIELD_PART_NAME || hasConfiguredShield(dataOrBody)),
  );
}

function getEffectiveSize(token) {
  const scaleX = Math.abs(token.scale?.x ?? 1);
  const scaleY = Math.abs(token.scale?.y ?? 1);
  return {
    width: (token.width || 140) * scaleX,
    height: (token.height || 140) * scaleY,
  };
}

async function getTokenMetrics(token) {
  const effectiveSize = getEffectiveSize(token);
  const center = token.position;
  const width = effectiveSize.width;
  const height = effectiveSize.height;

  const gridDpi = await getCachedGridDpi();

  const scaleFactor = Math.max(
    Math.abs(token.scale?.x ?? 1),
    Math.abs(token.scale?.y ?? 1),
    1,
  );
  const visibleDiameter = Math.max(
    width,
    height,
    effectiveSize.width,
    effectiveSize.height,
    gridDpi * scaleFactor,
    56,
  );
  const tokenRadius = visibleDiameter / 2;
  const tokenGap = 0;
  const torsoThickness = Math.max(5, visibleDiameter * 0.035);
  const torsoInnerRadius = tokenRadius + tokenGap;
  const torsoOuterRadius = torsoInnerRadius + torsoThickness;
  const ringGap = 0;
  const outerThickness = Math.max(8, visibleDiameter * 0.08);
  const outerInnerRadius = torsoOuterRadius + ringGap;
  const outerRadius = outerInnerRadius + outerThickness;
  const specialThickness = Math.max(4, visibleDiameter * 0.03);
  const specialInnerRadius = outerRadius;
  const specialOuterRadius = specialInnerRadius + specialThickness;
  const shieldThickness = Math.max(4, visibleDiameter * 0.028);
  const shieldOuterRadius = Math.max(10, visibleDiameter * 0.1);
  const shieldInnerRadius = Math.max(4, shieldOuterRadius - shieldThickness);
  const shieldOffsetY = -(specialOuterRadius + shieldOuterRadius + Math.max(5, visibleDiameter * 0.035));

  return {
    center,
    visibleDiameter,
    outerRadius,
    outerInnerRadius,
    torsoOuterRadius,
    torsoInnerRadius,
    specialOuterRadius,
    specialInnerRadius,
    shieldOuterRadius,
    shieldInnerRadius,
    shieldOffsetY,
  };
}

async function getCachedGridDpi(forceRefresh = false) {
  if (!forceRefresh && Number.isFinite(cachedGridDpi) && cachedGridDpi > 0) {
    return cachedGridDpi;
  }

  let gridDpi = 150;
  try {
    gridDpi = (await OBR.scene.grid.getDpi()) || gridDpi;
  } catch (error) {
    console.warn("[Body HP] Unable to read grid dpi, using fallback size", error);
  }

  cachedGridDpi = Math.max(1, Number(gridDpi) || 150);
  return cachedGridDpi;
}

function polar(radius, angle) {
  const radians = (angle * Math.PI) / 180;
  return {
    x: radius * Math.cos(radians),
    y: radius * Math.sin(radians),
  };
}

function arcPoints(radius, startAngle, endAngle, segments = 18) {
  const points = [];
  for (let index = 0; index <= segments; index += 1) {
    const ratio = index / segments;
    const angle = startAngle + (endAngle - startAngle) * ratio;
    points.push(polar(radius, angle));
  }
  return points;
}

function buildAnnulusCommands(radiusOuter, radiusInner, offsetX = 0, offsetY = 0) {
  const outer = arcPoints(radiusOuter, -180, 180, 36).map((point) => ({
    x: point.x + offsetX,
    y: point.y + offsetY,
  }));
  const inner = arcPoints(radiusInner, -180, 180, 36).map((point) => ({
    x: point.x + offsetX,
    y: point.y + offsetY,
  }));
  const commands = [[Command.MOVE, outer[0].x, outer[0].y]];

  for (const point of outer.slice(1)) {
    commands.push([Command.LINE, point.x, point.y]);
  }

  commands.push([Command.CLOSE]);
  commands.push([Command.MOVE, inner[0].x, inner[0].y]);

  for (const point of inner) {
    commands.push([Command.LINE, point.x, point.y]);
  }

  commands.push([Command.CLOSE]);
  return commands;
}

function buildSectorCommands(radiusOuter, radiusInner, centerAngle, spanAngle) {
  const startAngle = centerAngle - spanAngle / 2;
  const endAngle = centerAngle + spanAngle / 2;
  const outer = arcPoints(radiusOuter, startAngle, endAngle, 10);
  const inner = arcPoints(radiusInner, endAngle, startAngle, 10);
  const commands = [[Command.MOVE, outer[0].x, outer[0].y]];

  for (const point of outer.slice(1)) {
    commands.push([Command.LINE, point.x, point.y]);
  }

  for (const point of inner) {
    commands.push([Command.LINE, point.x, point.y]);
  }

  commands.push([Command.CLOSE]);
  return commands;
}

function hexToRgb(hex) {
  const normalized = String(hex).replace("#", "");
  const value = normalized.length === 3
    ? normalized
        .split("")
        .map((char) => `${char}${char}`)
        .join("")
    : normalized;

  return {
    r: Number.parseInt(value.slice(0, 2), 16) || 0,
    g: Number.parseInt(value.slice(2, 4), 16) || 0,
    b: Number.parseInt(value.slice(4, 6), 16) || 0,
  };
}

function rgbToHex({ r, g, b }) {
  return `#${[r, g, b]
    .map((channel) =>
      clamp(Math.round(channel), 0, 255).toString(16).padStart(2, "0").toUpperCase(),
    )
    .join("")}`;
}

function mixHexColors(startHex, endHex, ratio) {
  const safeRatio = clamp(ratio, 0, 1);
  const start = hexToRgb(startHex);
  const end = hexToRgb(endHex);

  return rgbToHex({
    r: start.r + (end.r - start.r) * safeRatio,
    g: start.g + (end.g - start.g) * safeRatio,
    b: start.b + (end.b - start.b) * safeRatio,
  });
}

function getHpColor(ratio) {
  const safeRatio = clamp(ratio, 0, 1);

  for (let index = 0; index < HP_COLOR_STOPS.length - 1; index += 1) {
    const upper = HP_COLOR_STOPS[index];
    const lower = HP_COLOR_STOPS[index + 1];

    if (safeRatio > upper.ratio || safeRatio < lower.ratio) {
      continue;
    }

    const span = upper.ratio - lower.ratio;
    if (span <= 0) return upper.color;
    const progress = (safeRatio - lower.ratio) / span;
    return mixHexColors(lower.color, upper.color, progress);
  }

  return HP_COLOR_STOPS.at(-1)?.color ?? RING_COLORS.base;
}

function getPartColor(part) {
  if (part.max <= 0) {
    return (Number(part?.armor) || 0) > 0 ? getHpColor(1) : getHpColor(0);
  }
  return getHpColor(part.current / part.max);
}

function getSpecialPartColor(part) {
  const ratio =
    (Number(part?.max) || 0) > 0
      ? clamp((Number(part?.current) || 0) / (Number(part?.max) || 1), 0, 1)
      : (Number(part?.current) || 0) > 0 || (Number(part?.armor) || 0) > 0
        ? 1
        : 0;
  return mixHexColors("#000000", SPECIAL_RING_COLOR, ratio);
}

function buildRingItem(
  token,
  metrics,
  kind,
  commands,
  fillColor,
  zIndex = 0,
  fillRule = "nonzero",
  signature = "",
  itemVisible = true,
) {
  return buildPath()
    .name(`${kind}: ${getCharacterName(token)}`)
    .commands(commands)
    .fillRule(fillRule)
    .fillColor(fillColor)
    .fillOpacity(1)
    .strokeColor(RING_COLORS.border)
    .strokeOpacity(1)
    .strokeWidth(0.75)
    .position(metrics.center)
    .rotation(0)
    .zIndex((token.zIndex ?? 0) + 100 + zIndex)
    .visible(itemVisible && token.visible !== false)
    .attachedTo(token.id)
    .disableAttachmentBehavior(["ROTATION"])
    .layer("ATTACHMENT")
    .locked(true)
    .disableHit(true)
    .metadata({
      [OVERLAY_KEY]: token.id,
      kind,
      visualVersion: VISUAL_VERSION,
      signature,
    })
    .build();
}

function applyOverlayItemState(target, source) {
  if (isImage(target) && isImage(source)) {
    target.image = source.image;
    target.grid = source.grid;
    target.position = source.position;
    target.rotation = source.rotation;
    target.scale = source.scale;
    target.zIndex = source.zIndex;
    target.visible = source.visible;
    target.attachedTo = source.attachedTo;
    target.disableAttachmentBehavior = source.disableAttachmentBehavior;
    target.layer = source.layer;
    target.locked = source.locked;
    target.disableHit = source.disableHit;
    target.metadata = {
      ...(target.metadata ?? {}),
      ...(source.metadata ?? {}),
    };
    return;
  }

  target.name = source.name;
  target.commands = source.commands;
  target.fillRule = source.fillRule;
  target.fillColor = source.fillColor;
  target.fillOpacity = source.fillOpacity;
  target.strokeColor = source.strokeColor;
  target.strokeOpacity = source.strokeOpacity;
  target.strokeWidth = source.strokeWidth;
  target.position = source.position;
  target.rotation = source.rotation;
  target.zIndex = source.zIndex;
  target.visible = source.visible;
  target.metadata = {
    ...(target.metadata ?? {}),
    ...(source.metadata ?? {}),
  };
}

function hasPatchableOverlaySet(token, overlayItems, expectedKinds) {
  if (OVERLAY_RENDER_MODE === "image") {
    if (overlayItems.length !== 1) {
      return false;
    }

    const [item] = overlayItems;
    return (
      isImage(item) &&
      item.attachedTo === token.id &&
      item.visible === true &&
      Number(item.metadata?.visualVersion ?? 0) === VISUAL_VERSION &&
      String(item.metadata?.kind ?? "") === OVERLAY_IMAGE_KIND
    );
  }

  if (overlayItems.length !== expectedKinds.length) {
    return false;
  }

  const seenKinds = new Set();
  return overlayItems.every((item) => {
    const kind = String(item.metadata?.kind ?? "");
    const valid =
      item.attachedTo === token.id &&
      Number(item.metadata?.visualVersion ?? 0) === VISUAL_VERSION &&
      expectedKinds.includes(kind) &&
      !seenKinds.has(kind);
    seenKinds.add(kind);
    return valid;
  });
}

function roundMetric(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function commandsToSvgPath(commands) {
  return commands
    .map((command) => {
      const [type, x = 0, y = 0] = command;
      if (type === Command.MOVE) {
        return `M ${roundMetric(x)} ${roundMetric(y)}`;
      }
      if (type === Command.LINE) {
        return `L ${roundMetric(x)} ${roundMetric(y)}`;
      }
      if (type === Command.CLOSE) {
        return "Z";
      }
      return "";
    })
    .filter(Boolean)
    .join(" ");
}

function buildOverlayBounds(metrics, data) {
  const specialActive = hasConfiguredSpecial(data);
  const shieldActive = hasConfiguredShield(data);
  const ringRadius = specialActive ? metrics.specialOuterRadius : metrics.outerRadius;
  const horizontalExtent = Math.max(
    ringRadius,
    shieldActive ? metrics.shieldOuterRadius : 0,
  );
  const topExtent = Math.max(
    ringRadius,
    shieldActive ? Math.abs(metrics.shieldOffsetY) + metrics.shieldOuterRadius : 0,
  );
  const bottomExtent = ringRadius;
  const padding = Math.max(2, metrics.visibleDiameter * 0.02);

  return {
    minX: -horizontalExtent - padding,
    maxX: horizontalExtent + padding,
    minY: -topExtent - padding,
    maxY: bottomExtent + padding,
  };
}

function buildOverlaySignature(token, data, metrics) {
  const bodySignature = BODY_ORDER.map((partName) => {
    const part = data.body?.[partName] ?? {};
    return [
      partName,
      Number(part.current) || 0,
      Number(part.max) || 0,
      Number(part.armor) || 0,
      Number(part.minor) || 0,
      Number(part.serious) || 0,
    ].join(":");
  }).join("|");

  return [
    VISUAL_VERSION,
    roundMetric(metrics.visibleDiameter),
    roundMetric(metrics.outerRadius),
    roundMetric(metrics.outerInnerRadius),
    roundMetric(metrics.torsoOuterRadius),
    roundMetric(metrics.torsoInnerRadius),
    roundMetric(metrics.specialOuterRadius),
    roundMetric(metrics.specialInnerRadius),
    roundMetric(metrics.shieldOuterRadius),
    roundMetric(metrics.shieldInnerRadius),
    roundMetric(metrics.shieldOffsetY),
    bodySignature,
  ].join(";");
}

function buildOverlaySvgMarkup(token, data, metrics) {
  const layers = [
    {
      d: commandsToSvgPath(buildAnnulusCommands(metrics.outerRadius, metrics.outerInnerRadius)),
      fill: RING_COLORS.base,
      fillRule: "evenodd",
    },
    ...OUTER_SEGMENTS.map((segment) => ({
      d: commandsToSvgPath(
        buildSectorCommands(
          metrics.outerRadius,
          metrics.outerInnerRadius,
          segment.angle,
          segment.span,
        ),
      ),
      fill: getPartColor(data.body[segment.part]),
      fillRule: "nonzero",
    })),
    {
      d: commandsToSvgPath(
        buildAnnulusCommands(metrics.torsoOuterRadius, metrics.torsoInnerRadius),
      ),
      fill: getPartColor(data.body.Torso),
      fillRule: "evenodd",
    },
  ];

  if (hasConfiguredSpecial(data)) {
    layers.push({
      d: commandsToSvgPath(
        buildAnnulusCommands(metrics.specialOuterRadius, metrics.specialInnerRadius),
      ),
      fill: getSpecialPartColor(data.body[SPECIAL_PART_NAME]),
      fillRule: "evenodd",
    });
  }

  if (hasConfiguredShield(data)) {
    layers.push({
      d: commandsToSvgPath(
        buildAnnulusCommands(
          metrics.shieldOuterRadius,
          metrics.shieldInnerRadius,
          0,
          metrics.shieldOffsetY,
        ),
      ),
      fill: getPartColor(data.body[SHIELD_PART_NAME]),
      fillRule: "evenodd",
    });
  }

  const bounds = buildOverlayBounds(metrics, data);
  const width = Math.max(1, roundMetric(bounds.maxX - bounds.minX));
  const height = Math.max(1, roundMetric(bounds.maxY - bounds.minY));

  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${roundMetric(bounds.minX)} ${roundMetric(bounds.minY)} ${width} ${height}" width="${width}" height="${height}">${layers
    .map(
      (layer) =>
        `<path d="${layer.d}" fill="${layer.fill}" fill-rule="${layer.fillRule}" stroke="${RING_COLORS.border}" stroke-width="${OVERLAY_STROKE_WIDTH}" stroke-opacity="1" vector-effect="non-scaling-stroke"/>`,
    )
    .join("")}</svg>`;

  return {
    svg,
    width,
    height,
    signature: buildOverlaySignature(token, data, metrics),
  };
}

function hashOverlaySignature(signature) {
  let hash = 0;
  for (let index = 0; index < signature.length; index += 1) {
    hash = (hash * 31 + signature.charCodeAt(index)) | 0;
  }
  return Math.abs(hash).toString(36);
}

function buildOverlayRuntimeUrl(tokenId, signature) {
  const runtimePath = `${OVERLAY_RUNTIME_PATH_SEGMENT}/${encodeURIComponent(tokenId)}-${hashOverlaySignature(signature)}.svg`;
  return new URL(`./${runtimePath}`, window.location.href).href;
}

async function cacheOverlaySvg(url, svg) {
  if (!("caches" in globalThis)) return false;
  const cache = await caches.open(OVERLAY_RUNTIME_CACHE);
  await cache.put(
    url,
    new Response(svg, {
      headers: {
        "Content-Type": "image/svg+xml",
        "Cache-Control": "no-store",
      },
    }),
  );
  return true;
}

async function deleteCachedOverlaySvg(url) {
  if (!url || !("caches" in globalThis)) return;
  const cache = await caches.open(OVERLAY_RUNTIME_CACHE);
  await cache.delete(url);
}

export async function ensureOverlayRuntimeReady() {
  if (OVERLAY_RENDER_MODE !== "image") return false;
  if (typeof window === "undefined") return false;
  if (!("serviceWorker" in navigator) || !("caches" in globalThis)) return false;

  if (!overlayRuntimeReadyPromise) {
    overlayRuntimeReadyPromise = (async () => {
      try {
        const swUrl = new URL(`${OVERLAY_RUNTIME_SW_PATH}?v=${VISUAL_VERSION}`, window.location.href);
        await navigator.serviceWorker.register(swUrl, {
          scope: new URL("./", window.location.href).pathname,
        });
        await navigator.serviceWorker.ready;
        return true;
      } catch (error) {
        console.warn("[Body HP] Overlay runtime registration failed", error);
        return false;
      }
    })();
  }

  return overlayRuntimeReadyPromise;
}

function encodeSvgDataUrl(svg) {
  return `data:image/svg+xml;base64,${btoa(svg)}`;
}

async function buildOverlayImageItem(token, data, metrics) {
  const { svg, width, height, signature } = buildOverlaySvgMarkup(token, data, metrics);
  const bounds = buildOverlayBounds(metrics, data);
  const dpi = await getCachedGridDpi();
  const runtimeReady = await ensureOverlayRuntimeReady();
  let url = encodeSvgDataUrl(svg);

  if (runtimeReady) {
    const runtimeUrl = buildOverlayRuntimeUrl(token.id, signature);
    await cacheOverlaySvg(runtimeUrl, svg);
    const previousUrl = overlayRuntimeUrlByTokenId.get(token.id);
    if (previousUrl && previousUrl !== runtimeUrl) {
      await deleteCachedOverlaySvg(previousUrl);
    }
    overlayRuntimeUrlByTokenId.set(token.id, runtimeUrl);
    url = runtimeUrl;
  }

  const image = {
    width: Math.max(1, Math.ceil(width)),
    height: Math.max(1, Math.ceil(height)),
    mime: "image/svg+xml",
    url,
  };
  const grid = {
    dpi,
    offset: {
      x: Math.max(0, roundMetric(-bounds.minX)),
      y: Math.max(0, roundMetric(-bounds.minY)),
    },
  };

  return buildImage(image, grid)
    .name(`Overlay: ${getCharacterName(token)}`)
    .position(metrics.center)
    .rotation(0)
    .scale({ x: 1, y: 1 })
    .zIndex((token.zIndex ?? 0) + 100)
    .visible(token.visible !== false)
    .attachedTo(token.id)
    .disableAttachmentBehavior(["ROTATION"])
    .disableAutoZIndex(true)
    .layer("ATTACHMENT")
    .locked(true)
    .disableHit(true)
    .metadata({
      [OVERLAY_KEY]: token.id,
      kind: OVERLAY_IMAGE_KIND,
      visualVersion: VISUAL_VERSION,
      signature,
    })
    .build();
}

function applyBodyEffects(body, bodyEffects) {
  if (!bodyEffects || typeof bodyEffects !== "object") return;

  for (const partName of BODY_ORDER) {
    const patch = bodyEffects[partName];
    if (!patch || typeof patch !== "object") continue;
    const part = body[partName];

    if (patch.max != null || patch.max_delta != null) {
      const baseMax = patch.max != null ? Number(patch.max) || 0 : part.max;
      const deltaMax = Number(patch.max_delta) || 0;
      part.max = clamp(baseMax + deltaMax, 0, 99);
    }

    if (patch.current != null || patch.current_delta != null) {
      const baseCurrent =
        patch.current != null ? Number(patch.current) || 0 : part.current;
      const deltaCurrent = Number(patch.current_delta) || 0;
      part.current = clamp(baseCurrent + deltaCurrent, 0, part.max);
    } else {
      part.current = clamp(part.current, 0, part.max);
    }

    if (patch.armor != null || patch.armor_delta != null) {
      const baseArmor = patch.armor != null ? Number(patch.armor) || 0 : part.armor;
      const deltaArmor = Number(patch.armor_delta) || 0;
      part.armor = clamp(baseArmor + deltaArmor, 0, 99);
    }

    if (patch.minor != null || patch.minor_delta != null) {
      const baseMinor = patch.minor != null ? Number(patch.minor) || 0 : part.minor;
      const deltaMinor = Number(patch.minor_delta) || 0;
      part.minor = clamp(baseMinor + deltaMinor, 0, 3);
    }

    if (patch.serious != null || patch.serious_delta != null) {
      const baseSerious = patch.serious != null ? Number(patch.serious) || 0 : part.serious;
      const deltaSerious = Number(patch.serious_delta) || 0;
      part.serious = clamp(baseSerious + deltaSerious, 0, 1);
    }
  }
}

export function applyRollEventToData(current, event) {
  const next = sanitizeTrackerData(current);
  const payload = event?.payload ?? {};
  const effects = payload.effects ?? {};
  const rollSummary = sanitizeRollSummary({
    eventId: event?.id,
    actorName: event?.actor_name ?? payload.actor_name,
    summary: event?.summary ?? payload.summary,
    total: payload.total,
    outcome: payload.outcome,
    targetPart: payload.target_part,
    timestamp: event?.created_at,
    source: "bridge",
  });

  next.minor = clamp(
    effects.minor != null
      ? Number(effects.minor) || 0
      : next.minor + (Number(effects.minor_delta) || 0),
    0,
    4,
  );
  next.serious = clamp(
    effects.serious != null
      ? Number(effects.serious) || 0
      : next.serious + (Number(effects.serious_delta) || 0),
    0,
    2,
  );

  applyBodyEffects(next.body, effects.body);

  if (effects.identity && typeof effects.identity === "object") {
    next.identity.playerId = String(
      effects.identity.player_id ?? next.identity.playerId,
    ).trim();
    next.identity.characterId = String(
      effects.identity.character_id ?? next.identity.characterId,
    ).trim();
  }

  next.sync.lastEventId = Math.max(next.sync.lastEventId, Number(event?.id) || 0);
  next.sync.lastSyncedAt = event?.created_at ?? new Date().toISOString();

  if (rollSummary) {
    next.lastRoll = rollSummary;
    next.history = [rollSummary, ...next.history].slice(0, ROLL_HISTORY_LIMIT);
  }

  return next;
}

export async function updateTrackerData(tokenId, updater) {
  await OBR.scene.items.updateItems([tokenId], (items) => {
    const token = items[0];
    if (!token) return;
    token.metadata ??= {};
    token.metadata[META_KEY] = sanitizeTrackerData(
      updater(getTrackerData(token)),
    );
  });
}

export function buildOverlayItems(token, data, metrics, signature = "") {
  const items = [];
  const specialVisible = hasConfiguredSpecial(data);
  const shieldVisible = hasConfiguredShield(data);

  items.push(
    buildRingItem(
      token,
      metrics,
      "outer-base",
      buildAnnulusCommands(metrics.outerRadius, metrics.outerInnerRadius),
        RING_COLORS.base,
        0,
        "evenodd",
        signature,
        true,
      ),
    );

  for (const segment of OUTER_SEGMENTS) {
    items.push(
      buildRingItem(
        token,
        metrics,
        `segment-${segment.part}`,
        buildSectorCommands(
          metrics.outerRadius,
          metrics.outerInnerRadius,
          segment.angle,
          segment.span,
        ),
        getPartColor(data.body[segment.part]),
        1,
        "nonzero",
        signature,
        true,
      ),
    );
  }

  items.push(
    buildRingItem(
      token,
      metrics,
      "torso-ring",
      buildAnnulusCommands(metrics.torsoOuterRadius, metrics.torsoInnerRadius),
      getPartColor(data.body.Torso),
      2,
      "evenodd",
      signature,
      true,
    ),
  );

  items.push(
    buildRingItem(
      token,
      metrics,
      "special-ring",
      buildAnnulusCommands(metrics.specialOuterRadius, metrics.specialInnerRadius),
      getSpecialPartColor(data.body[SPECIAL_PART_NAME]),
      3,
      "evenodd",
      signature,
      specialVisible,
    ),
  );

  items.push(
    buildRingItem(
      token,
      metrics,
      "shield-ring",
      buildAnnulusCommands(
        metrics.shieldOuterRadius,
        metrics.shieldInnerRadius,
        0,
        metrics.shieldOffsetY,
      ),
      getPartColor(data.body[SHIELD_PART_NAME]),
      4,
      "evenodd",
      signature,
      shieldVisible,
    ),
  );

  return items;
}

function getExpectedOverlayKinds(data) {
  if (OVERLAY_RENDER_MODE === "image") {
    return [OVERLAY_IMAGE_KIND];
  }
  return FIXED_OVERLAY_KINDS;
}

export async function removeOverlaysForToken(tokenId, items) {
  const sceneItems = items ?? (await OBR.scene.items.getItems());
  const overlayIds = sceneItems
    .filter((item) => item.metadata?.[OVERLAY_KEY] === tokenId)
    .map((item) => item.id);

  if (overlayIds.length) {
    await OBR.scene.items.deleteItems(overlayIds);
  }

  const cachedUrl = overlayRuntimeUrlByTokenId.get(tokenId);
  if (cachedUrl) {
    overlayRuntimeUrlByTokenId.delete(tokenId);
    await deleteCachedOverlaySvg(cachedUrl);
  }
}

async function ensureOverlayForTokenInternal(tokenId, items) {
  const sceneItems = items ?? (await OBR.scene.items.getItems());
  const token = sceneItems.find((item) => item.id === tokenId);
  if (!token || !isCharacterToken(token)) return;
  const overlayItems = sceneItems.filter((item) => item.metadata?.[OVERLAY_KEY] === tokenId);

  if (!isTrackedCharacter(token) || token.visible === false) {
    if (overlayItems.length) {
      await removeOverlaysForToken(tokenId, sceneItems);
    }
    return;
  }

  const data = getTrackerData(token);
  const metrics = await getTokenMetrics(token);
  const overlaySignature = buildOverlaySignature(token, data, metrics);
  const expectedKinds = getExpectedOverlayKinds(data);

  if (hasPatchableOverlaySet(token, overlayItems, expectedKinds)) {
    const signaturesMatch = overlayItems.every(
      (item) => String(item.metadata?.signature ?? "") === overlaySignature,
    );
    if (signaturesMatch) {
      return;
    }

    try {
      if (OVERLAY_RENDER_MODE === "image") {
        const nextOverlayItem = await buildOverlayImageItem(token, data, metrics);
        await OBR.scene.items.updateItems(
          overlayItems.map((item) => item.id),
          (itemsToUpdate) => {
            for (const overlayItem of itemsToUpdate) {
              applyOverlayItemState(overlayItem, nextOverlayItem);
            }
          },
        );
      } else {
        const nextOverlayItems = buildOverlayItems(token, data, metrics, overlaySignature);
        const nextOverlayByKind = new Map(
          nextOverlayItems.map((item) => [String(item.metadata?.kind ?? ""), item]),
        );
        await OBR.scene.items.updateItems(
          overlayItems.map((item) => item.id),
          (itemsToUpdate) => {
            for (const overlayItem of itemsToUpdate) {
              const kind = String(overlayItem.metadata?.kind ?? "");
              const nextItem = nextOverlayByKind.get(kind);
              if (!nextItem) continue;
              applyOverlayItemState(overlayItem, nextItem);
            }
          },
        );
      }
      return;
    } catch (error) {
      console.warn("[Body HP] Overlay patch failed, falling back to rebuild", error);
    }
  }

  await removeOverlaysForToken(tokenId);

  if (OVERLAY_RENDER_MODE === "image") {
    const overlayItem = await buildOverlayImageItem(token, data, metrics);
    await OBR.scene.items.addItems([overlayItem]);
  } else {
    await OBR.scene.items.addItems(
      buildOverlayItems(token, data, metrics, overlaySignature),
    );
  }
}

export async function ensureOverlayForToken(tokenId, items) {
  const previous = overlayEnsureQueue.get(tokenId) ?? Promise.resolve();
  const next = previous
    .catch(() => {})
    .then(() => ensureOverlayForTokenInternal(tokenId, items));

  overlayEnsureQueue.set(tokenId, next);

  try {
    await next;
  } finally {
    if (overlayEnsureQueue.get(tokenId) === next) {
      overlayEnsureQueue.delete(tokenId);
    }
  }
}

export async function setTrackedState(tokenId, enabled) {
  if (enabled) {
    await updateTrackerData(tokenId, (current) => ({
      ...current,
      enabled: true,
    }));
    await ensureOverlayForToken(tokenId);
    return;
  }

  await OBR.scene.items.updateItems([tokenId], (items) => {
    const token = items[0];
    if (!token) return;
    token.metadata ??= {};
    delete token.metadata[META_KEY];
  });

  await removeOverlaysForToken(tokenId);
}

export async function applyRemoteRollEvent(event) {
  if (!event?.token_id) return false;

  const sceneItems = await OBR.scene.items.getItems();
  const token = sceneItems.find((item) => item.id === event.token_id);
  if (!token || !isTrackedCharacter(token)) return false;

  await updateTrackerData(token.id, (current) => applyRollEventToData(current, event));
  await ensureOverlayForToken(token.id);
  return true;
}

export async function syncTrackedOverlays() {
  const items = await OBR.scene.items.getItems();
  const byId = new Map(items.map((item) => [item.id, item]));
  const overlaysByTokenId = new Map();

  for (const item of items.filter(isOverlayItem)) {
    const tokenId = String(item.metadata?.[OVERLAY_KEY] ?? "");
    if (!tokenId) continue;
    const bucket = overlaysByTokenId.get(tokenId) ?? [];
    bucket.push(item);
    overlaysByTokenId.set(tokenId, bucket);
  }

  const staleOverlayIds = items
    .filter(isOverlayItem)
    .filter((item) => {
      const token = byId.get(item.metadata[OVERLAY_KEY]);
      return (
        !token ||
        !isTrackedCharacter(token) ||
        token.visible === false ||
        Number(item.metadata?.visualVersion ?? 0) !== VISUAL_VERSION
      );
    })
    .map((item) => item.id);

  if (staleOverlayIds.length) {
    await OBR.scene.items.deleteItems(staleOverlayIds);
  }

  const trackedTokens = items.filter((item) => isTrackedCharacter(item) && item.visible !== false);
  for (const token of trackedTokens) {
    const overlayItems = overlaysByTokenId.get(token.id) ?? [];
    const data = getTrackerData(token);
    const expectedKinds = getExpectedOverlayKinds(data);
    const patchable = hasPatchableOverlaySet(token, overlayItems, expectedKinds);
    let needsRebuild = !patchable;

    if (!needsRebuild && overlayItems.length) {
      const metrics = await getTokenMetrics(token);
      const expectedSignature = buildOverlaySignature(token, data, metrics);
      needsRebuild = overlayItems.some(
        (item) => String(item.metadata?.signature ?? "") !== expectedSignature,
      );
    }

    if (needsRebuild) {
      await ensureOverlayForToken(token.id);
    }
  }
}
