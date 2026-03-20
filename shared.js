import OBR, { Command, buildPath, isImage } from "@owlbear-rodeo/sdk";

export { OBR };

export const EXTENSION_ID = "com.codex.body-hp";
export const META_KEY = `${EXTENSION_ID}/data`;
export const OVERLAY_KEY = `${EXTENSION_ID}/overlayFor`;
export const BODY_ORDER = ["Head", "L.Arm", "R.Arm", "Torso", "L.Leg", "R.Leg"];
export const ROLL_HISTORY_LIMIT = 12;
export const COMBAT_SKILL_CATEGORY = "combat";
export const APPLIED_SKILL_CATEGORY = "applied";
export const MELEE_SKILL_NAME = "Melee";
export const PARRY_SKILL_NAME = "Parry";
const LEGACY_MELEE_SKILL_NAMES = new Set(["Hand", "Cold", "\u0420\u0443\u043A\u043E\u043F\u0430\u0448\u043D\u044B\u0439"]);
const LEGACY_REMOVED_SKILLS = new Set(["Hand", "Cold", "Throwing", "Rifle", "Turrets"]);
const VISUAL_VERSION = 3;
const RING_COLORS = {
  full: "#73FF5A",
  half: "#FFAF22",
  kaputt: "#FF460D",
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
    next.attributes[key] = clamp(Number(fallbackValue) || 0, 0, 15);
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
  return [
    `Head ${body["Head"].current}/${body["Head"].max}(${body["Head"].armor}) | L.Arm ${body["L.Arm"].current}/${body["L.Arm"].max}(${body["L.Arm"].armor}) | R.Arm ${body["R.Arm"].current}/${body["R.Arm"].max}(${body["R.Arm"].armor})`,
    `Torso ${body["Torso"].current}/${body["Torso"].max}(${body["Torso"].armor}) | L.Leg ${body["L.Leg"].current}/${body["L.Leg"].max}(${body["L.Leg"].armor}) | R.Leg ${body["R.Leg"].current}/${body["R.Leg"].max}(${body["R.Leg"].armor})`,
  ].join("\n");
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
  return BODY_ORDER.reduce(
    (accumulator, partName) => {
      accumulator.current += data.body[partName].current;
      accumulator.max += data.body[partName].max;
      return accumulator;
    },
    { current: 0, max: 0 },
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
  let center = token.position;
  let width = effectiveSize.width;
  let height = effectiveSize.height;

  try {
    const bounds = await OBR.scene.items.getItemBounds([token.id]);
    if (bounds?.width > 0 && bounds?.height > 0) {
      center = bounds.center;
      width = bounds.width;
      height = bounds.height;
    }
  } catch (error) {
    console.warn("[Body HP] Unable to read token bounds, using fallback size", error);
  }

  let gridDpi = 150;
  try {
    gridDpi = (await OBR.scene.grid.getDpi()) || gridDpi;
  } catch (error) {
    console.warn("[Body HP] Unable to read grid dpi, using fallback size", error);
  }

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

  return {
    center,
    visibleDiameter,
    outerRadius,
    outerInnerRadius,
    torsoOuterRadius,
    torsoInnerRadius,
  };
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

function buildAnnulusCommands(radiusOuter, radiusInner) {
  const outer = arcPoints(radiusOuter, -180, 180, 36);
  const inner = arcPoints(radiusInner, -180, 180, 36);
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

function getPartColor(part) {
  if (part.max <= 0 || part.current <= 0) return RING_COLORS.kaputt;
  if (part.current < part.max) return RING_COLORS.half;
  return RING_COLORS.full;
}

function buildRingItem(token, metrics, kind, commands, fillColor, zIndex = 0, fillRule = "nonzero") {
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
    .zIndex(Date.now() + zIndex)
    .attachedTo(token.id)
    .disableAttachmentBehavior(["ROTATION"])
    .layer("ATTACHMENT")
    .locked(true)
    .disableHit(true)
    .metadata({
      [OVERLAY_KEY]: token.id,
      kind,
      visualVersion: VISUAL_VERSION,
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

export function buildOverlayItems(token, data, metrics) {
  const items = [];

  items.push(
    buildRingItem(
      token,
      metrics,
      "outer-base",
      buildAnnulusCommands(metrics.outerRadius, metrics.outerInnerRadius),
      RING_COLORS.base,
      0,
      "evenodd",
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
    ),
  );

  return items;
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

export async function removeOverlaysForToken(tokenId, items) {
  const sceneItems = items ?? (await OBR.scene.items.getItems());
  const overlayIds = sceneItems
    .filter((item) => item.metadata?.[OVERLAY_KEY] === tokenId)
    .map((item) => item.id);

  if (overlayIds.length) {
    await OBR.scene.items.deleteItems(overlayIds);
  }
}

export async function ensureOverlayForToken(tokenId, items) {
  const sceneItems = items ?? (await OBR.scene.items.getItems());
  const token = sceneItems.find((item) => item.id === tokenId);
  if (!token || !isCharacterToken(token)) return;

  await removeOverlaysForToken(tokenId, sceneItems);

  if (!isTrackedCharacter(token)) return;

  const metrics = await getTokenMetrics(token);
  await OBR.scene.items.addItems(
    buildOverlayItems(token, getTrackerData(token), metrics),
  );
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

  const staleOverlayIds = items
    .filter(isOverlayItem)
    .filter((item) => {
      const token = byId.get(item.metadata[OVERLAY_KEY]);
      return !token || !isTrackedCharacter(token);
    })
    .map((item) => item.id);

  if (staleOverlayIds.length) {
    await OBR.scene.items.deleteItems(staleOverlayIds);
  }

  const trackedTokens = items.filter(isTrackedCharacter);
  for (const token of trackedTokens) {
    await ensureOverlayForToken(token.id, items);
  }
}
