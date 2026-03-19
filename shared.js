import OBR, { buildLabel, buildShape, isImage } from "@owlbear-rodeo/sdk";

export { OBR };

export const EXTENSION_ID = "com.codex.body-hp";
export const META_KEY = `${EXTENSION_ID}/data`;
export const OVERLAY_KEY = `${EXTENSION_ID}/overlayFor`;
export const BODY_ORDER = ["L.Arm", "Head", "R.Arm", "L.Leg", "Torso", "R.Leg"];
export const ROLL_HISTORY_LIMIT = 12;

export const BODY_DEFAULTS = {
  "L.Arm": { current: 2, max: 2, armor: 2 },
  Head: { current: 1, max: 1, armor: 0 },
  "R.Arm": { current: 2, max: 2, armor: 2 },
  "L.Leg": { current: 2, max: 2, armor: 2 },
  Torso: { current: 3, max: 3, armor: 6 },
  "R.Leg": { current: 2, max: 2, armor: 2 },
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
    skills: {
      Hand: 0,
      Cold: 0,
      Throwing: 0,
      Rifle: 0,
      Turrets: 0,
    },
    attributes: {
      Strength: 0,
      Agility: 0,
      Reaction: 0,
      Endurance: 0,
      Perception: 0,
      Intelligence: 0,
      Charisma: 0,
      Willpower: 0,
      Psionics: 0,
      Parry: 0,
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
    part.max = clamp(Number(source.max ?? part.max) || part.max, 0, 99);
    part.current = clamp(
      Number(source.current ?? part.current) || part.current,
      0,
      part.max,
    );
    part.armor = clamp(Number(source.armor ?? part.armor) || part.armor, 0, 99);
  }

  return next;
}

export function sanitizeOdysseyData(raw) {
  const next = deepClone(DEFAULT_TRACKER_DATA.odyssey);
  if (!raw || typeof raw !== "object") return next;

  next.owner.playerId = String(raw.owner?.playerId ?? "").trim();
  next.owner.playerName = String(raw.owner?.playerName ?? "").trim();

  for (const key of Object.keys(next.skills)) {
    next.skills[key] = clamp(Number(raw.skills?.[key] ?? 0) || 0, 0, 10);
  }

  for (const key of Object.keys(next.attributes)) {
    next.attributes[key] = clamp(Number(raw.attributes?.[key] ?? 0) || 0, 0, 10);
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
    `L.Arm ${body["L.Arm"].current}/${body["L.Arm"].max}(${body["L.Arm"].armor}) | Head ${body["Head"].current}/${body["Head"].max}(${body["Head"].armor}) | R.Arm ${body["R.Arm"].current}/${body["R.Arm"].max}(${body["R.Arm"].armor})`,
    `L.Leg ${body["L.Leg"].current}/${body["L.Leg"].max}(${body["L.Leg"].armor}) | Torso ${body["Torso"].current}/${body["Torso"].max}(${body["Torso"].armor}) | R.Leg ${body["R.Leg"].current}/${body["R.Leg"].max}(${body["R.Leg"].armor})`,
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

  let gridDpi = 0;
  try {
    gridDpi = await OBR.scene.grid.getDpi();
  } catch (error) {
    console.warn("[Body HP] Unable to read grid dpi, using size fallback", error);
  }

  const scaleFactor = Math.max(
    Math.abs(token.scale?.x ?? 1),
    Math.abs(token.scale?.y ?? 1),
    1,
  );
  const markerSize = Math.max(
    width,
    height,
    effectiveSize.width,
    effectiveSize.height,
    gridDpi * scaleFactor,
    56,
  );

  return {
    center,
    width,
    height,
    markerSize,
  };
}

function getWorldPosition(token, center, offsetX, offsetY) {
  const radians = ((token.rotation ?? 0) * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  return {
    x: center.x + offsetX * cos - offsetY * sin,
    y: center.y + offsetX * sin + offsetY * cos,
  };
}

function buildOverlayCard(token, data, metrics) {
  const width = Math.max(360, Math.round(metrics.width * 2.55));
  const height = data.lastRoll ? 96 : 72;
  const offsetX = metrics.width / 2 + width / 2 + 18;

  return buildLabel()
    .name(`Body HP: ${getCharacterName(token)}`)
    .plainText(formatOverlayText(data))
    .width(width)
    .height(height)
    .padding(10)
    .fontSize(13)
    .fontWeight(600)
    .lineHeight(1.18)
    .textAlign("LEFT")
    .textAlignVertical("MIDDLE")
    .fillColor("#f8fafc")
    .backgroundColor("#020617")
    .backgroundOpacity(0.58)
    .strokeColor("#cbd5e1")
    .strokeOpacity(0.45)
    .strokeWidth(1)
    .cornerRadius(12)
    .pointerDirection("LEFT")
    .pointerWidth(10)
    .pointerHeight(12)
    .position(getWorldPosition(token, metrics.center, offsetX, 0))
    .attachedTo(token.id)
    .layer("ATTACHMENT")
    .locked(true)
    .disableHit(true)
    .metadata({ [OVERLAY_KEY]: token.id, kind: "body-card" })
    .build();
}

function buildMinorDots(token, data, metrics) {
  const items = [];
  const startX = -metrics.markerSize / 2 + 12;
  const y = metrics.markerSize / 2 - 12;

  for (let index = 0; index < data.minor; index += 1) {
    items.push(
      buildShape()
        .shapeType("CIRCLE")
        .width(8)
        .height(8)
        .position(getWorldPosition(token, metrics.center, startX + index * 10, y))
        .attachedTo(token.id)
        .layer("ATTACHMENT")
        .locked(true)
        .disableHit(true)
        .fillColor("#f59e0b")
        .fillOpacity(0.98)
        .strokeColor("#111827")
        .strokeWidth(1)
        .metadata({ [OVERLAY_KEY]: token.id, kind: "minor", index })
        .build(),
    );
  }

  return items;
}

function buildSeriousBars(token, data, metrics) {
  const items = [];
  const x = metrics.markerSize / 2 - 12;
  const startY = -metrics.markerSize / 2 + 13;

  for (let index = 0; index < data.serious; index += 1) {
    items.push(
      buildShape()
        .shapeType("RECTANGLE")
        .width(4)
        .height(18)
        .position(getWorldPosition(token, metrics.center, x - index * 8, startY))
        .attachedTo(token.id)
        .layer("ATTACHMENT")
        .locked(true)
        .disableHit(true)
        .fillColor("#ef4444")
        .fillOpacity(0.98)
        .strokeColor("#111827")
        .strokeWidth(1)
        .metadata({ [OVERLAY_KEY]: token.id, kind: "serious", index })
        .build(),
    );
  }

  return items;
}

export function buildBodyFigure(token, data, metrics) {
  const parts = [];
  const spacing = 14;

  const positions = {
    Head: [0, -spacing * 2],
    Torso: [0, 0],
    "L.Arm": [-spacing, 0],
    "R.Arm": [spacing, 0],
    "L.Leg": [-spacing / 2, spacing * 2],
    "R.Leg": [spacing / 2, spacing * 2],
  };

  for (const partName of BODY_ORDER) {
    const part = data.body[partName];
    const [x, y] = positions[partName];
    const hpRatio = part.max > 0 ? part.current / part.max : 0;

    let color = "#22c55e";
    if (hpRatio < 0.5) color = "#f59e0b";
    if (hpRatio < 0.25) color = "#ef4444";

    parts.push(
      buildShape()
        .shapeType("RECTANGLE")
        .width(10)
        .height(10)
        .position(getWorldPosition(token, metrics.center, x, y))
        .attachedTo(token.id)
        .layer("ATTACHMENT")
        .locked(true)
        .disableHit(true)
        .fillColor(color)
        .fillOpacity(0.95)
        .strokeColor("#111")
        .strokeWidth(1)
        .metadata({
          [OVERLAY_KEY]: token.id,
          kind: "body-part",
          part: partName,
        })
        .build(),
    );
  }

  return parts;
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
  return [
    buildOverlayCard(token, data, metrics),
    ...buildBodyFigure(token, data, metrics),
  ];
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
