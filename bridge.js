import { OBR, META_KEY, getCharacterName, getTrackerData, isTrackedCharacter } from "./shared.js";

export const BRIDGE_ROOM_KEY = "com.codex.body-hp/bridge";

const DEFAULT_BRIDGE_CONFIG = {
  enabled: true,
  baseUrl: "http://127.0.0.1:8765",
  pollIntervalMs: 2500,
};

function clampInterval(value) {
  const numeric = Number(value) || DEFAULT_BRIDGE_CONFIG.pollIntervalMs;
  return Math.max(1000, Math.min(numeric, 15000));
}

export function normalizeBridgeConfig(raw) {
  return {
    enabled: raw?.enabled !== false,
    baseUrl: String(raw?.baseUrl || DEFAULT_BRIDGE_CONFIG.baseUrl).replace(/\/+$/, ""),
    pollIntervalMs: clampInterval(raw?.pollIntervalMs),
  };
}

export async function getBridgeConfig() {
  try {
    const metadata = await OBR.room.getMetadata();
    return normalizeBridgeConfig(metadata?.[BRIDGE_ROOM_KEY]);
  } catch (error) {
    console.warn("[Body HP] Unable to read bridge config, using defaults", error);
    return { ...DEFAULT_BRIDGE_CONFIG };
  }
}

async function requestJson(config, path, options = {}) {
  const response = await fetch(`${config.baseUrl}${path}`, {
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
    ...options,
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `Bridge request failed with ${response.status}`);
  }

  return response.json();
}

export function serializeTrackedToken(token) {
  const data = getTrackerData(token);
  return {
    token_id: token.id,
    token_name: getCharacterName(token),
    player_id: data.identity?.playerId || "",
    character_id: data.identity?.characterId || "",
    tracker_data: {
      ...data,
      enabled: isTrackedCharacter(token),
    },
  };
}

export async function pushTokenSnapshots(tokens) {
  if (!tokens.length) return;

  const config = await getBridgeConfig();
  if (!config.enabled) return;

  const roomId = OBR.room.id;
  await requestJson(config, "/api/tokens/upsert", {
    method: "POST",
    body: JSON.stringify({
      room_id: roomId,
      tokens: tokens.map(serializeTrackedToken),
    }),
  });
}

export async function fetchRollEvents(sinceId) {
  const config = await getBridgeConfig();
  if (!config.enabled) {
    return { items: [], pollIntervalMs: config.pollIntervalMs };
  }

  const roomId = encodeURIComponent(OBR.room.id);
  const response = await requestJson(
    config,
    `/api/events?room_id=${roomId}&since_id=${encodeURIComponent(String(sinceId || 0))}`,
    { method: "GET" },
  );

  return {
    items: Array.isArray(response.items) ? response.items : [],
    pollIntervalMs: config.pollIntervalMs,
  };
}

export async function acknowledgeRollEvents(eventIds) {
  if (!eventIds.length) return;

  const config = await getBridgeConfig();
  if (!config.enabled) return;

  await requestJson(config, "/api/events/ack", {
    method: "POST",
    body: JSON.stringify({
      room_id: OBR.room.id,
      event_ids: eventIds,
    }),
  });
}

export function extractTrackedTokens(items) {
  return items.filter((item) => item.metadata?.[META_KEY]?.enabled === true);
}
