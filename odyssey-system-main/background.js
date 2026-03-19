import {
  OBR,
  META_KEY,
  applyRemoteRollEvent,
  getCharacterName,
  isCharacterToken,
  isTrackedCharacter,
  setTrackedState,
  syncTrackedOverlays,
} from "./shared.js";
import {
  acknowledgeRollEvents,
  extractTrackedTokens,
  fetchRollEvents,
  getBridgeConfig,
  pushTokenSnapshots,
} from "./bridge.js";

const EXTENSION_MENU_ID = "com.codex.body-hp/context-menu";
let currentRole = "PLAYER";
let lastBridgeEventId = 0;
let bridgePollTimer = null;
let pushStateTimer = null;
const ADD_ICON_URL = new URL("./add.svg", import.meta.url).href;
const REMOVE_ICON_URL = new URL("./remove.svg", import.meta.url).href;

async function updateBadge() {
  try {
    const items = await OBR.scene.items.getItems();
    const trackedCount = items.filter(isTrackedCharacter).length;
    await OBR.action.setBadgeText(trackedCount ? String(trackedCount) : undefined);
  } catch (error) {
    console.warn("[Body HP] Unable to update badge", error);
  }
}

async function toggleTracking(items) {
  const characters = items.filter(isCharacterToken);
  if (!characters.length) return;

  const tracked = characters.filter(isTrackedCharacter);
  const untracked = characters.filter((item) => !isTrackedCharacter(item));
  const shouldEnable = untracked.length > 0;
  const targets = shouldEnable ? untracked : tracked;

  for (const character of targets) {
    await setTrackedState(character.id, shouldEnable);
  }

  console.log(
    `[Body HP] ${shouldEnable ? "Tracking" : "Untracking"}: ${targets
      .map(getCharacterName)
      .join(", ")}`
  );

  if (currentRole === "GM") {
    const sceneItems = await OBR.scene.items.getItems();
    await pushTokenSnapshots(extractTrackedTokens(sceneItems));
  }

  await updateBadge();
}

function scheduleTokenSync(delayMs = 800) {
  if (pushStateTimer) {
    clearTimeout(pushStateTimer);
  }

  pushStateTimer = setTimeout(() => {
    pushStateTimer = null;
    if (currentRole !== "GM") return;

    void OBR.scene.items
      .getItems()
      .then((items) => pushTokenSnapshots(extractTrackedTokens(items)))
      .catch((error) => {
        console.warn("[Body HP] Unable to push token snapshots", error);
      });
  }, delayMs);
}

async function pollBridgeOnce() {
  if (currentRole !== "GM") return;

  const { items } = await fetchRollEvents(lastBridgeEventId);
  if (!items.length) return;

  const appliedEventIds = [];
  for (const event of items) {
    const applied = await applyRemoteRollEvent(event);
    lastBridgeEventId = Math.max(lastBridgeEventId, Number(event.id) || 0);
    if (applied) {
      appliedEventIds.push(event.id);
    }
  }

  if (appliedEventIds.length) {
    await acknowledgeRollEvents(appliedEventIds);
    scheduleTokenSync(200);
  }
}

async function restartBridgePolling() {
  if (bridgePollTimer) {
    clearTimeout(bridgePollTimer);
    bridgePollTimer = null;
  }

  try {
    await pollBridgeOnce();
    const { pollIntervalMs } = await getBridgeConfig();
    bridgePollTimer = setTimeout(() => {
      void restartBridgePolling();
    }, pollIntervalMs);
  } catch (error) {
    console.warn("[Body HP] Bridge polling failed", error);
    bridgePollTimer = setTimeout(() => {
      void restartBridgePolling();
    }, 5000);
  }
}

async function setupContextMenu() {
  await OBR.contextMenu.create({
    id: EXTENSION_MENU_ID,
    icons: [
      {
        icon: ADD_ICON_URL,
        label: "Track Body HP",
        filter: {
          roles: ["GM"],
          every: [{ key: "layer", value: "CHARACTER" }],
          some: [{ key: ["metadata", META_KEY, "enabled"], value: true, operator: "!=" }],
        },
      },
      {
        icon: REMOVE_ICON_URL,
        label: "Remove Body HP",
        filter: {
          roles: ["GM"],
          every: [
            { key: "layer", value: "CHARACTER" },
            { key: ["metadata", META_KEY, "enabled"], value: true },
          ],
        },
      },
    ],
    onClick(context) {
      return toggleTracking(context.items).catch((error) => {
        console.error("[Body HP] Context menu failed", error);
      });
    },
  });
}

OBR.onReady(async () => {
  try {
    currentRole = await OBR.player.getRole();

    await setupContextMenu();
    await updateBadge();

    if (currentRole === "GM") {
      await syncTrackedOverlays();
      scheduleTokenSync(100);
      void restartBridgePolling();
    }

    OBR.scene.items.onChange(() => {
      void updateBadge();
      scheduleTokenSync();
    });

    OBR.player.onChange(async () => {
      const nextRole = await OBR.player.getRole();
      if (nextRole !== currentRole && nextRole === "GM") {
        await syncTrackedOverlays();
        scheduleTokenSync(100);
        void restartBridgePolling();
      }
      currentRole = nextRole;
    });

    console.log("[Body HP] Background ready");
  } catch (error) {
    console.error("[Body HP] Background init failed", error);
  }
});
