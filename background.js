import {
  OBR,
  applyRemoteRollEvent,
  isTrackedCharacter,
  syncTrackedOverlays,
} from "./shared.js";
import {
  acknowledgeRollEvents,
  extractTrackedTokens,
  fetchRollEvents,
  getBridgeConfig,
  pushTokenSnapshots,
} from "./bridge.js";

let currentRole = "PLAYER";
let lastBridgeEventId = 0;
let bridgePollTimer = null;
let pushStateTimer = null;

async function updateBadge() {
  try {
    const items = await OBR.scene.items.getItems();
    const trackedCount = items.filter(isTrackedCharacter).length;
    await OBR.action.setBadgeText(trackedCount ? String(trackedCount) : undefined);
  } catch (error) {
    console.warn("[Body HP] Unable to update badge", error);
  }
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

OBR.onReady(async () => {
  try {
    currentRole = await OBR.player.getRole();
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
