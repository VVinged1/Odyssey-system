import OBR from "@owlbear-rodeo/sdk";

// Shared keys with the main Odyssey System extension.
const DEBUG_LOG_KEY = "com.codex.body-hp/debugLog";
const DEBUG_BROADCAST_CHANNEL = "com.codex.body-hp/debug";
const ENTRY_LIMIT = 50;
const POLL_INTERVAL_MS = 2000;
const SIZE_STORAGE_KEY = "odyssey-combat-log/window-size";
const VIEW_CUTOFF_STORAGE_KEY = "odyssey-combat-log/view-cutoff-id";
const DEFAULT_WINDOW_SIZE = { width: 520, height: 780 };
const COMPACT_WINDOW_SIZE = { width: 440, height: 640 };
const LARGE_WINDOW_SIZE = { width: 660, height: 960 };
const WIDTH_STEP = 40;
const HEIGHT_STEP = 60;
const MIN_WINDOW_WIDTH = 420;
const MAX_WINDOW_WIDTH = 960;
const MIN_WINDOW_HEIGHT = 560;
const MAX_WINDOW_HEIGHT = 1200;

const ui = {
  refreshBtn: document.getElementById("refreshBtn"),
  clearViewBtn: document.getElementById("clearViewBtn"),
  restoreViewBtn: document.getElementById("restoreViewBtn"),
  clearBtn: document.getElementById("clearBtn"),
  sizeLabel: document.getElementById("sizeLabel"),
  sizeCompactBtn: document.getElementById("sizeCompactBtn"),
  sizeDefaultBtn: document.getElementById("sizeDefaultBtn"),
  sizeLargeBtn: document.getElementById("sizeLargeBtn"),
  widthDownBtn: document.getElementById("widthDownBtn"),
  widthUpBtn: document.getElementById("widthUpBtn"),
  heightDownBtn: document.getElementById("heightDownBtn"),
  heightUpBtn: document.getElementById("heightUpBtn"),
  liveBadge: document.getElementById("liveBadge"),
  entryCount: document.getElementById("entryCount"),
  statusBox: document.getElementById("statusBox"),
  viewerName: document.getElementById("viewerName"),
  viewerRole: document.getElementById("viewerRole"),
  lastSync: document.getElementById("lastSync"),
  emptyState: document.getElementById("emptyState"),
  emptyTitle: document.getElementById("emptyTitle"),
  emptyBody: document.getElementById("emptyBody"),
  logEntries: document.getElementById("logEntries"),
  resizeHandle: document.getElementById("resizeHandle"),
};

let sharedEntries = [];
let viewerName = "Unknown";
let viewerRole = "PLAYER";
let lastSyncLabel = "Not synced yet";
let roomRefreshTimer = null;
let windowSize = { ...DEFAULT_WINDOW_SIZE };
let localViewCutoffId = 0;

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function sanitizeDebugEntries(raw) {
  if (!Array.isArray(raw)) return [];
  return raw
    .filter((entry) => entry && typeof entry === "object")
    .map((entry) => ({
      id: Number(entry.id) || Date.now(),
      title: String(entry.title ?? "Debug"),
      body: String(entry.body ?? ""),
      kind: String(entry.kind ?? "info"),
      timestamp: String(entry.timestamp ?? ""),
    }))
    .slice(0, ENTRY_LIMIT);
}

function mergeDebugEntries(...entryGroups) {
  const merged = new Map();

  for (const group of entryGroups) {
    for (const entry of sanitizeDebugEntries(group)) {
      merged.set(entry.id, entry);
    }
  }

  return [...merged.values()]
    .sort((left, right) => Number(right.id) - Number(left.id))
    .slice(0, ENTRY_LIMIT);
}

function kindClass(kind) {
  switch (kind) {
    case "success":
      return "kind-success";
    case "error":
      return "kind-error";
    case "warning":
      return "kind-warning";
    default:
      return "kind-info";
  }
}

function formatKind(kind) {
  switch (kind) {
    case "success":
      return "Success";
    case "error":
      return "Error";
    case "warning":
      return "Warning";
    default:
      return "Info";
  }
}

function setStatus(message) {
  ui.statusBox.textContent = message;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function normalizeWindowSize(width, height) {
  return {
    width: clamp(Number(width) || DEFAULT_WINDOW_SIZE.width, MIN_WINDOW_WIDTH, MAX_WINDOW_WIDTH),
    height: clamp(Number(height) || DEFAULT_WINDOW_SIZE.height, MIN_WINDOW_HEIGHT, MAX_WINDOW_HEIGHT),
  };
}

function formatWindowSize(size = windowSize) {
  return `${size.width} x ${size.height}`;
}

function renderWindowSize() {
  if (ui.sizeLabel) {
    ui.sizeLabel.textContent = formatWindowSize();
  }
}

function saveWindowSize(size = windowSize) {
  try {
    window.localStorage.setItem(SIZE_STORAGE_KEY, JSON.stringify(size));
  } catch (error) {
    console.warn("[Odyssey Combat Log] Unable to store window size", error);
  }
}

function loadStoredWindowSize() {
  try {
    const raw = window.localStorage.getItem(SIZE_STORAGE_KEY);
    if (!raw) return { ...DEFAULT_WINDOW_SIZE };
    const parsed = JSON.parse(raw);
    return normalizeWindowSize(parsed?.width, parsed?.height);
  } catch (error) {
    console.warn("[Odyssey Combat Log] Unable to read stored window size", error);
    return { ...DEFAULT_WINDOW_SIZE };
  }
}

function saveViewCutoff() {
  try {
    window.localStorage.setItem(VIEW_CUTOFF_STORAGE_KEY, String(localViewCutoffId));
  } catch (error) {
    console.warn("[Odyssey Combat Log] Unable to store local view cutoff", error);
  }
}

function loadStoredViewCutoff() {
  try {
    const raw = window.localStorage.getItem(VIEW_CUTOFF_STORAGE_KEY);
    return Math.max(0, Number(raw) || 0);
  } catch (error) {
    console.warn("[Odyssey Combat Log] Unable to read local view cutoff", error);
    return 0;
  }
}

function getVisibleEntries() {
  return sharedEntries.filter((entry) => Number(entry.id) > localViewCutoffId).slice(0, ENTRY_LIMIT);
}

function hasHiddenEntries() {
  return sharedEntries.some((entry) => Number(entry.id) <= localViewCutoffId);
}

async function applyWindowSize(nextSize, label = "Window resized") {
  const normalized = normalizeWindowSize(nextSize?.width, nextSize?.height);
  await Promise.all([
    OBR.action.setWidth(normalized.width),
    OBR.action.setHeight(normalized.height),
  ]);
  windowSize = normalized;
  saveWindowSize(windowSize);
  renderWindowSize();
  if (label) setStatus(`${label}: ${formatWindowSize(windowSize)}.`);
}

function bindPointerResize() {
  const handle = ui.resizeHandle;
  if (!handle) return;

  let startX = 0;
  let startY = 0;
  let startSize = { ...windowSize };
  let pendingSize = null;
  let frame = 0;

  const commit = () => {
    frame = 0;
    if (!pendingSize) return;
    void applyWindowSize(pendingSize, "").catch((error) => {
      console.warn("[Odyssey Combat Log] Pointer resize failed", error);
    });
    pendingSize = null;
  };

  handle.addEventListener("pointerdown", (event) => {
    startX = event.clientX;
    startY = event.clientY;
    startSize = { ...windowSize };
    handle.setPointerCapture(event.pointerId);
    event.preventDefault();
  });

  handle.addEventListener("pointermove", (event) => {
    if (!handle.hasPointerCapture(event.pointerId)) return;
    pendingSize = normalizeWindowSize(
      startSize.width + event.clientX - startX,
      startSize.height + event.clientY - startY,
    );
    if (!frame) frame = window.requestAnimationFrame(commit);
  });

  const finish = (event) => {
    if (handle.hasPointerCapture(event.pointerId)) handle.releasePointerCapture(event.pointerId);
    if (frame) window.cancelAnimationFrame(frame);
    frame = 0;
    if (pendingSize) {
      const finalSize = pendingSize;
      pendingSize = null;
      void applyWindowSize(finalSize, "Window resized").catch((error) => {
        setStatus(error?.message ?? "Unable to resize window.");
      });
    }
  };

  handle.addEventListener("pointerup", finish);
  handle.addEventListener("pointercancel", finish);
}

function setSyncState(label) {
  lastSyncLabel = `${label} at ${new Date().toLocaleTimeString()}`;
  ui.lastSync.textContent = lastSyncLabel;
  ui.liveBadge.textContent = label;
}

function renderHeader() {
  ui.viewerName.textContent = viewerName;
  ui.viewerRole.textContent = viewerRole;
  const visibleEntries = getVisibleEntries();
  ui.entryCount.textContent = `${visibleEntries.length} ${visibleEntries.length === 1 ? "entry" : "entries"}`;
  ui.lastSync.textContent = lastSyncLabel;
}

function renderControlState() {
  if (ui.clearViewBtn) {
    ui.clearViewBtn.disabled = !getVisibleEntries().length;
  }
  if (ui.restoreViewBtn) {
    ui.restoreViewBtn.disabled = !hasHiddenEntries();
  }
}

function renderEntries() {
  renderWindowSize();
  renderHeader();
  renderControlState();
  const visibleEntries = getVisibleEntries();

  if (!visibleEntries.length) {
    ui.emptyState.hidden = false;
    ui.logEntries.innerHTML = "";
    if (sharedEntries.length && hasHiddenEntries()) {
      ui.emptyTitle.textContent = "Log view cleared";
      ui.emptyBody.textContent =
        "Current entries are hidden only on this client. New events will appear automatically, or use Restore View.";
    } else {
      ui.emptyTitle.textContent = "No combat entries yet";
      ui.emptyBody.textContent =
        "As soon as the main Odyssey extension rolls attacks or checks, the shared room log will appear here.";
    }
    return;
  }

  ui.emptyState.hidden = true;
  ui.logEntries.innerHTML = visibleEntries
    .map(
      (entry) => `
        <article class="entry-card">
          <div class="entry-head">
            <div class="entry-title">${escapeHtml(entry.title)}</div>
            <div class="entry-time">${escapeHtml(entry.timestamp || "Unknown time")}</div>
          </div>
          <div class="kind-pill ${kindClass(entry.kind)}">${escapeHtml(formatKind(entry.kind))}</div>
          <pre class="entry-body">${escapeHtml(entry.body)}</pre>
        </article>`,
    )
    .join("");
}

function haveEntriesChanged(nextEntries) {
  if (sharedEntries.length !== nextEntries.length) return true;
  return sharedEntries.some((entry, index) => entry.id !== nextEntries[index]?.id);
}

async function refreshFromRoom(label = "Room refresh", options = {}) {
  const { quiet = false } = options;
  const metadata = await OBR.room.getMetadata();
  const nextEntries = sanitizeDebugEntries(metadata?.[DEBUG_LOG_KEY]);
  const changed = haveEntriesChanged(nextEntries);
  sharedEntries = nextEntries;
  if (!sharedEntries.length && localViewCutoffId) {
    localViewCutoffId = 0;
    saveViewCutoff();
  }

  if (quiet && !changed) return;

  setSyncState(label);
  setStatus("Connected to the shared Odyssey combat log.");
  renderEntries();
}

function clearLocalView() {
  if (!getVisibleEntries().length) {
    setStatus("There are no visible combat entries to clear.");
    return;
  }

  localViewCutoffId = Math.max(
    localViewCutoffId,
    ...sharedEntries.map((entry) => Number(entry.id) || 0),
  );
  saveViewCutoff();
  setSyncState("Local view cleared");
  setStatus("Log output cleared locally. New entries will still appear.");
  renderEntries();
}

function restoreLocalView() {
  if (!hasHiddenEntries()) {
    setStatus("There are no hidden combat entries to restore.");
    return;
  }

  localViewCutoffId = 0;
  saveViewCutoff();
  setSyncState("View restored");
  setStatus("Hidden combat entries restored to this client.");
  renderEntries();
}

async function clearSharedLog() {
  if (viewerRole !== "GM") {
    setStatus("Only the GM can clear the shared combat log.");
    return;
  }

  sharedEntries = [];
  setSyncState("Log cleared");
  setStatus("Shared Odyssey combat log cleared.");
  renderEntries();

  await OBR.broadcast.sendMessage(
    DEBUG_BROADCAST_CHANNEL,
    { type: "debug-clear" },
    { destination: "ALL" },
  );
  await OBR.room.setMetadata({
    [DEBUG_LOG_KEY]: [],
  });
}

function bindUiEvents() {
  ui.refreshBtn?.addEventListener("click", () => {
    setStatus("Refreshing combat log...");
    void refreshFromRoom("Manual refresh").catch((error) => {
      console.warn("[Odyssey Combat Log] Unable to refresh log", error);
      setStatus(error?.message ?? "Unable to refresh combat log.");
    });
  });

  ui.clearViewBtn?.addEventListener("click", () => {
    clearLocalView();
  });

  ui.restoreViewBtn?.addEventListener("click", () => {
    restoreLocalView();
  });

  ui.clearBtn?.addEventListener("click", () => {
    setStatus("Clearing shared combat log...");
    void clearSharedLog().catch((error) => {
      console.warn("[Odyssey Combat Log] Unable to clear log", error);
      setStatus(error?.message ?? "Unable to clear shared combat log.");
    });
  });

  ui.sizeCompactBtn?.addEventListener("click", () => {
    void applyWindowSize(COMPACT_WINDOW_SIZE, "Compact size").catch((error) => {
      console.warn("[Odyssey Combat Log] Unable to apply compact size", error);
      setStatus(error?.message ?? "Unable to resize window.");
    });
  });

  ui.sizeDefaultBtn?.addEventListener("click", () => {
    void applyWindowSize(DEFAULT_WINDOW_SIZE, "Default size").catch((error) => {
      console.warn("[Odyssey Combat Log] Unable to apply default size", error);
      setStatus(error?.message ?? "Unable to resize window.");
    });
  });

  ui.sizeLargeBtn?.addEventListener("click", () => {
    void applyWindowSize(LARGE_WINDOW_SIZE, "Large size").catch((error) => {
      console.warn("[Odyssey Combat Log] Unable to apply large size", error);
      setStatus(error?.message ?? "Unable to resize window.");
    });
  });

  ui.widthDownBtn?.addEventListener("click", () => {
    void applyWindowSize(
      { width: windowSize.width - WIDTH_STEP, height: windowSize.height },
      "Width updated",
    ).catch((error) => {
      console.warn("[Odyssey Combat Log] Unable to reduce width", error);
      setStatus(error?.message ?? "Unable to resize window.");
    });
  });

  ui.widthUpBtn?.addEventListener("click", () => {
    void applyWindowSize(
      { width: windowSize.width + WIDTH_STEP, height: windowSize.height },
      "Width updated",
    ).catch((error) => {
      console.warn("[Odyssey Combat Log] Unable to increase width", error);
      setStatus(error?.message ?? "Unable to resize window.");
    });
  });

  ui.heightDownBtn?.addEventListener("click", () => {
    void applyWindowSize(
      { width: windowSize.width, height: windowSize.height - HEIGHT_STEP },
      "Height updated",
    ).catch((error) => {
      console.warn("[Odyssey Combat Log] Unable to reduce height", error);
      setStatus(error?.message ?? "Unable to resize window.");
    });
  });

  ui.heightUpBtn?.addEventListener("click", () => {
    void applyWindowSize(
      { width: windowSize.width, height: windowSize.height + HEIGHT_STEP },
      "Height updated",
    ).catch((error) => {
      console.warn("[Odyssey Combat Log] Unable to increase height", error);
      setStatus(error?.message ?? "Unable to resize window.");
    });
  });
}

OBR.onReady(async () => {
  try {
    const [name, role] = await Promise.all([
      OBR.player.getName(),
      OBR.player.getRole(),
    ]);
    viewerName = name ?? viewerName;
    viewerRole = role ?? viewerRole;
    windowSize = loadStoredWindowSize();
    localViewCutoffId = loadStoredViewCutoff();

    bindUiEvents();
    bindPointerResize();
    renderWindowSize();
    renderEntries();
    await applyWindowSize(windowSize, "Restored size");
    await refreshFromRoom("Initial sync");

    OBR.player.onChange((player) => {
      viewerName = player?.name ?? viewerName;
      viewerRole = player?.role ?? viewerRole;
      renderHeader();
    });

    OBR.broadcast.onMessage(DEBUG_BROADCAST_CHANNEL, (event) => {
      const payload = event?.data;
      if (!payload || typeof payload !== "object") return;

      if (payload.type === "debug-clear") {
        sharedEntries = [];
        localViewCutoffId = 0;
        saveViewCutoff();
        setSyncState("Live clear");
        setStatus("Shared combat log cleared.");
        renderEntries();
        return;
      }

      if (payload.type !== "debug-entry") return;

      sharedEntries = mergeDebugEntries([payload.entry], sharedEntries);
      setSyncState("Live event");
      setStatus("Received a live Odyssey combat event.");
      renderEntries();
    });

    OBR.room.onMetadataChange((metadata) => {
      sharedEntries = sanitizeDebugEntries(metadata?.[DEBUG_LOG_KEY]);
      if (!sharedEntries.length && localViewCutoffId) {
        localViewCutoffId = 0;
        saveViewCutoff();
      }
      setSyncState("Room update");
      setStatus("Room log updated.");
      renderEntries();
    });

    roomRefreshTimer = window.setInterval(() => {
      void refreshFromRoom("Fallback poll", { quiet: true }).catch((error) => {
        console.warn("[Odyssey Combat Log] Poll refresh failed", error);
      });
    }, POLL_INTERVAL_MS);
  } catch (error) {
    console.error("[Odyssey Combat Log] Initialization failed", error);
    setStatus(error?.message ?? "Combat log extension failed to initialize.");
  }
});
