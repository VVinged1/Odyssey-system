import {
  BODY_ORDER,
  OBR,
  canPlayerControlToken,
  clamp,
  ensureOverlayForToken,
  formatOverlayText,
  getAvailableWeapons,
  getBodyTotals,
  getCharacterName,
  getOdysseyData,
  getTrackerData,
  isCharacterToken,
  isTrackedCharacter,
  sortCharacters,
  syncTrackedOverlays,
  updateTrackerData,
} from "./shared.js";
import { resolveAttack, rollDice } from "./odyssey_rules.js";

const DEBUG_LOG_KEY = "com.codex.body-hp/debugLog";

const ui = {
  roleBadge: document.getElementById("roleBadge"),
  refreshBtn: document.getElementById("refreshBtn"),
  syncBtn: document.getElementById("syncBtn"),
  statusBox: document.getElementById("statusBox"),
  selectionHint: document.getElementById("selectionHint"),
  selectedTokenPanel: document.getElementById("selectedTokenPanel"),
  debugConsole: document.getElementById("debugConsole"),
  trackedSection: document.getElementById("trackedSection"),
  trackedCount: document.getElementById("trackedCount"),
  trackedList: document.getElementById("trackedList"),
  allTokensSection: document.getElementById("allTokensSection"),
  allCount: document.getElementById("allCount"),
  allTokensList: document.getElementById("allTokensList"),
};

let playerRole = "PLAYER";
let playerId = "";
let playerName = "";
let sceneItems = [];
let selectionIds = [];
let activeTokenId = null;
let debugEntries = [];
const inputAutosaveTimers = new Map();
let selectionPollTimer = null;

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
    .slice(0, 30);
}

function setStatus(message, kind = "info") {
  ui.statusBox.textContent = message;
  ui.statusBox.className = `status ${kind}`;
  console[kind === "error" ? "error" : "log"](`[Body HP] ${message}`);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function getCharacters() {
  return sortCharacters(sceneItems.filter(isCharacterToken));
}

function getTrackedCharacters() {
  return getCharacters().filter(isTrackedCharacter);
}

function getControllableCharacters() {
  return getTrackedCharacters().filter((token) =>
    canPlayerControlToken(playerRole, playerId, token)
  );
}

function getCharacterById(tokenId) {
  return getCharacters().find((item) => item.id === tokenId) ?? null;
}

function resolveActiveTokenId() {
  const characters = getCharacters();
  const selectedCharacterId = selectionIds.find((id) =>
    characters.some((character) => character.id === id)
  );

  if (selectedCharacterId) return selectedCharacterId;
  if (activeTokenId && characters.some((character) => character.id === activeTokenId)) {
    return activeTokenId;
  }

  const firstTracked = getTrackedCharacters()[0];
  if (firstTracked) return firstTracked.id;

  return characters[0]?.id ?? null;
}

function isEditable() {
  return playerRole === "GM";
}

function canUseToken(token) {
  return canPlayerControlToken(playerRole, playerId, token);
}

function canEditTokenData(token) {
  return canUseToken(token);
}

async function initializeCharacterToken(tokenId) {
  const token = getCharacterById(tokenId);
  if (!token || !isCharacterToken(token)) return;
  await updateTrackerData(tokenId, (current) => current);
  await ensureOverlayForToken(tokenId);
}

function resolveDefaultTargetTokenId(attackerId) {
  const otherSelected = selectionIds.find((id) => id !== attackerId);
  if (otherSelected) return otherSelected;
  const fallback = getTrackedCharacters().find((token) => token.id !== attackerId);
  return fallback?.id ?? "";
}

async function pushDebugEntry(title, body, kind = "info") {
  const nextEntries = [
    {
      id: Date.now(),
      title,
      body,
      kind,
      timestamp: new Date().toLocaleTimeString(),
    },
    ...debugEntries,
  ].slice(0, 30);

  debugEntries = nextEntries;
  renderDebugConsole();
  await OBR.room.setMetadata({
    [DEBUG_LOG_KEY]: nextEntries,
  });
}

function renderDebugConsole() {
  if (!debugEntries.length) {
    ui.debugConsole.innerHTML = `
      <div class="hint-box">
        <div class="field-label">Current viewer</div>
        <pre class="console-output">Name: ${escapeHtml(playerName || "Unknown")}
Player ID: ${escapeHtml(playerId || "Unavailable")}

Roll debug will appear here after Attack or Roll Dice.</pre>
      </div>`;
    return;
  }

  ui.debugConsole.innerHTML = debugEntries
    .map(
      (entry) => `
        <div class="debug-entry">
          <div class="debug-head">
            <div class="debug-title">${escapeHtml(entry.title)}</div>
            <div class="muted">${escapeHtml(entry.timestamp)}</div>
          </div>
          <pre class="console-output">${escapeHtml(entry.body)}</pre>
        </div>`
    )
    .join("");
}

async function loadSharedDebugConsole() {
  const metadata = await OBR.room.getMetadata();
  debugEntries = sanitizeDebugEntries(metadata?.[DEBUG_LOG_KEY]);
}

function arraysEqual(left, right) {
  if (left.length !== right.length) return false;
  return left.every((value, index) => value === right[index]);
}

function startSelectionPolling() {
  if (selectionPollTimer) {
    clearInterval(selectionPollTimer);
  }

  selectionPollTimer = setInterval(() => {
    void OBR.player
      .getSelection()
      .then((selection) => selection ?? [])
      .then((selection) => {
        if (!arraysEqual(selectionIds, selection)) {
          return syncState();
        }
        return null;
      })
      .catch((error) => {
        console.warn("[Body HP] Selection polling failed", error);
      });
  }, 350);
}

function formatAttackDebug({
  attackerName,
  targetName,
  targetPart,
  attackSkillName,
  attackSkillValue,
  weaponDamage,
  attackBonuses,
  attackPenalties,
  defenseBonuses,
  defensePenalties,
  targetParry,
  targetArmor,
  result,
  beforeHp,
  afterHp,
}) {
  const lines = [
    `Attacker: ${attackerName}`,
    `Target: ${targetName}`,
    `Target Part: ${targetPart}`,
    "",
    `Attack Roll: ${result.attackRoll}`,
    `Attack Skill: ${attackSkillName} (${attackSkillValue} -> ${attackSkillValue * 10})`,
    `Attack Bonuses: ${attackBonuses}`,
    `Attack Penalties: ${attackPenalties}`,
    `Attack Total: ${result.attackTotal}`,
    "",
    `Defense Roll: ${result.defenseRoll}`,
    `Target Parry: ${targetParry} -> ${targetParry * 10}`,
    `Defense Bonuses: ${defenseBonuses}`,
    `Defense Penalties: ${defensePenalties}`,
    `Defense Total: ${result.defenseTotal}`,
    "",
    `Weapon Damage: ${weaponDamage}`,
    `Target Armor: ${targetArmor}`,
    `Final Attack: ${result.damage?.totalAttack ?? result.attackTotal}`,
    `Final Defense: ${result.damage?.totalDefense ?? result.defenseTotal}`,
    `Outcome: ${result.outcome}`,
    `Damage Label: ${result.damage?.label ?? "No damage"}`,
    `HP Change: ${beforeHp} -> ${afterHp}`,
  ];
  return lines.join("\n");
}

function formatDiceDebug({ tokenName, result }) {
  return [
    `Actor: ${tokenName}`,
    `Dice: d${result.sides}`,
    `Raw Roll: ${result.roll}`,
    `Modifier: ${result.modifier}`,
    `Total: ${result.total}`,
  ].join("\n");
}

function renderOwnerFields(data, disabledAttr) {
  return `
    <div class="preview-box">
      <div class="field-label">Odyssey ownership</div>
      <div class="hint-box">
        <div class="field-label">Current viewer</div>
        <pre class="console-output">Name: ${escapeHtml(playerName || "Unknown")}
Player ID: ${escapeHtml(playerId || "Unavailable")}</pre>
      </div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Owner Player ID</span>
          <input type="text" value="${escapeHtml(data.odyssey.owner.playerId)}" data-action="set-owner" data-field="playerId" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Owner Name</span>
          <input type="text" value="${escapeHtml(data.odyssey.owner.playerName)}" data-action="set-owner" data-field="playerName" ${disabledAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" class="secondary" data-action="assign-owner-to-me" ${disabledAttr}>Assign To Current Viewer</button>
      </div>
    </div>
  `;
}

function renderOdysseyStats(token, data, disabledAttr) {
  const skillInputs = Object.entries(data.odyssey.skills)
    .map(
      ([key, value]) => `
        <label class="field-stack">
          <span class="field-label">${escapeHtml(key)}</span>
          <input type="number" min="0" max="10" value="${value}" data-action="set-odyssey-skill" data-skill="${escapeHtml(key)}" ${disabledAttr}>
        </label>`
    )
    .join("");

  const primaryWeapon = getAvailableWeapons(token, "melee")[0] ?? { name: "Default", damage: 0 };

  return `
    <div class="preview-box">
      <div class="field-label">Odyssey stats</div>
      <div class="form-grid">${skillInputs}</div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Parry</span>
          <input type="number" min="0" max="10" value="${data.odyssey.attributes.Parry}" data-action="set-odyssey-attribute" data-attribute="Parry" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Primary weapon name</span>
          <input type="text" value="${escapeHtml(primaryWeapon.name)}" data-action="set-weapon-name" data-weapon-index="0" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Default weapon damage</span>
          <input type="number" min="-99" max="99" value="${primaryWeapon.damage}" data-action="set-weapon-damage" data-weapon-index="0" ${disabledAttr}>
        </label>
      </div>
    </div>
  `;
}

function renderOdysseyActions(token, data, tokenLocked) {
  const targetCharacters = getCharacters().filter((item) => item.id !== token.id);
  const defaultTargetId = resolveDefaultTargetTokenId(token.id);
  const disabledAttr = tokenLocked || !targetCharacters.length ? "disabled" : "";
  const skillOptions = Object.entries(data.odyssey.skills)
    .map(([key, value]) => `<option value="${escapeHtml(key)}">${escapeHtml(key)} (${value})</option>`)
    .join("");
  const defaultWeapon = getAvailableWeapons(token, "melee")[0] ?? { damage: 0 };

  return `
    <div class="preview-box">
      <div class="field-label">Odyssey actions</div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Attack skill</span>
          <select data-attack-field="skill">${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Target token</span>
          <select data-attack-field="targetTokenId" ${disabledAttr}>
            ${targetCharacters
              .map(
                (target) =>
                  `<option value="${target.id}" ${target.id === defaultTargetId ? "selected" : ""}>${escapeHtml(
                    getCharacterName(target)
                  )}</option>`
              )
              .join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Target body part</span>
          <select data-attack-field="targetPart">
            ${BODY_ORDER.map((part) => `<option value="${part}">${part}</option>`).join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon damage</span>
          <input type="number" value="${defaultWeapon.damage}" data-attack-field="weaponDamage" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack bonuses</span>
          <input type="number" value="0" data-attack-field="attackBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack penalties</span>
          <input type="number" value="0" data-attack-field="attackPenalties" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense bonuses</span>
          <input type="number" value="0" data-attack-field="defenseBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense penalties</span>
          <input type="number" value="0" data-attack-field="defensePenalties" ${disabledAttr}>
        </label>
      </div>
      <div class="muted">${
        targetCharacters.length
          ? "Attack goes from the selected attacker token to the chosen target token."
          : "Add at least two character tokens to perform an attack."
      }</div>
      <div class="row row-gap">
        <button type="button" class="success" data-action="perform-attack" ${disabledAttr}>Attack</button>
      </div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Dice sides</span>
          <input type="number" min="2" max="1000" value="20" data-roll-field="dice" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Modifier</span>
          <input type="number" value="0" data-roll-field="modifier" ${disabledAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-dice" ${disabledAttr}>Roll Dice</button>
      </div>
    </div>
  `;
}

function renderSelectedToken() {
  activeTokenId = resolveActiveTokenId();
  const token = getCharacterById(activeTokenId);

  if (!token) {
    ui.selectionHint.textContent = "No character token selected";
    ui.selectedTokenPanel.innerHTML =
      '<div class="empty">Add a character token to the map from Owlbear Rodeo Characters, then select it.</div>';
    return;
  }

  const tracked = isTrackedCharacter(token);
  const data = getTrackerData(token);
  const odyssey = getOdysseyData(token);
  const totals = getBodyTotals(data);
  const selected = selectionIds.includes(token.id);
  const tokenLocked = !canUseToken(token);
  const lastRollText = data.lastRoll
    ? escapeHtml(data.lastRoll.summary || "Last roll recorded")
    : "No rolls synced yet";

  ui.selectionHint.textContent = selected ? "Selected on map" : "Showing current focus";

  const fieldDisabled = !canEditTokenData(token) ? "disabled" : "";
  const odysseyOwnerDisabled = !isEditable() ? "disabled" : "";

  ui.selectedTokenPanel.innerHTML = `
    <div class="selected-card">
      <div class="selected-head">
        <div>
          <div class="token-name">${escapeHtml(getCharacterName(token))}</div>
          <div class="token-meta">${escapeHtml(token.id.slice(0, 8))} - ${
            tracked ? "Initialized" : "Auto-init on selection"
          } - ${tokenLocked ? "Read only" : "Controllable"}</div>
        </div>
        <div class="row row-gap">
          <button type="button" data-action="focus-token" class="secondary">Select On Map</button>
        </div>
      </div>

      <div class="summary-strip">
        <div class="stat-chip">
          <span class="chip-label">Body HP</span>
          <span class="chip-value">${totals.current}/${totals.max}</span>
        </div>
        <div class="stat-chip">
          <span class="chip-label">Owner</span>
          <span class="chip-value">${escapeHtml(odyssey.owner.playerName || odyssey.owner.playerId || "Unassigned")}</span>
        </div>
      </div>

      ${renderOwnerFields({ odyssey }, odysseyOwnerDisabled)}
      ${renderOdysseyStats(token, { odyssey }, fieldDisabled)}
      ${renderOdysseyActions(token, { odyssey }, tokenLocked)}

      <div class="preview-box">
        <div class="field-label">Last roll summary</div>
        <pre class="console-output">${lastRollText}</pre>
      </div>

      <div class="body-table-wrap">
        <table class="body-table">
          <thead>
            <tr>
              <th>Part</th>
              <th>Current</th>
              <th>Max</th>
              <th>Armor</th>
            </tr>
          </thead>
          <tbody>
            ${BODY_ORDER.map((partName) => {
              const part = data.body[partName];
              return `
                <tr>
                  <td class="part-name">${escapeHtml(partName)}</td>
                  <td>
                    <div class="inline-stepper">
                      <button type="button" data-action="change-part" data-part="${escapeHtml(
                        partName
                      )}" data-field="current" data-delta="-1" ${fieldDisabled}>-</button>
                      <input type="number" min="0" max="${part.max}" value="${part.current}" data-action="set-field" data-part="${escapeHtml(
                        partName
                      )}" data-field="current" ${fieldDisabled}>
                      <button type="button" data-action="change-part" data-part="${escapeHtml(
                        partName
                      )}" data-field="current" data-delta="1" ${fieldDisabled}>+</button>
                    </div>
                  </td>
                  <td>
                    <input class="compact-input" type="number" min="0" max="99" value="${part.max}" data-action="set-field" data-part="${escapeHtml(
                      partName
                    )}" data-field="max" ${fieldDisabled}>
                  </td>
                  <td>
                    <input class="compact-input" type="number" min="0" max="99" value="${part.armor}" data-action="set-field" data-part="${escapeHtml(
                      partName
                    )}" data-field="armor" ${fieldDisabled}>
                  </td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      </div>

      <div class="preview-box">
        <div class="field-label">Overlay preview</div>
        <pre>${escapeHtml(formatOverlayText(data))}</pre>
      </div>
    </div>`;
}

function renderTrackedList() {
  const trackedCharacters = getTrackedCharacters();
  ui.trackedCount.textContent = String(trackedCharacters.length);

  if (!trackedCharacters.length) {
    ui.trackedList.innerHTML =
      '<div class="empty">No initialized characters yet. Click a token on the map to initialize it automatically.</div>';
    return;
  }

  ui.trackedList.innerHTML = trackedCharacters
    .map((token) => {
      const data = getTrackerData(token);
      const totals = getBodyTotals(data);
      const controllable = canUseToken(token);
      return `
        <button type="button" class="list-item${
          token.id === activeTokenId ? " active" : ""
        }" data-action="select-character" data-token-id="${token.id}">
          <div class="list-item-head">
            <span>${escapeHtml(getCharacterName(token))}</span>
            <span class="pill hp">${totals.current}/${totals.max}</span>
          </div>
          <div class="list-item-sub">${controllable ? "Playable" : "Read only"}</div>
        </button>`;
    })
    .join("");
}

function renderAllCharacters() {
  const characters = getCharacters();
  ui.allCount.textContent = String(characters.length);

  if (!characters.length) {
    ui.allTokensList.innerHTML =
      '<div class="empty">No character tokens are on the scene yet.</div>';
    return;
  }

  ui.allTokensList.innerHTML = characters
    .map((token) => {
      const tracked = isTrackedCharacter(token);
      const controllable = canUseToken(token);
      return `
        <div class="token-row${token.id === activeTokenId ? " active" : ""}">
          <div>
            <div class="token-row-name">${escapeHtml(getCharacterName(token))}</div>
            <div class="token-row-sub">${escapeHtml(token.id.slice(0, 8))} - ${controllable ? "Playable" : "Read only"}</div>
          </div>
          <div class="row row-gap">
            <button type="button" class="secondary" data-action="select-character" data-token-id="${
              token.id
            }">Select</button>
            <span class="pill ${tracked ? "hp" : "armor"}">${
              tracked ? "Initialized" : "Ready"
            }</span>
          </div>
        </div>`;
    })
    .join("");
}

function render() {
  ui.roleBadge.textContent = playerRole === "GM" ? "GM" : "PLAYER";
  ui.trackedSection.classList.toggle("hidden", playerRole !== "GM");
  ui.allTokensSection.classList.toggle("hidden", playerRole !== "GM");
  renderSelectedToken();
  renderDebugConsole();
  if (playerRole === "GM") {
    renderTrackedList();
    renderAllCharacters();
  }
}

async function syncState(showToast = false) {
  const [role, id, name, items, selection] = await Promise.all([
    OBR.player.getRole(),
    OBR.player.getId(),
    OBR.player.getName(),
    OBR.scene.items.getItems(),
    OBR.player.getSelection(),
  ]);

  playerRole = role;
  playerId = id;
  playerName = name;
  sceneItems = items;
  selectionIds = selection ?? [];

  const selectedCharacterId = selectionIds.find((selectionId) =>
    sceneItems.some((item) => item.id === selectionId && isCharacterToken(item))
  );
  if (selectedCharacterId) {
    await initializeCharacterToken(selectedCharacterId);
    sceneItems = await OBR.scene.items.getItems();
  }

  render();

  if (showToast) {
    setStatus(
      `Loaded ${getCharacters().length} character token(s), ${getTrackedCharacters().length} tracked.`,
      "success"
    );
  }
}

async function selectCharacter(tokenId) {
  activeTokenId = tokenId;
  await OBR.player.select([tokenId], true);
  await initializeCharacterToken(tokenId);
  sceneItems = await OBR.scene.items.getItems();
  render();
}

async function changeBodyField(partName, field, delta) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can edit this token.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    const part = next.body[partName];
    if (!part) return next;

    if (field === "current") {
      part.current = clamp(part.current + delta, 0, part.max);
    } else if (field === "max") {
      part.max = clamp(part.max + delta, 0, 99);
      part.current = clamp(part.current, 0, part.max);
    } else if (field === "armor") {
      part.armor = clamp(part.armor + delta, 0, 99);
    }

    return next;
  });
  await ensureOverlayForToken(token.id);
  await syncState();
}

async function setBodyField(partName, field, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can edit this token.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    const part = next.body[partName];
    if (!part) return next;

    const numericValue = clamp(Number(value) || 0, 0, 99);
    if (field === "current") {
      part.current = clamp(numericValue, 0, part.max);
    } else if (field === "max") {
      part.max = numericValue;
      part.current = clamp(part.current, 0, part.max);
    } else if (field === "armor") {
      part.armor = numericValue;
    }

    return next;
  });
  await ensureOverlayForToken(token.id);
  await syncState();
}

async function setOwnerField(field, value) {
  if (!isEditable()) {
    setStatus("Only the GM can assign token owners.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey ??= structuredClone(getTrackerData(token).odyssey);
    next.odyssey.owner ??= { playerId: "", playerName: "" };
    next.odyssey.owner[field] = String(value || "").trim();
    return next;
  });
  await ensureOverlayForToken(token.id);
  await syncState();
}

async function setOdysseySkill(skill, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can edit this token.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.skills[skill] = clamp(Number(value) || 0, 0, 10);
    return next;
  });
  await syncState();
}

async function setOdysseyAttribute(attribute, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can edit this token.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.attributes[attribute] = clamp(Number(value) || 0, 0, 10);
    return next;
  });
  await syncState();
}

async function setWeaponDamage(index, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can edit this token.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.weapons.melee ??= [];
    if (!next.odyssey.weapons.melee[index]) {
      next.odyssey.weapons.melee[index] = { name: "Default", damage: 0 };
    }
    next.odyssey.weapons.melee[index].damage = clamp(Number(value) || 0, -99, 99);
    return next;
  });
  await syncState();
}

async function setWeaponName(index, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can edit this token.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.weapons.melee ??= [];
    if (!next.odyssey.weapons.melee[index]) {
      next.odyssey.weapons.melee[index] = { name: "Default", damage: 0 };
    }
    next.odyssey.weapons.melee[index].name = String(value || "").trim() || "Default";
    return next;
  });
  await syncState();
}

async function autosaveDraftField(draft) {
  const token = getCharacterById(draft.tokenId);
  if (!token) return;

  if (draft.action === "set-owner") {
    if (!isEditable()) return;
    await updateTrackerData(token.id, (current) => {
      const next = structuredClone(current);
      next.odyssey ??= structuredClone(getTrackerData(token).odyssey);
      next.odyssey.owner ??= { playerId: "", playerName: "" };
      next.odyssey.owner[draft.field] = String(draft.value || "").trim();
      return next;
    });
    return;
  }

  if (!canEditTokenData(token)) return;

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);

    if (draft.action === "set-odyssey-skill") {
      next.odyssey.skills[draft.skill] = clamp(Number(draft.value) || 0, 0, 10);
      return next;
    }

    if (draft.action === "set-odyssey-attribute") {
      next.odyssey.attributes[draft.attribute] = clamp(Number(draft.value) || 0, 0, 10);
      return next;
    }

    if (draft.action === "set-weapon-damage") {
      next.odyssey.weapons.melee ??= [];
      if (!next.odyssey.weapons.melee[draft.weaponIndex]) {
        next.odyssey.weapons.melee[draft.weaponIndex] = { name: "Default", damage: 0 };
      }
      next.odyssey.weapons.melee[draft.weaponIndex].damage = clamp(Number(draft.value) || 0, -99, 99);
      return next;
    }

    if (draft.action === "set-weapon-name") {
      next.odyssey.weapons.melee ??= [];
      if (!next.odyssey.weapons.melee[draft.weaponIndex]) {
        next.odyssey.weapons.melee[draft.weaponIndex] = { name: "Default", damage: 0 };
      }
      next.odyssey.weapons.melee[draft.weaponIndex].name = String(draft.value || "").trim() || "Default";
      return next;
    }

    if (draft.action === "set-field") {
      const part = next.body[draft.partName];
      if (!part) return next;
      const numericValue = clamp(Number(draft.value) || 0, 0, 99);
      if (draft.field === "current") {
        part.current = clamp(numericValue, 0, part.max);
      } else if (draft.field === "max") {
        part.max = numericValue;
        part.current = clamp(part.current, 0, part.max);
      } else if (draft.field === "armor") {
        part.armor = numericValue;
      }
      return next;
    }

    return next;
  });

  if (draft.action === "set-field") {
    await ensureOverlayForToken(token.id);
  }
}

function queueInputAutosave(draft) {
  const key = [
    draft.tokenId,
    draft.action,
    draft.field ?? "",
    draft.skill ?? "",
    draft.attribute ?? "",
    draft.weaponIndex ?? "",
    draft.partName ?? "",
  ].join("|");
  const existing = inputAutosaveTimers.get(key);
  if (existing) {
    clearTimeout(existing);
  }

  const timeoutId = setTimeout(() => {
    inputAutosaveTimers.delete(key);
    void autosaveDraftField(draft).catch((error) => {
      console.warn("[Body HP] Autosave failed", error);
    });
  }, 250);

  inputAutosaveTimers.set(key, timeoutId);
}

function getActionFieldValue(selector) {
  const tokenPanel = ui.selectedTokenPanel;
  const field = tokenPanel.querySelector(selector);
  if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) {
    return "";
  }
  return field.value;
}

async function performAttack() {
  const attacker = getCharacterById(activeTokenId);
  if (!attacker) {
    setStatus("Select an attacker token first.", "error");
    return;
  }
  if (!canUseToken(attacker)) {
    setStatus("You cannot roll for this attacker token.", "error");
    return;
  }

  const targetTokenId =
    getActionFieldValue('[data-attack-field="targetTokenId"]') ||
    resolveDefaultTargetTokenId(attacker.id);
  const target = getCharacterById(targetTokenId);
  if (!target) {
    setStatus("Choose a valid target token.", "error");
    return;
  }
  if (target.id === attacker.id) {
    setStatus("Attacker and target must be different tokens.", "error");
    return;
  }

  const attackerData = getTrackerData(attacker);
  const attackerOdyssey = getOdysseyData(attacker);
  const targetData = getTrackerData(target);
  const targetOdyssey = getOdysseyData(target);
  const skillName = getActionFieldValue('[data-attack-field="skill"]');
  const targetPart = getActionFieldValue('[data-attack-field="targetPart"]');
  const weaponDamage = Number(getActionFieldValue('[data-attack-field="weaponDamage"]')) || 0;
  const attackBonuses = Number(getActionFieldValue('[data-attack-field="attackBonuses"]')) || 0;
  const attackPenalties = Number(getActionFieldValue('[data-attack-field="attackPenalties"]')) || 0;
  const defenseBonuses = Number(getActionFieldValue('[data-attack-field="defenseBonuses"]')) || 0;
  const defensePenalties = Number(getActionFieldValue('[data-attack-field="defensePenalties"]')) || 0;
  const targetArmor = targetData.body[targetPart]?.armor ?? 0;
  const beforeHp = targetData.body[targetPart]?.current ?? 0;
  const targetParry = targetOdyssey.attributes.Parry ?? 0;

  const result = resolveAttack({
    attackSkill: attackerOdyssey.skills[skillName] ?? 0,
    weaponDamage,
    defenseBonuses,
    defensePenalties,
    attackBonuses,
    attackPenalties,
    parry: targetParry,
    targetPart,
    targetArmor,
  });
  const afterHp = result.hit
    ? clamp(beforeHp + result.bodyDelta, 0, targetData.body[targetPart]?.max ?? beforeHp)
    : beforeHp;

  await updateTrackerData(attacker.id, (current) => {
    const next = structuredClone(current);
    next.lastRoll = {
      eventId: 0,
      actorName: playerName || "Owlbear Player",
      summary: `${getCharacterName(attacker)} -> ${getCharacterName(target)}: ${result.summary}`,
      outcome: result.outcome,
      total: result.attackTotal,
      targetPart: result.targetPart,
      timestamp: new Date().toISOString(),
      source: "owlbear-extension",
    };
    next.history = [next.lastRoll, ...(next.history ?? [])].slice(0, 12);
    return next;
  });

  await updateTrackerData(target.id, (current) => {
    const next = structuredClone(current);
    if (result.hit && next.body[result.targetPart]) {
      next.body[result.targetPart].current = clamp(
        next.body[result.targetPart].current + result.bodyDelta,
        0,
        next.body[result.targetPart].max,
      );
    }
    next.lastRoll = {
      eventId: 0,
      actorName: getCharacterName(attacker),
      summary: result.summary,
      outcome: result.outcome,
      total: result.attackTotal,
      targetPart: result.targetPart,
      timestamp: new Date().toISOString(),
      source: "owlbear-extension",
    };
    next.history = [next.lastRoll, ...(next.history ?? [])].slice(0, 12);
    return next;
  });

  await ensureOverlayForToken(attacker.id);
  await ensureOverlayForToken(target.id);
  await pushDebugEntry(
    `${getCharacterName(attacker)} attacks ${getCharacterName(target)}`,
    formatAttackDebug({
      attackerName: getCharacterName(attacker),
      targetName: getCharacterName(target),
      targetPart,
      attackSkillName: skillName,
      attackSkillValue: attackerOdyssey.skills[skillName] ?? 0,
      weaponDamage,
      attackBonuses,
      attackPenalties,
      defenseBonuses,
      defensePenalties,
      targetParry,
      targetArmor,
      result,
      beforeHp,
      afterHp,
    }),
    result.hit ? "success" : "info",
  );
  await syncState();
  setStatus(`${getCharacterName(attacker)} -> ${getCharacterName(target)}: ${result.summary}`, result.hit ? "success" : "info");
}

async function performRollDice() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot roll for this token.", "error");
    return;
  }

  const dice = Number(getActionFieldValue('[data-roll-field="dice"]')) || 20;
  const modifier = Number(getActionFieldValue('[data-roll-field="modifier"]')) || 0;
  const result = rollDice(dice, modifier);

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.lastRoll = {
      eventId: 0,
      actorName: playerName || "Owlbear Player",
      summary: `Rolled d${result.sides}: ${result.roll}${modifier ? ` ${modifier >= 0 ? "+" : ""}${modifier}` : ""} = ${result.total}`,
      outcome: "roll",
      total: result.total,
      targetPart: "",
      timestamp: new Date().toISOString(),
      source: "owlbear-extension",
    };
    next.history = [next.lastRoll, ...(next.history ?? [])].slice(0, 12);
    return next;
  });

  await ensureOverlayForToken(token.id);
  await pushDebugEntry(`${getCharacterName(token)} rolls dice`, formatDiceDebug({
    tokenName: getCharacterName(token),
    result,
  }), "success");
  await syncState();
  setStatus(`d${result.sides} rolled ${result.total}.`, "success");
}

function bindUiEvents() {
  ui.refreshBtn.addEventListener("click", () => {
    void syncState(true).catch((error) => {
      setStatus(error?.message ?? "Refresh failed.", "error");
    });
  });

  ui.syncBtn.addEventListener("click", () => {
    if (!isEditable()) {
      setStatus("Only the GM can rebuild overlays.", "error");
      return;
    }

    void syncTrackedOverlays()
      .then(() => syncState())
      .then(() => {
        setStatus("Tracked overlays rebuilt.", "success");
      })
      .catch((error) => {
        setStatus(error?.message ?? "Overlay rebuild failed.", "error");
      });
  });

  document.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    const actionNode = target.closest("[data-action]");
    if (!(actionNode instanceof HTMLElement)) return;

    const action = actionNode.dataset.action;
    const tokenId = actionNode.dataset.tokenId;
    const partName = actionNode.dataset.part;
    const field = actionNode.dataset.field;
    const delta = Number(actionNode.dataset.delta ?? 0);

    if (action === "select-character" && tokenId) {
      void selectCharacter(tokenId).catch((error) => {
        setStatus(error?.message ?? "Unable to select token.", "error");
      });
    }

    if (action === "assign-owner-to-me") {
      void Promise.all([
        setOwnerField("playerId", playerId),
        setOwnerField("playerName", playerName),
      ]).catch((error) => {
        setStatus(error?.message ?? "Unable to assign current viewer.", "error");
      });
    }

    if (action === "focus-token" && activeTokenId) {
      void selectCharacter(activeTokenId).catch((error) => {
        setStatus(error?.message ?? "Unable to focus token.", "error");
      });
    }

    if (action === "change-part" && partName && field) {
      void changeBodyField(partName, field, delta).catch((error) => {
        setStatus(error?.message ?? "Unable to update body value.", "error");
      });
    }

    if (action === "perform-attack") {
      void performAttack().catch((error) => {
        setStatus(error?.message ?? "Unable to resolve attack.", "error");
      });
    }

    if (action === "perform-roll-dice") {
      void performRollDice().catch((error) => {
        setStatus(error?.message ?? "Unable to roll dice.", "error");
      });
    }
  });

  document.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement)) return;

    if (target.dataset.action === "set-owner") {
      const field = target.dataset.field;
      if (!field) return;
      void setOwnerField(field, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save owner.", "error");
      });
      return;
    }

    if (target.dataset.action === "set-odyssey-skill") {
      const skill = target.dataset.skill;
      if (!skill) return;
      void setOdysseySkill(skill, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save skill.", "error");
      });
      return;
    }

    if (target.dataset.action === "set-odyssey-attribute") {
      const attribute = target.dataset.attribute;
      if (!attribute) return;
      void setOdysseyAttribute(attribute, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save attribute.", "error");
      });
      return;
    }

    if (target.dataset.action === "set-weapon-damage") {
      const index = Number(target.dataset.weaponIndex ?? 0);
      void setWeaponDamage(index, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save weapon damage.", "error");
      });
      return;
    }

    if (target.dataset.action === "set-weapon-name") {
      const index = Number(target.dataset.weaponIndex ?? 0);
      void setWeaponName(index, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save weapon name.", "error");
      });
      return;
    }

    if (target.dataset.action !== "set-field") return;

    const partName = target.dataset.part;
    const field = target.dataset.field;
    if (!partName || !field) return;

    void setBodyField(partName, field, target.value).catch((error) => {
      setStatus(error?.message ?? "Unable to save field.", "error");
    });
  });

  document.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (!activeTokenId) return;

    if (target.dataset.action === "set-owner") {
      const field = target.dataset.field;
      if (!field) return;
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-owner",
        field,
        value: target.value,
      });
      return;
    }

    if (target.dataset.action === "set-odyssey-skill") {
      const skill = target.dataset.skill;
      if (!skill) return;
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-odyssey-skill",
        skill,
        value: target.value,
      });
      return;
    }

    if (target.dataset.action === "set-odyssey-attribute") {
      const attribute = target.dataset.attribute;
      if (!attribute) return;
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-odyssey-attribute",
        attribute,
        value: target.value,
      });
      return;
    }

    if (target.dataset.action === "set-weapon-damage") {
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-weapon-damage",
        weaponIndex: Number(target.dataset.weaponIndex ?? 0),
        value: target.value,
      });
      return;
    }

    if (target.dataset.action === "set-weapon-name") {
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-weapon-name",
        weaponIndex: Number(target.dataset.weaponIndex ?? 0),
        value: target.value,
      });
      return;
    }

    if (target.dataset.action === "set-field") {
      const partName = target.dataset.part;
      const field = target.dataset.field;
      if (!partName || !field) return;
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-field",
        partName,
        field,
        value: target.value,
      });
    }
  });
}

OBR.onReady(async () => {
  try {
    bindUiEvents();
    await loadSharedDebugConsole();
    await syncState(true);
    startSelectionPolling();
    setStatus(
      "Ready. Select a character token on the map to edit it here.",
      "info"
    );

    OBR.scene.items.onChange((items) => {
      sceneItems = items;
      render();
    });

    OBR.player.onChange((player) => {
      playerRole = player.role;
      playerId = player.id ?? playerId;
      playerName = player.name ?? playerName;
      selectionIds = player.selection ?? [];
      const selectedCharacterId = selectionIds.find((selectionId) =>
        sceneItems.some((item) => item.id === selectionId && isCharacterToken(item))
      );
      if (selectedCharacterId) {
        void initializeCharacterToken(selectedCharacterId)
          .then(() => OBR.scene.items.getItems())
          .then((items) => {
            sceneItems = items;
            render();
          })
          .catch((error) => {
            console.warn("[Body HP] Auto-init on selection failed", error);
            render();
          });
        return;
      }
      render();
    });

    OBR.room.onMetadataChange((metadata) => {
      debugEntries = sanitizeDebugEntries(metadata?.[DEBUG_LOG_KEY]);
      renderDebugConsole();
    });
  } catch (error) {
    setStatus(error?.message ?? "Extension failed to initialize.", "error");
  }
});
