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
  setTrackedState,
  sortCharacters,
  syncTrackedOverlays,
  updateTrackerData,
} from "./shared.js";
import { resolveAttack, rollDice } from "./odyssey_rules.js";

const ui = {
  roleBadge: document.getElementById("roleBadge"),
  refreshBtn: document.getElementById("refreshBtn"),
  syncBtn: document.getElementById("syncBtn"),
  statusBox: document.getElementById("statusBox"),
  selectionHint: document.getElementById("selectionHint"),
  selectedTokenPanel: document.getElementById("selectedTokenPanel"),
  trackedCount: document.getElementById("trackedCount"),
  trackedList: document.getElementById("trackedList"),
  allCount: document.getElementById("allCount"),
  allTokensList: document.getElementById("allTokensList"),
};

let playerRole = "PLAYER";
let playerId = "";
let playerName = "";
let sceneItems = [];
let selectionIds = [];
let activeTokenId = null;

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

function renderOwnerFields(data, disabledAttr) {
  return `
    <div class="preview-box">
      <div class="field-label">Odyssey ownership</div>
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

  const weaponOptions = getAvailableWeapons(token, "melee")
    .map(
      (weapon) =>
        `<option value="${escapeHtml(weapon.name)}">${escapeHtml(weapon.name)} (${weapon.damage >= 0 ? "+" : ""}${weapon.damage})</option>`
    )
    .join("");

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
          <span class="field-label">Default weapon damage</span>
          <input type="number" min="-99" max="99" value="${getAvailableWeapons(token, "melee")[0]?.damage ?? 0}" data-action="set-weapon-damage" data-weapon-index="0" ${disabledAttr}>
        </label>
      </div>
      <div class="field-label">Primary melee weapon</div>
      <div class="muted">${weaponOptions ? "Default weapon can be edited above." : "No weapons configured yet."}</div>
    </div>
  `;
}

function renderOdysseyActions(token, data, tokenLocked) {
  const disabledAttr = tokenLocked ? "disabled" : "";
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
    ? escapeHtml(
        [
          data.lastRoll.actorName || "Unknown",
          data.lastRoll.total != null ? `roll ${data.lastRoll.total}` : "",
          data.lastRoll.outcome || "",
          data.lastRoll.targetPart ? `target ${data.lastRoll.targetPart}` : "",
          data.lastRoll.summary || "",
        ]
          .filter(Boolean)
          .join(" | ")
      )
    : "No rolls synced yet";

  ui.selectionHint.textContent = selected ? "Selected on map" : "Showing current focus";

  const toggleButton = isEditable()
    ? `<button type="button" data-action="toggle-tracking" class="${
        tracked ? "danger" : "success"
      }">${tracked ? "Remove Tracking" : "Track Character"}</button>`
    : "";

  const damageDisabled = !tracked || !isEditable() ? "disabled" : "";
  const fieldDisabled = !tracked || !isEditable() ? "disabled" : "";
  const odysseyOwnerDisabled = !tracked || !isEditable() ? "disabled" : "";

  ui.selectedTokenPanel.innerHTML = `
    <div class="selected-card">
      <div class="selected-head">
        <div>
          <div class="token-name">${escapeHtml(getCharacterName(token))}</div>
          <div class="token-meta">${escapeHtml(token.id.slice(0, 8))} - ${
            tracked ? "Tracked" : "Not tracked"
          } - ${tokenLocked ? "Read only" : "Controllable"}</div>
        </div>
        <div class="row row-gap">
          <button type="button" data-action="focus-token" class="secondary">Select On Map</button>
          ${toggleButton}
        </div>
      </div>

      <div class="summary-strip">
        <div class="stat-chip">
          <span class="chip-label">Body HP</span>
          <span class="chip-value">${totals.current}/${totals.max}</span>
        </div>
        <div class="stat-chip">
          <span class="chip-label">Minor</span>
          <span class="chip-value">${data.minor}</span>
        </div>
        <div class="stat-chip">
          <span class="chip-label">Serious</span>
          <span class="chip-value">${data.serious}</span>
        </div>
        <div class="stat-chip">
          <span class="chip-label">Owner</span>
          <span class="chip-value">${escapeHtml(odyssey.owner.playerName || odyssey.owner.playerId || "Unassigned")}</span>
        </div>
      </div>

      ${renderOwnerFields({ odyssey }, odysseyOwnerDisabled)}
      ${renderOdysseyStats(token, { odyssey }, odysseyOwnerDisabled)}
      ${renderOdysseyActions(token, { odyssey }, !tracked || tokenLocked)}

      <div class="preview-box">
        <div class="field-label">Bridge identity</div>
        <div class="damage-grid">
          <div class="damage-card">
            <div class="field-label">Player ID</div>
            <input class="compact-input" type="text" value="${escapeHtml(
              data.identity.playerId
            )}" data-action="set-identity" data-field="playerId" ${fieldDisabled}>
          </div>
          <div class="damage-card">
            <div class="field-label">Character ID</div>
            <input class="compact-input" type="text" value="${escapeHtml(
              data.identity.characterId
            )}" data-action="set-identity" data-field="characterId" ${fieldDisabled}>
          </div>
        </div>
        <div class="field-label">Last synced roll</div>
        <pre>${lastRollText}</pre>
      </div>

      <div class="damage-grid">
        <div class="damage-card">
          <div class="field-label">Minor damage dots</div>
          <div class="stepper">
            <button type="button" data-action="change-damage" data-kind="minor" data-delta="-1" ${damageDisabled}>-</button>
            <span>${data.minor}/4</span>
            <button type="button" data-action="change-damage" data-kind="minor" data-delta="1" ${damageDisabled}>+</button>
          </div>
        </div>
        <div class="damage-card">
          <div class="field-label">Serious damage bars</div>
          <div class="stepper">
            <button type="button" data-action="change-damage" data-kind="serious" data-delta="-1" ${damageDisabled}>-</button>
            <span>${data.serious}/2</span>
            <button type="button" data-action="change-damage" data-kind="serious" data-delta="1" ${damageDisabled}>+</button>
          </div>
        </div>
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
      '<div class="empty">No tracked characters yet. A GM can track them from this panel or from the token context menu.</div>';
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
          <div class="list-item-sub">Minor ${data.minor} - Serious ${data.serious} - ${controllable ? "Playable" : "Read only"}</div>
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
            ${
              isEditable()
                ? `<button type="button" class="${
                    tracked ? "danger" : "success"
                  }" data-action="toggle-track-specific" data-token-id="${token.id}">${
                    tracked ? "Untrack" : "Track"
                  }</button>`
                : `<span class="pill ${tracked ? "hp" : "armor"}">${
                    tracked ? "Tracked" : "Viewer"
                  }</span>`
            }
          </div>
        </div>`;
    })
    .join("");
}

function render() {
  ui.roleBadge.textContent = playerRole === "GM" ? "GM" : "PLAYER";
  renderSelectedToken();
  renderTrackedList();
  renderAllCharacters();
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
  render();
}

async function toggleTracking(tokenId) {
  if (!isEditable()) {
    setStatus("Only the GM can change tracked characters.", "error");
    return;
  }

  const token = getCharacterById(tokenId);
  if (!token) return;

  const enableTracking = !isTrackedCharacter(token);
  await setTrackedState(tokenId, enableTracking);
  activeTokenId = tokenId;
  await syncState();
  setStatus(
    enableTracking
      ? `Tracking enabled for ${getCharacterName(token)}.`
      : `Tracking removed for ${getCharacterName(token)}.`,
    "success"
  );
}

async function changeDamage(kind, delta) {
  if (!isEditable()) {
    setStatus("Only the GM can edit damage values.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token || !isTrackedCharacter(token)) {
    setStatus("Select a tracked character first.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => ({
    ...current,
    [kind]: clamp(
      (current[kind] ?? 0) + delta,
      0,
      kind === "minor" ? 4 : 2
    ),
  }));
  await ensureOverlayForToken(token.id);
  await syncState();
}

async function changeBodyField(partName, field, delta) {
  if (!isEditable()) {
    setStatus("Only the GM can edit body values.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token || !isTrackedCharacter(token)) {
    setStatus("Select a tracked character first.", "error");
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
  if (!isEditable()) {
    setStatus("Only the GM can edit body values.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token || !isTrackedCharacter(token)) {
    setStatus("Select a tracked character first.", "error");
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

async function setIdentityField(field, value) {
  if (!isEditable()) {
    setStatus("Only the GM can edit bridge identity.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token || !isTrackedCharacter(token)) {
    setStatus("Select a tracked character first.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.identity ??= { playerId: "", characterId: "" };
    next.identity[field] = String(value || "").trim();
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
  if (!token || !isTrackedCharacter(token)) {
    setStatus("Select a tracked character first.", "error");
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
  if (!isEditable()) {
    setStatus("Only the GM can edit Odyssey skills.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token || !isTrackedCharacter(token)) {
    setStatus("Select a tracked character first.", "error");
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
  if (!isEditable()) {
    setStatus("Only the GM can edit Odyssey attributes.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token || !isTrackedCharacter(token)) {
    setStatus("Select a tracked character first.", "error");
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
  if (!isEditable()) {
    setStatus("Only the GM can edit weapon damage.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token || !isTrackedCharacter(token)) {
    setStatus("Select a tracked character first.", "error");
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

function getActionFieldValue(selector) {
  const tokenPanel = ui.selectedTokenPanel;
  const field = tokenPanel.querySelector(selector);
  if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) {
    return "";
  }
  return field.value;
}

async function performAttack() {
  const token = getCharacterById(activeTokenId);
  if (!token || !isTrackedCharacter(token)) {
    setStatus("Select a tracked character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot roll for this token.", "error");
    return;
  }

  const trackerData = getTrackerData(token);
  const odyssey = getOdysseyData(token);
  const skillName = getActionFieldValue('[data-attack-field="skill"]');
  const targetPart = getActionFieldValue('[data-attack-field="targetPart"]');
  const weaponDamage = Number(getActionFieldValue('[data-attack-field="weaponDamage"]')) || 0;
  const attackBonuses = Number(getActionFieldValue('[data-attack-field="attackBonuses"]')) || 0;
  const attackPenalties = Number(getActionFieldValue('[data-attack-field="attackPenalties"]')) || 0;
  const defenseBonuses = Number(getActionFieldValue('[data-attack-field="defenseBonuses"]')) || 0;
  const defensePenalties = Number(getActionFieldValue('[data-attack-field="defensePenalties"]')) || 0;
  const targetArmor = trackerData.body[targetPart]?.armor ?? 0;

  const result = resolveAttack({
    attackSkill: odyssey.skills[skillName] ?? 0,
    weaponDamage,
    defenseBonuses,
    defensePenalties,
    attackBonuses,
    attackPenalties,
    parry: odyssey.attributes.Parry ?? 0,
    targetPart,
    targetArmor,
  });

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.lastRoll = {
      eventId: 0,
      actorName: playerName || "Owlbear Player",
      summary: result.summary,
      outcome: result.outcome,
      total: result.attackTotal,
      targetPart: result.targetPart,
      timestamp: new Date().toISOString(),
      source: "owlbear-extension",
    };
    next.history = [next.lastRoll, ...(next.history ?? [])].slice(0, 12);

    if (result.hit && next.body[result.targetPart]) {
      next.body[result.targetPart].current = clamp(
        next.body[result.targetPart].current + result.bodyDelta,
        0,
        next.body[result.targetPart].max,
      );
    }
    return next;
  });

  await ensureOverlayForToken(token.id);
  await syncState();
  setStatus(result.summary, result.hit ? "success" : "info");
}

async function performRollDice() {
  const token = getCharacterById(activeTokenId);
  if (!token || !isTrackedCharacter(token)) {
    setStatus("Select a tracked character first.", "error");
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

    if (action === "toggle-track-specific" && tokenId) {
      void toggleTracking(tokenId).catch((error) => {
        setStatus(error?.message ?? "Unable to toggle tracking.", "error");
      });
    }

    if (action === "toggle-tracking" && activeTokenId) {
      void toggleTracking(activeTokenId).catch((error) => {
        setStatus(error?.message ?? "Unable to toggle tracking.", "error");
      });
    }

    if (action === "focus-token" && activeTokenId) {
      void selectCharacter(activeTokenId).catch((error) => {
        setStatus(error?.message ?? "Unable to focus token.", "error");
      });
    }

    if (action === "change-damage") {
      const kind = actionNode.dataset.kind;
      if (!kind) return;
      void changeDamage(kind, delta).catch((error) => {
        setStatus(error?.message ?? "Unable to update damage.", "error");
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

    if (target.dataset.action === "set-identity") {
      const field = target.dataset.field;
      if (!field) return;
      void setIdentityField(field, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save identity.", "error");
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
}

OBR.onReady(async () => {
  try {
    bindUiEvents();
    await syncState(true);
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
      render();
    });
  } catch (error) {
    setStatus(error?.message ?? "Extension failed to initialize.", "error");
  }
});
