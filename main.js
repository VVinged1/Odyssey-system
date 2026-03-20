import { buildShape } from "@owlbear-rodeo/sdk";
import {
  APPLIED_SKILL_CATEGORY,
  BODY_ORDER,
  COMBAT_SKILL_CATEGORY,
  DEFAULT_ODYSSEY_SKILLS,
  MELEE_SKILL_NAME,
  OBR,
  PARRY_SKILL_NAME,
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
const DEBUG_BROADCAST_CHANNEL = "com.codex.body-hp/debug";
const TARGET_PICK_TOOL_ID = "com.codex.body-hp/attack-target-picker";
const TARGET_PICK_MODE_ID = "pick-target";
const TARGET_HIGHLIGHT_KEY = "com.codex.body-hp/local-attack-target";
const EXTENSION_ICON_URL = new URL("./icon.svg", window.location.href).href;
const TARGET_PICK_CURSOR = `url("data:image/svg+xml,${encodeURIComponent(
  `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32">
    <circle cx="16" cy="16" r="3.25" fill="#ef4444" stroke="#7f1d1d" stroke-width="1.5"/>
    <path d="M16 2v8M16 22v8M2 16h8M22 16h8" stroke="#ef4444" stroke-width="2.5" stroke-linecap="round"/>
    <path d="M16 6v4M16 22v4M6 16h4M22 16h4" stroke="#7f1d1d" stroke-width="1.2" stroke-linecap="round"/>
  </svg>`,
)}") 16 16, crosshair`;
const CORE_COMBAT_SKILLS = Object.keys(DEFAULT_ODYSSEY_SKILLS);
const ATTACK_ONLY_EXCLUDED_SKILLS = new Set([PARRY_SKILL_NAME]);
const ATTRIBUTE_FIELDS = [
  ["Strength", "Strength"],
  ["Agility", "Agility"],
  ["Reaction", "Reaction"],
  ["Endurance", "Endurance"],
  ["Perception", "Perception"],
  ["Intelligence", "Intelligence"],
  ["Charisma", "Charisma"],
  ["Willpower", "Willpower"],
  ["Magic", "Magic"],
];

const ATTRIBUTE_UI_FIELDS = [
  ["Strength", "Strength"],
  ["Agility", "Agility"],
  ["Reaction", "Reaction"],
  ["Endurance", "Endurance"],
  ["Perception", "Perception"],
  ["Intelligence", "Intelligence"],
  ["Charisma", "Charisma"],
  ["Willpower", "Willpower"],
  ["Magic", "Magic"],
];

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
let playerColor = "#facc15";
let sceneItems = [];
let selectionIds = [];
let activeTokenId = null;
let debugEntries = [];
let partyPlayers = [];
let gmPrivateEntries = [];
const collapsibleSectionState = new Map();
const attackFormDrafts = new Map();
const inputAutosaveTimers = new Map();
let selectionPollTimer = null;
const targetPickState = {
  active: false,
  attackerTokenId: null,
  previousToolId: "",
  previousModeId: undefined,
  toolReady: false,
  restoring: false,
};

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

function mergeDebugEntries(...entryGroups) {
  const merged = new Map();

  for (const group of entryGroups) {
    for (const entry of sanitizeDebugEntries(group)) {
      merged.set(entry.id, entry);
    }
  }

  return [...merged.values()]
    .sort((left, right) => Number(right.id) - Number(left.id))
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

function getSkillCategory(odyssey, skillName) {
  return odyssey?.skillCategories?.[skillName] === COMBAT_SKILL_CATEGORY
    ? COMBAT_SKILL_CATEGORY
    : APPLIED_SKILL_CATEGORY;
}

function getSkillStrengthBonusFlag(odyssey, skillName) {
  return Boolean(odyssey?.skillStrengthBonuses?.[skillName]);
}

function getSortedSkillEntries(odyssey) {
  return Object.entries(odyssey?.skills ?? {}).sort(([left], [right]) =>
    left.localeCompare(right),
  );
}

function getCombatSkillEntries(odyssey) {
  return getSortedSkillEntries(odyssey).filter(
    ([skillName]) => getSkillCategory(odyssey, skillName) === COMBAT_SKILL_CATEGORY,
  );
}

function getAttackSkillEntries(odyssey) {
  return getCombatSkillEntries(odyssey).filter(
    ([skillName]) => !ATTACK_ONLY_EXCLUDED_SKILLS.has(skillName),
  );
}

function getAppliedSkillEntries(odyssey) {
  return getSortedSkillEntries(odyssey).filter(
    ([skillName]) => getSkillCategory(odyssey, skillName) === APPLIED_SKILL_CATEGORY,
  );
}

function buildSkillOptions(skillEntries, selectedValue = "") {
  return skillEntries
    .map(
      ([key, value]) => `<option value="${escapeHtml(key)}" ${
        key === selectedValue ? "selected" : ""
      }>${escapeHtml(key)} (${value})</option>`,
    )
    .join("");
}

function buildGroupedSkillOptions(odyssey, selectedValue = "") {
  const combatOptions = buildSkillOptions(getCombatSkillEntries(odyssey), selectedValue);
  const appliedOptions = buildSkillOptions(getAppliedSkillEntries(odyssey), selectedValue);

  return [
    combatOptions ? `<optgroup label="Combat">${combatOptions}</optgroup>` : "",
    appliedOptions ? `<optgroup label="Applied">${appliedOptions}</optgroup>` : "",
  ]
    .filter(Boolean)
    .join("");
}

function getTransientFieldKey(field) {
  if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) {
    return "";
  }

  if (field.dataset.attackField) return `attack:${field.dataset.attackField}`;
  if (field.dataset.rollField) return `roll:${field.dataset.rollField}`;
  if (field.dataset.rollCharField) return `roll-char:${field.dataset.rollCharField}`;
  if (field.dataset.rollSkillField) return `roll-skill:${field.dataset.rollSkillField}`;
  if (field.dataset.gmRollField) return `gm-roll:${field.dataset.gmRollField}`;
  if (field.dataset.skillField) return `new-skill:${field.dataset.skillField}`;

  if (field.dataset.action === "select-owner-player") return "owner";
  if (field.dataset.action === "set-odyssey-skill") return `skill:${field.dataset.skill ?? ""}`;
  if (field.dataset.action === "set-skill-strength-bonus") {
    return `skill-strength:${field.dataset.skill ?? ""}`;
  }
  if (field.dataset.action === "set-odyssey-attribute") {
    return `attribute:${field.dataset.attribute ?? ""}`;
  }
  if (field.dataset.action === "set-field") {
    return `part:${field.dataset.part ?? ""}:${field.dataset.field ?? ""}`;
  }

  return "";
}

function shouldPreserveFieldValue(fieldKey, focusedKey) {
  return (
    fieldKey === focusedKey ||
    fieldKey.startsWith("attack:") ||
    fieldKey.startsWith("roll:") ||
    fieldKey.startsWith("roll-char:") ||
    fieldKey.startsWith("roll-skill:") ||
    fieldKey.startsWith("gm-roll:") ||
    fieldKey.startsWith("new-skill:")
  );
}

function captureSelectedPanelState() {
  if (!activeTokenId || !ui.selectedTokenPanel.childElementCount) return null;

  let focusedKey = "";
  let selectionStart = null;
  let selectionEnd = null;
  const activeField = document.activeElement;
  if (
    (activeField instanceof HTMLInputElement || activeField instanceof HTMLSelectElement) &&
    ui.selectedTokenPanel.contains(activeField)
  ) {
    focusedKey = getTransientFieldKey(activeField);
    if (activeField instanceof HTMLInputElement && activeField.type !== "number") {
      selectionStart = activeField.selectionStart;
      selectionEnd = activeField.selectionEnd;
    }
  }

  const fields = [];
  ui.selectedTokenPanel.querySelectorAll("input, select").forEach((field) => {
    if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) return;
    const key = getTransientFieldKey(field);
    if (!key || !shouldPreserveFieldValue(key, focusedKey)) return;
    fields.push([
      key,
      field instanceof HTMLInputElement && field.type === "checkbox"
        ? { kind: "checkbox", checked: field.checked }
        : { kind: "value", value: field.value },
    ]);
  });

  return {
    tokenId: activeTokenId,
    fields,
    focusedKey,
    selectionStart,
    selectionEnd,
  };
}

function restoreSelectedPanelState(panelState) {
  if (!panelState || panelState.tokenId !== activeTokenId) return;

  const fieldValues = new Map(panelState.fields ?? []);
  let focusedField = null;

  ui.selectedTokenPanel.querySelectorAll("input, select").forEach((field) => {
    if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) return;
    const key = getTransientFieldKey(field);
    if (!key || !fieldValues.has(key)) return;

    const nextValue = fieldValues.get(key);
    if (
      field instanceof HTMLInputElement &&
      field.type === "checkbox" &&
      nextValue &&
      typeof nextValue === "object" &&
      nextValue.kind === "checkbox"
    ) {
      field.checked = Boolean(nextValue.checked);
    } else {
      const normalizedValue =
        nextValue && typeof nextValue === "object" && nextValue.kind === "value"
          ? nextValue.value
          : typeof nextValue === "string"
            ? nextValue
            : "";
      if (field instanceof HTMLSelectElement) {
        const hasMatchingOption = Array.from(field.options).some(
          (option) => option.value === normalizedValue,
        );
        if (!hasMatchingOption) return;
      }
      field.value = normalizedValue;
    }
    if (key === panelState.focusedKey) {
      focusedField = field;
    }
  });

  if (!focusedField) return;

  focusedField.focus({ preventScroll: true });
  if (
    focusedField instanceof HTMLInputElement &&
    typeof panelState.selectionStart === "number" &&
    typeof panelState.selectionEnd === "number"
  ) {
    try {
      focusedField.setSelectionRange(panelState.selectionStart, panelState.selectionEnd);
    } catch (error) {
      console.warn("[Body HP] Unable to restore cursor position", error);
    }
  }
}

function getSortedPartyPlayers() {
  return [...partyPlayers].sort((left, right) =>
    String(left?.name ?? "").localeCompare(String(right?.name ?? ""))
  );
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
  if (!token || !isCharacterToken(token)) return false;
  const shouldInitialize = !isTrackedCharacter(token);
  if (shouldInitialize) {
    await updateTrackerData(tokenId, (current) => current);
  }
  await ensureOverlayForToken(tokenId);
  return shouldInitialize;
}

function resolveDefaultTargetTokenId(attackerId) {
  const visibleTargets = getCharacters().filter(
    (token) => token.id !== attackerId && token.visible !== false,
  );
  const otherSelected = selectionIds.find(
    (id) => id !== attackerId && visibleTargets.some((token) => token.id === id),
  );
  if (otherSelected) return otherSelected;
  const fallback = visibleTargets[0];
  return fallback?.id ?? "";
}

async function pushDebugEntry(title, body, kind = "info") {
  const entry = {
    id: Date.now() * 1000 + Math.floor(Math.random() * 1000),
    title,
    body,
    kind,
    timestamp: new Date().toLocaleTimeString(),
  };
  const metadata = await OBR.room.getMetadata();
  const nextEntries = mergeDebugEntries([entry], metadata?.[DEBUG_LOG_KEY], debugEntries);
  debugEntries = nextEntries;
  renderDebugConsole();
  await OBR.broadcast.sendMessage(
    DEBUG_BROADCAST_CHANNEL,
    { type: "debug-entry", entry },
    { destination: "REMOTE" },
  );

  if (playerRole === "GM") {
    await OBR.room.setMetadata({
      [DEBUG_LOG_KEY]: nextEntries,
    });
  }
}

function renderDebugConsole() {
  if (!debugEntries.length) {
    ui.debugConsole.innerHTML = `
      <div class="hint-box">
        <div class="field-label">Current viewer</div>
        <pre class="console-output">Name: ${escapeHtml(playerName || "Unknown")}
Player ID: ${escapeHtml(playerId || "Unavailable")}

Actions from all players and the GM will appear here after rolls and attacks.</pre>
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
  debugEntries = mergeDebugEntries(metadata?.[DEBUG_LOG_KEY], debugEntries);
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
  }, 200);
}

function formatAttackDebug({
  attackerName,
  targetName,
  targetPart,
  attackSkillName,
  attackSkillValue,
  weaponDamage,
  strengthBonus,
  attackBonuses,
  manualAttackPenalties,
  automaticTargetPenalty,
  totalAttackPenalties,
  defenseBonuses,
  defensePenalties,
  targetParry,
  targetArmor,
  result,
  beforeHp,
  afterHp,
  beforeMinor,
  afterMinor,
  beforeSerious,
  afterSerious,
  critApplied,
}) {
  const accuracyTable = formatTextTable(
    ["Side", "Attacking", "Defending"],
    [
      [
        "Accuracy",
        `${result.attackRoll} + ${attackSkillValue * 10} + ${attackBonuses} - ${totalAttackPenalties} = ${result.attackTotal}`,
        `${result.defenseRoll} + ${targetParry * 10} + ${defenseBonuses} - ${defensePenalties} = ${result.defenseTotal}`,
      ],
      [
        "Damage",
        `${result.attackTotal} + ${weaponDamage}`,
        `${result.defenseTotal} + ${targetArmor}`,
      ],
      [
        "Result",
        `${result.damage?.totalAttack ?? result.attackTotal}`,
        `${result.damage?.totalDefense ?? result.defenseTotal}`,
      ],
    ],
  );

  const damageTable = formatTextTable(
    ["Parameter", "Value"],
    [
      ["Attacker", attackerName],
      ["Target", `${targetName} -> ${targetPart}`],
      ["Attack Skill", `${attackSkillName} (${attackSkillValue})`],
      ["Strength Bonus", strengthBonus],
      ["Manual Attack Penalty", manualAttackPenalties],
      ["Auto Target Penalty", automaticTargetPenalty],
      ["Outcome", result.outcome],
      ["Damage Diff", result.damage?.damageDiff ?? 0],
      ["Damage Label", result.damage?.label ?? "No damage"],
      ["Applied Min/Sir/Crit", `${result.damage?.minor ?? 0} / ${result.damage?.serious ?? 0} / ${result.damage?.crit ?? 0}`],
      ["Converted Crit", critApplied],
      ["Crit State", `${beforeHp} -> ${afterHp}`],
      ["Minor State", `${beforeMinor} -> ${afterMinor}`],
      ["Serious State", `${beforeSerious} -> ${afterSerious}`],
    ],
  );

  return `${accuracyTable}\n\n${damageTable}`;
}

function projectPartDamage(part, damage) {
  const next = {
    current: Number(part?.current) || 0,
    max: Number(part?.max) || 0,
    armor: Number(part?.armor) || 0,
    minor: Number(part?.minor) || 0,
    serious: Number(part?.serious) || 0,
  };

  next.minor = Math.max(0, next.minor + (Number(damage?.minor) || 0));
  next.serious = Math.max(0, next.serious + (Number(damage?.serious) || 0));

  const promotedSerious = Math.floor(next.minor / 4);
  next.minor %= 4;
  next.serious += promotedSerious;

  const convertedCrit = Math.floor(next.serious / 2);
  next.serious %= 2;

  const directCrit = Math.max(0, Number(damage?.crit) || 0);
  const totalCrit = directCrit + convertedCrit;
  next.current = clamp(next.current - totalCrit, 0, next.max);

  return {
    ...next,
    critApplied: totalCrit,
  };
}

function formatDiceDebug({ tokenName, result }) {
  return formatTextTable(
    ["Parameter", "Value"],
    [
      ["Actor", tokenName],
      ["Dice", `${result.count}d${result.sides}`],
      ["Rolls", result.rolls.join(", ")],
      ["Subtotal", result.subtotal],
      ["Modifier", result.modifier],
      ["Total", result.total],
    ],
  );
}

function getCurrentPlayerColor() {
  return String(
    playerColor ||
      partyPlayers.find((player) => player?.id === playerId)?.color ||
      "#facc15",
  );
}

function getAutomaticTargetPenalty(targetPart) {
  if (targetPart === "Head") return 30;
  if (targetPart === "L.Arm" || targetPart === "R.Arm" || targetPart === "L.Leg" || targetPart === "R.Leg") {
    return 15;
  }
  return 0;
}

function formatRollCharDebug({ tokenName, attributeLabel, result }) {
  return [
    `Character: ${tokenName}`,
    `Characteristic: ${attributeLabel}`,
    `${result.result}`,
    "",
    formatTextTable(
      ["Roll", "Base Attribute", "Modifier", "Final Attribute"],
      [[result.roll, result.baseAttribute, result.modifier, result.finalAttribute]],
    ),
  ].join("\n");
}

function formatRollSkillDebug({ tokenName, skillName, result }) {
  return [
    `Character: ${tokenName}`,
    `Skill: ${skillName}`,
    `${result.result}`,
    "",
    formatTextTable(
      ["Parameter", "Value"],
      [
        ["First Roll", `${result.rollPrimary} + ${result.baseSkill * 10} + ${result.modifier} = ${result.totalPrimary}`],
        ["Second Roll", `${result.rollSecondary} = ${result.totalSecondary}`],
      ],
    ),
  ].join("\n");
}

function formatTextTable(headers, rows) {
  const normalizedHeaders = headers.map((cell) => String(cell ?? ""));
  const normalizedRows = rows.map((row) => row.map((cell) => String(cell ?? "")));
  const widths = normalizedHeaders.map((header, columnIndex) =>
    Math.max(
      header.length,
      ...normalizedRows.map((row) => (row[columnIndex] ?? "").length),
    ),
  );

  const renderBorder = (left, middle, right, fill) =>
    `${left}${widths.map((width) => fill.repeat(width + 2)).join(middle)}${right}`;
  const renderRow = (row) =>
    `│ ${row
      .map((cell, columnIndex) => String(cell ?? "").padEnd(widths[columnIndex], " "))
      .join(" │ ")} │`;

  return [
    renderBorder("╒", "╤", "╕", "═"),
    renderRow(normalizedHeaders),
    renderBorder("╞", "╪", "╡", "═"),
    ...normalizedRows.map(renderRow),
    renderBorder("╘", "╧", "╛", "═"),
  ].join("\n");
}

function getAttackDraft(token, data, targetCharacters) {
  const defaultWeapon = getAvailableWeapons(token, "melee")[0] ?? { damage: 0 };
  const stored = attackFormDrafts.get(token.id) ?? {};
  const combatSkillNames = getAttackSkillEntries(data.odyssey).map(([skillName]) => skillName);
  const fallbackSkill =
    combatSkillNames[0] ??
    CORE_COMBAT_SKILLS.find((key) => key in data.odyssey.skills) ??
    CORE_COMBAT_SKILLS[0];

  return {
    skill: combatSkillNames.includes(stored.skill)
      ? stored.skill
      : fallbackSkill,
    targetTokenId: targetCharacters.some((target) => target.id === stored.targetTokenId)
      ? stored.targetTokenId
      : resolveDefaultTargetTokenId(token.id),
    targetPart: BODY_ORDER.includes(stored.targetPart) ? stored.targetPart : "Torso",
    weaponDamage: stored.weaponDamage ?? String(defaultWeapon.damage),
    attackBonuses: stored.attackBonuses ?? "0",
    attackPenalties: stored.attackPenalties ?? "0",
    defenseBonuses: stored.defenseBonuses ?? "0",
    defensePenalties: stored.defensePenalties ?? "0",
  };
}

function saveAttackDraftValue(tokenId, field, value) {
  if (!tokenId || !field) return;
  const current = attackFormDrafts.get(tokenId) ?? {};
  attackFormDrafts.set(tokenId, {
    ...current,
    [field]: value,
  });
}

function isLocalTargetHighlight(item) {
  return Boolean(item?.metadata?.[TARGET_HIGHLIGHT_KEY]);
}

async function clearTargetHighlight() {
  const localItems = await OBR.scene.local.getItems();
  const highlightIds = localItems
    .filter(isLocalTargetHighlight)
    .map((item) => item.id);

  if (highlightIds.length) {
    await OBR.scene.local.deleteItems(highlightIds);
  }
}

async function buildTargetHighlightItem(targetToken) {
  let bounds = null;
  try {
    bounds = await OBR.scene.items.getItemBounds([targetToken.id]);
  } catch (error) {
    console.warn("[Body HP] Unable to read target bounds for highlight", error);
  }

  const width = Math.max(
    bounds?.width ?? 0,
    (targetToken.width || 140) * Math.abs(targetToken.scale?.x ?? 1),
    56,
  );
  const height = Math.max(
    bounds?.height ?? 0,
    (targetToken.height || 140) * Math.abs(targetToken.scale?.y ?? 1),
    56,
  );
  const diameter = Math.max(24, Math.min(width, height) - 4);
  const center = bounds?.center ?? targetToken.position;
  const position = {
    x: center.x - diameter / 2,
    y: center.y - diameter / 2,
  };

  return buildShape()
    .name(`Attack Target: ${getCharacterName(targetToken)}`)
    .shapeType("CIRCLE")
    .width(diameter)
    .height(diameter)
    .position(position)
    .rotation(0)
    .layer("ATTACHMENT")
    .locked(true)
    .disableHit(true)
    .fillColor(getCurrentPlayerColor())
    .fillOpacity(0.22)
    .strokeColor(getCurrentPlayerColor())
    .strokeOpacity(0)
    .strokeWidth(0)
    .metadata({
      [TARGET_HIGHLIGHT_KEY]: {
        playerId,
        playerName,
        attackerTokenId: activeTokenId ?? "",
        targetTokenId: targetToken.id,
      },
    })
    .build();
}

async function syncTargetHighlight() {
  const attacker = getCharacterById(activeTokenId);
  if (!attacker || !isCharacterToken(attacker)) {
    await clearTargetHighlight();
    return;
  }

  const targetCharacters = getCharacters().filter(
    (item) => item.id !== attacker.id && item.visible !== false,
  );
  const draft = getAttackDraft(attacker, getTrackerData(attacker), targetCharacters);
  const target = getCharacterById(draft.targetTokenId);
  const localItems = await OBR.scene.local.getItems();
  const existingHighlights = localItems.filter(isLocalTargetHighlight);

  if (!target || !isCharacterToken(target) || target.visible === false || target.id === attacker.id) {
    if (existingHighlights.length) {
      await OBR.scene.local.deleteItems(existingHighlights.map((item) => item.id));
    }
    return;
  }

  if (
    existingHighlights.length === 1 &&
    existingHighlights[0]?.metadata?.[TARGET_HIGHLIGHT_KEY]?.targetTokenId === target.id &&
    existingHighlights[0]?.metadata?.[TARGET_HIGHLIGHT_KEY]?.attackerTokenId === attacker.id
  ) {
    return;
  }

  if (existingHighlights.length) {
    await OBR.scene.local.deleteItems(existingHighlights.map((item) => item.id));
  }

  await OBR.scene.local.addItems([await buildTargetHighlightItem(target)]);
}

async function teardownTargetPickerTool() {
  if (!targetPickState.toolReady) return;

  try {
    await OBR.tool.removeMode(TARGET_PICK_MODE_ID);
  } catch (error) {
    console.warn("[Body HP] Unable to remove target picker mode", error);
  }

  try {
    await OBR.tool.remove(TARGET_PICK_TOOL_ID);
  } catch (error) {
    console.warn("[Body HP] Unable to remove target picker tool", error);
  }

  targetPickState.toolReady = false;
}

async function restorePreviousTool() {
  const previousToolId = targetPickState.previousToolId;
  const previousModeId = targetPickState.previousModeId;
  if (!previousToolId || previousToolId === TARGET_PICK_TOOL_ID) return;

  targetPickState.restoring = true;
  try {
    await OBR.tool.activateTool(previousToolId);
    if (previousModeId) {
      try {
        await OBR.tool.activateMode(previousToolId, previousModeId);
      } catch (error) {
        console.warn("[Body HP] Unable to restore previous tool mode", error);
      }
    }
  } finally {
    targetPickState.restoring = false;
  }
}

async function stopTargetPick(statusMessage = "", statusKind = "info") {
  const wasActive = targetPickState.active;
  targetPickState.active = false;
  targetPickState.attackerTokenId = null;
  render();

  if (wasActive) {
    await restorePreviousTool();
  }

  await teardownTargetPickerTool();

  targetPickState.previousToolId = "";
  targetPickState.previousModeId = undefined;

  if (statusMessage) {
    setStatus(statusMessage, statusKind);
  }
}

async function ensureTargetPickerTool() {
  if (targetPickState.toolReady) return;

  await OBR.tool.create({
    id: TARGET_PICK_TOOL_ID,
    icons: [{ icon: EXTENSION_ICON_URL, label: "Pick Attack Target" }],
    defaultMode: TARGET_PICK_MODE_ID,
  });

  await OBR.tool.createMode({
    id: TARGET_PICK_MODE_ID,
    icons: [{ icon: EXTENSION_ICON_URL, label: "Pick Attack Target" }],
    cursors: [{ cursor: TARGET_PICK_CURSOR }],
    onToolClick: async (_context, event) => {
      if (!targetPickState.active) return false;

      const attacker = getCharacterById(targetPickState.attackerTokenId);
      const target = event.target;

      if (!attacker || !isCharacterToken(attacker)) {
        await stopTargetPick("Select an attacker token first.", "error");
        return false;
      }

      if (!target || !isCharacterToken(target)) {
        setStatus("Click a visible character token to use it as target.", "error");
        return false;
      }

      if (target.visible === false) {
        setStatus("Hidden tokens cannot be targeted.", "error");
        return false;
      }

      if (target.id === attacker.id) {
        setStatus("Attacker and target must be different tokens.", "error");
        return false;
      }

      saveAttackDraftValue(attacker.id, "targetTokenId", target.id);
      render();
      await syncTargetHighlight();
      await stopTargetPick(`Target set to ${getCharacterName(target)}.`, "success");
      return false;
    },
    onKeyDown: (_context, event) => {
      if (event.key === "Escape" && targetPickState.active) {
        void stopTargetPick("Target picking cancelled.", "info");
      }
    },
    onDeactivate: () => {
      if (targetPickState.active && !targetPickState.restoring) {
        void stopTargetPick("Target picking cancelled.", "info");
      }
    },
  });

  targetPickState.toolReady = true;
}

async function startTargetPick() {
  const attacker = getCharacterById(activeTokenId);
  if (!attacker) {
    setStatus("Select an attacker token first.", "error");
    return;
  }
  if (!canUseToken(attacker)) {
    setStatus("You cannot roll for this attacker token.", "error");
    return;
  }

  const targetCharacters = getCharacters().filter(
    (item) => item.id !== attacker.id && item.visible !== false,
  );
  if (!targetCharacters.length) {
    setStatus("Add at least one visible target token.", "error");
    return;
  }

  if (targetPickState.active && targetPickState.attackerTokenId === attacker.id) {
    await stopTargetPick("Target picking cancelled.", "info");
    return;
  }

  if (targetPickState.active) {
    await stopTargetPick();
  }

  targetPickState.previousToolId = await OBR.tool.getActiveTool();
  targetPickState.previousModeId = await OBR.tool.getActiveToolMode();
  targetPickState.attackerTokenId = attacker.id;
  targetPickState.active = true;

  await ensureTargetPickerTool();
  await OBR.tool.activateTool(TARGET_PICK_TOOL_ID);
  await OBR.tool.activateMode(TARGET_PICK_TOOL_ID, TARGET_PICK_MODE_ID);
  render();
  setStatus("Click a visible target token on the map to assign it.", "info");
}

function renderCollapsibleSection(title, content, open = false, sectionKey = "") {
  const scopedSectionKey = `${activeTokenId ?? "global"}:${sectionKey || title}`;
  const resolvedOpen = collapsibleSectionState.has(scopedSectionKey)
    ? collapsibleSectionState.get(scopedSectionKey)
    : open;

  return `
    <details class="collapsible-block" data-section-key="${escapeHtml(scopedSectionKey)}" ${resolvedOpen ? "open" : ""}>
      <summary class="collapsible-title">${escapeHtml(title)}</summary>
      <div class="collapsible-body">${content}</div>
    </details>
  `;
}

function pushPrivateGmEntry(title, body) {
  gmPrivateEntries = [
    {
      id: Date.now(),
      title,
      body,
      timestamp: new Date().toLocaleTimeString(),
    },
    ...gmPrivateEntries,
  ].slice(0, 12);
}

function rollCharacterCheck(attributeValue, modifier = 0) {
  const baseAttribute = Number(attributeValue) || 0;
  const finalAttribute = Math.max(0, baseAttribute + (Number(modifier) || 0));
  const roll = Math.floor(Math.random() * 20) + 1;
  const result = roll <= finalAttribute ? "Check Passed" : "Check Failed";

  return {
    roll,
    baseAttribute,
    modifier: Number(modifier) || 0,
    finalAttribute,
    result,
  };
}

function rollSkillCheck(skillValue, modifier = 0) {
  const baseSkill = Number(skillValue) || 0;
  const rollPrimary = Math.floor(Math.random() * 100) + 1;
  const rollSecondary = Math.floor(Math.random() * 100) + 1;
  const totalPrimary = rollPrimary + baseSkill * 10 + (Number(modifier) || 0);
  const totalSecondary = rollSecondary;
  const result = totalPrimary > totalSecondary ? "Check Passed" : "Check Failed";

  return {
    rollPrimary,
    rollSecondary,
    baseSkill,
    modifier: Number(modifier) || 0,
    totalPrimary,
    totalSecondary,
    result,
  };
}

function renderOwnerFields(data, disabledAttr) {
  const playerOptions = [
    `<option value="">Unassigned</option>`,
    ...getSortedPartyPlayers().map(
      (player) => `
        <option value="${escapeHtml(player.id)}" ${
          data.odyssey.owner.playerId === player.id ? "selected" : ""
        }>${escapeHtml(player.name || player.id)}</option>`
    ),
  ].join("");

  return renderCollapsibleSection(
    "Ownership",
    `
      <div class="hint-box">
        <div class="field-label">Current viewer</div>
        <pre class="console-output">Name: ${escapeHtml(playerName || "Unknown")}
Player ID: ${escapeHtml(playerId || "Unavailable")}</pre>
      </div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Assigned Player</span>
          <select data-action="select-owner-player" ${disabledAttr}>${playerOptions}</select>
        </label>
      </div>
    `,
    false,
  );
}

function legacyRenderCharacteristicsBlock(data, disabledAttr) {
  const attributeInputs = ATTRIBUTE_FIELDS.map(
    ([key, label]) => `
      <label class="field-stack">
        <span class="field-label">${escapeHtml(label)}</span>
        <input type="number" min="0" max="15" value="${data.odyssey.attributes[key] ?? 0}" data-action="set-odyssey-attribute" data-attribute="${escapeHtml(key)}" ${disabledAttr}>
      </label>`
  ).join("");

  return renderCollapsibleSection(
    "Characteristics",
    `<div class="form-grid">${attributeInputs}</div>`,
    false,
  );
}

function legacyRenderSkillsBlock(data, disabledAttr) {
  const skillRows = Object.entries(data.odyssey.skills)
    .map(
      ([key, value]) => `
        <div class="skill-row">
          <div class="skill-name">${escapeHtml(key)}</div>
          <input type="number" min="0" max="10" value="${value}" data-action="set-odyssey-skill" data-skill="${escapeHtml(key)}" ${disabledAttr}>
          <button type="button" class="danger" data-action="remove-skill" data-skill="${escapeHtml(key)}" ${disabledAttr}>Remove</button>
        </div>`
    )
    .join("");

  return renderCollapsibleSection(
    "Навыки",
    `
      <div class="list">${skillRows || '<div class="empty">No skills yet.</div>'}</div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Название Навыка</span>
          <input type="text" data-skill-field="new-name" placeholder="New skill" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Значение</span>
          <input type="number" min="0" max="10" value="0" data-skill-field="new-value" ${disabledAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" class="secondary" data-action="add-skill" ${disabledAttr}>Add Skill</button>
      </div>
    `,
    false,
  );
}

function legacyRenderCombatBlock(token, data, tokenLocked) {
  const targetCharacters = getCharacters().filter((item) => item.id !== token.id);
  const defaultTargetId = resolveDefaultTargetTokenId(token.id);
  const disabledAttr = tokenLocked || !targetCharacters.length ? "disabled" : "";
  const skillOptions = Object.entries(data.odyssey.skills)
    .map(([key, value]) => `<option value="${escapeHtml(key)}">${escapeHtml(key)} (${value})</option>`)
    .join("");
  const defaultWeapon = getAvailableWeapons(token, "melee")[0] ?? { damage: 0 };

  return renderCollapsibleSection(
    "Атака",
    `
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
                  `<option value="${target.id}" ${target.id === draft.targetTokenId ? "selected" : ""}>${escapeHtml(
                    getCharacterName(target)
                  )}</option>`
              )
              .join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Target body part</span>
          <select data-attack-field="targetPart">
            ${BODY_ORDER.map(
              (part) =>
                `<option value="${part}" ${part === draft.targetPart ? "selected" : ""}>${part}</option>`
            ).join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon damage</span>
          <input type="number" value="${draft.weaponDamage}" data-attack-field="weaponDamage" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack bonuses</span>
          <input type="number" value="${draft.attackBonuses}" data-attack-field="attackBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack penalties</span>
          <input type="number" value="${draft.attackPenalties}" data-attack-field="attackPenalties" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense bonuses</span>
          <input type="number" value="${draft.defenseBonuses}" data-attack-field="defenseBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense penalties</span>
          <input type="number" value="${draft.defensePenalties}" data-attack-field="defensePenalties" ${disabledAttr}>
        </label>
      </div>
      <div class="muted">${
        targetCharacters.length
          ? "Attack goes from the selected attacker token to the chosen target token."
          : "Add at least two character tokens to perform an attack."
      }</div>
      <div class="muted">For Hand/Cold attacks, Strength above 10 adds bonus weapon damage automatically.</div>
      <div class="row row-gap">
        <button type="button" class="success" data-action="perform-attack" ${disabledAttr}>Attack</button>
      </div>
    `,
    true,
  );
}

function legacyRenderDiceBlock(token, data, tokenLocked) {
  const attributeOptions = ATTRIBUTE_FIELDS
    .filter(([key]) => key !== "Parry")
    .map(
      ([key, label]) =>
        `<option value="${escapeHtml(key)}">${escapeHtml(label)} (${data.odyssey.attributes[key] ?? 0})</option>`
    )
    .join("");
  const skillOptions = Object.entries(data.odyssey.skills)
    .map(([key, value]) => `<option value="${escapeHtml(key)}">${escapeHtml(key)} (${value})</option>`)
    .join("");
  const tokenLockedAttr = tokenLocked ? "disabled" : "";

  return renderCollapsibleSection(
    "Dice",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Dice sides</span>
          <input type="number" min="2" max="1000" value="20" data-roll-field="dice">
        </label>
        <label class="field-stack">
          <span class="field-label">Modifier</span>
          <input type="number" value="0" data-roll-field="modifier">
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-dice">Roll Dice</button>
      </div>

      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Roll_Char</span>
          <select data-roll-char-field="attribute" ${tokenLockedAttr}>${attributeOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Bonus / Penalty</span>
          <input type="number" value="0" data-roll-char-field="modifier" ${tokenLockedAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-char" ${tokenLockedAttr}>Roll Char</button>
      </div>

      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Roll_Skill</span>
          <select data-roll-skill-field="skill" ${tokenLockedAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Bonus / Penalty</span>
          <input type="number" value="0" data-roll-skill-field="modifier" ${tokenLockedAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-skill" ${tokenLockedAttr}>Roll Skill</button>
      </div>
    `,
    false,
  );
}

function legacyRenderPrivateGmDiceBlock() {
  if (!isEditable()) return "";

  const privateLog = gmPrivateEntries.length
    ? gmPrivateEntries
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
        .join("")
    : '<div class="empty">Private GM rolls will stay visible only here.</div>';

  return renderCollapsibleSection(
    "GM Private Dice",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Dice sides</span>
          <input type="number" min="2" max="1000" value="20" data-gm-roll-field="dice">
        </label>
        <label class="field-stack">
          <span class="field-label">Modifier</span>
          <input type="number" value="0" data-gm-roll-field="modifier">
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-gm-private-roll">GM Roll</button>
      </div>
      <div class="list">${privateLog}</div>
    `,
    false,
  );
}

function legacyRenderSelectedToken() {
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
      ${renderCharacteristicsBlock({ odyssey }, odysseyOwnerDisabled)}
      ${renderSkillsBlock({ odyssey }, odysseyOwnerDisabled)}
      ${renderCombatBlock(token, { odyssey }, tokenLocked)}
      ${renderDiceBlock(token, { odyssey }, tokenLocked)}
      ${renderPrivateGmDiceBlock()}
      ${renderCollapsibleSection(
        "Last roll summary",
        `<pre class="console-output">${lastRollText}</pre>`,
        false,
      )}
      ${renderCollapsibleSection(
        "Part",
        `
          <div class="body-table-wrap">
            <table class="body-table">
              <thead>
                <tr>
                  <th>Part</th>
                  <th>Crit</th>
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
        `,
        true,
      )}
      ${renderCollapsibleSection(
        "Overlay preview",
        `<pre class="console-output">${escapeHtml(formatOverlayText(data))}</pre>`,
        false,
      )}
    </div>`;
}

function renderCharacteristicsBlock(data, disabledAttr) {
  const attributeInputs = ATTRIBUTE_UI_FIELDS.map(
    ([key, label]) => `
      <label class="field-stack">
        <span class="field-label">${escapeHtml(label)}</span>
        <input type="number" min="0" max="15" value="${data.odyssey.attributes[key] ?? 0}" data-action="set-odyssey-attribute" data-attribute="${escapeHtml(key)}" ${disabledAttr}>
      </label>`
  ).join("");

  return renderCollapsibleSection(
    "Characteristics",
    `<div class="form-grid">${attributeInputs}</div>`,
    false,
  );
}

function renderSkillsBlock(data, disabledAttr) {
  const skillRows = Object.entries(data.odyssey.skills)
    .map(
      ([key, value]) => `
        <div class="skill-row">
          <div class="skill-name">${escapeHtml(key)}</div>
          <input type="number" min="0" max="10" value="${value}" data-action="set-odyssey-skill" data-skill="${escapeHtml(key)}" ${disabledAttr}>
          <button type="button" class="danger" data-action="remove-skill" data-skill="${escapeHtml(key)}" ${
            CORE_COMBAT_SKILLS.includes(key) ? "disabled" : disabledAttr
          }>${CORE_COMBAT_SKILLS.includes(key) ? "Core" : "Remove"}</button>
        </div>`
    )
    .join("");

  return renderCollapsibleSection(
    "Навыки",
    `
      <div class="list">${skillRows || '<div class="empty">No skills yet.</div>'}</div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Название навыка</span>
          <input type="text" data-skill-field="new-name" placeholder="New skill" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Значение</span>
          <input type="number" min="0" max="10" value="0" data-skill-field="new-value" ${disabledAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" class="secondary" data-action="add-skill" ${disabledAttr}>Add Skill</button>
      </div>
    `,
    false,
  );
}

function renderCombatBlock(token, data, tokenLocked) {
  const targetCharacters = getCharacters().filter(
    (item) => item.id !== token.id && item.visible !== false,
  );
  const disabledAttr = tokenLocked || !targetCharacters.length ? "disabled" : "";
  const draft = getAttackDraft(token, data, targetCharacters);
  const skillOptions = CORE_COMBAT_SKILLS.map(
    (key) =>
      `<option value="${escapeHtml(key)}" ${
        draft.skill === key ? "selected" : ""
      }>${escapeHtml(key)} (${data.odyssey.skills[key] ?? 0})</option>`
  ).join("");

  return renderCollapsibleSection(
    "Атака",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Attack skill</span>
          <select data-attack-field="skill" ${disabledAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Target token</span>
          <select data-attack-field="targetTokenId" ${disabledAttr}>
            ${targetCharacters
              .map(
                (target) =>
                  `<option value="${target.id}" ${target.id === draft.targetTokenId ? "selected" : ""}>${escapeHtml(
                    getCharacterName(target)
                  )}</option>`
              )
              .join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Target body part</span>
          <select data-attack-field="targetPart" ${disabledAttr}>
            ${BODY_ORDER.map(
              (part) =>
                `<option value="${part}" ${part === draft.targetPart ? "selected" : ""}>${part}</option>`
            ).join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon damage</span>
          <input type="number" value="${draft.weaponDamage}" data-attack-field="weaponDamage" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack bonuses</span>
          <input type="number" value="${draft.attackBonuses}" data-attack-field="attackBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack penalties</span>
          <input type="number" value="${draft.attackPenalties}" data-attack-field="attackPenalties" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense bonuses</span>
          <input type="number" value="${draft.defenseBonuses}" data-attack-field="defenseBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense penalties</span>
          <input type="number" value="${draft.defensePenalties}" data-attack-field="defensePenalties" ${disabledAttr}>
        </label>
      </div>
      <div class="muted">${
        targetCharacters.length
          ? "Attack goes from the selected attacker token to the chosen target token."
          : "Add at least two character tokens to perform an attack."
      }</div>
      <div class="muted">For Hand/Cold attacks, Strength above 10 adds bonus weapon damage automatically.</div>
      <div class="row row-gap">
        <button type="button" class="success" data-action="perform-attack" ${disabledAttr}>Attack</button>
      </div>
    `,
    true,
  );
}

function renderDiceBlock(token, data, tokenLocked) {
  const attributeOptions = ATTRIBUTE_UI_FIELDS
    .filter(([key]) => key !== "Parry")
    .map(
      ([key, label]) =>
        `<option value="${escapeHtml(key)}">${escapeHtml(label)} (${data.odyssey.attributes[key] ?? 0})</option>`
    )
    .join("");
  const skillOptions = Object.entries(data.odyssey.skills)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([key, value]) => `<option value="${escapeHtml(key)}">${escapeHtml(key)} (${value})</option>`)
    .join("");
  const tokenLockedAttr = tokenLocked ? "disabled" : "";

  return renderCollapsibleSection(
    "Dice",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Dice sides</span>
          <input type="number" min="2" max="1000" value="20" data-roll-field="dice">
        </label>
        <label class="field-stack">
          <span class="field-label">Modifier</span>
          <input type="number" value="0" data-roll-field="modifier">
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-dice">Roll Dice</button>
      </div>

      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Roll_Char</span>
          <select data-roll-char-field="attribute" ${tokenLockedAttr}>${attributeOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Bonus / Penalty</span>
          <input type="number" value="0" data-roll-char-field="modifier" ${tokenLockedAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-char" ${tokenLockedAttr}>Roll Char</button>
      </div>

      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Roll_Skill</span>
          <select data-roll-skill-field="skill" ${tokenLockedAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Bonus / Penalty</span>
          <input type="number" value="0" data-roll-skill-field="modifier" ${tokenLockedAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-skill" ${tokenLockedAttr}>Roll Skill</button>
      </div>
    `,
    false,
  );
}

function renderOdysseySkillRows(odyssey, skillEntries, disabledAttr) {
  return skillEntries
    .map(
      ([key, value]) => `
        <div class="skill-row">
          <div class="skill-name">${escapeHtml(key)}</div>
          <input type="number" min="0" max="10" value="${value}" data-action="set-odyssey-skill" data-skill="${escapeHtml(key)}" ${disabledAttr}>
          <label class="skill-toggle">
            <input type="checkbox" data-action="set-skill-strength-bonus" data-skill="${escapeHtml(key)}" ${disabledAttr} ${
              getSkillStrengthBonusFlag(odyssey, key) ? "checked" : ""
            }>
            <span>STR Bonus</span>
          </label>
          <button type="button" class="danger" data-action="remove-skill" data-skill="${escapeHtml(key)}" ${
            CORE_COMBAT_SKILLS.includes(key) ? "disabled" : disabledAttr
          }>${CORE_COMBAT_SKILLS.includes(key) ? "Core" : "Remove"}</button>
        </div>`
    )
    .join("");
}

function renderOdysseySkillsBlock(data, disabledAttr) {
  const combatSkillRows = renderOdysseySkillRows(
    data.odyssey,
    getCombatSkillEntries(data.odyssey),
    disabledAttr,
  );
  const appliedSkillRows = renderOdysseySkillRows(
    data.odyssey,
    getAppliedSkillEntries(data.odyssey),
    disabledAttr,
  );

  return renderCollapsibleSection(
    "РќР°РІС‹РєРё",
    `
      <div class="field-label">Боевые</div>
      <div class="list">${combatSkillRows || '<div class="empty">Нет боевых навыков.</div>'}</div>
      <div class="field-label">Прикладные</div>
      <div class="list">${appliedSkillRows || '<div class="empty">Нет прикладных навыков.</div>'}</div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">РќР°Р·РІР°РЅРёРµ РЅР°РІС‹РєР°</span>
          <input type="text" data-skill-field="new-name" placeholder="New skill" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Р—РЅР°С‡РµРЅРёРµ</span>
          <input type="number" min="0" max="10" value="0" data-skill-field="new-value" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Категория</span>
          <select data-skill-field="new-category" ${disabledAttr}>
            <option value="${COMBAT_SKILL_CATEGORY}">Боевой</option>
            <option value="${APPLIED_SKILL_CATEGORY}" selected>Прикладной</option>
          </select>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" class="secondary" data-action="add-skill" ${disabledAttr}>Add Skill</button>
      </div>
    `,
    false,
  );
}

function renderOdysseyAttackBlock(token, data, tokenLocked) {
  const targetCharacters = getCharacters().filter(
    (item) => item.id !== token.id && item.visible !== false,
  );
  const disabledAttr = tokenLocked || !targetCharacters.length ? "disabled" : "";
  const draft = getAttackDraft(token, data, targetCharacters);
  const combatSkillEntries = getCombatSkillEntries(data.odyssey);
  const skillOptions = buildSkillOptions(combatSkillEntries, draft.skill);

  return renderCollapsibleSection(
    "РђС‚Р°РєР°",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Attack skill</span>
          <select data-attack-field="skill" ${disabledAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Target token</span>
          <select data-attack-field="targetTokenId" ${disabledAttr}>
            ${targetCharacters
              .map(
                (target) =>
                  `<option value="${target.id}" ${target.id === draft.targetTokenId ? "selected" : ""}>${escapeHtml(
                    getCharacterName(target)
                  )}</option>`
              )
              .join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Target body part</span>
          <select data-attack-field="targetPart" ${disabledAttr}>
            ${BODY_ORDER.map(
              (part) =>
                `<option value="${part}" ${part === draft.targetPart ? "selected" : ""}>${part}</option>`
            ).join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon damage</span>
          <input type="number" value="${draft.weaponDamage}" data-attack-field="weaponDamage" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack bonuses</span>
          <input type="number" value="${draft.attackBonuses}" data-attack-field="attackBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack penalties</span>
          <input type="number" value="${draft.attackPenalties}" data-attack-field="attackPenalties" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense bonuses</span>
          <input type="number" value="${draft.defenseBonuses}" data-attack-field="defenseBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense penalties</span>
          <input type="number" value="${draft.defensePenalties}" data-attack-field="defensePenalties" ${disabledAttr}>
        </label>
      </div>
      <div class="muted">${
        targetCharacters.length
          ? "Attack goes from the selected attacker token to the chosen target token."
          : "Add at least two character tokens to perform an attack."
      }</div>
      <div class="muted">Для навыка "${escapeHtml(MELEE_SKILL_NAME)}" сила выше 10 автоматически добавляется к урону оружия.</div>
      <div class="row row-gap">
        <button type="button" class="success" data-action="perform-attack" ${disabledAttr}>Attack</button>
      </div>
    `,
    true,
  );
}

function renderOdysseyDiceBlock(token, data, tokenLocked) {
  const attributeOptions = ATTRIBUTE_UI_FIELDS
    .filter(([key]) => key !== "Parry")
    .map(
      ([key, label]) =>
        `<option value="${escapeHtml(key)}">${escapeHtml(label)} (${data.odyssey.attributes[key] ?? 0})</option>`
    )
    .join("");
  const skillOptions = buildGroupedSkillOptions(data.odyssey);
  const tokenLockedAttr = tokenLocked ? "disabled" : "";

  return renderCollapsibleSection(
    "Dice",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Dice sides</span>
          <input type="number" min="2" max="1000" value="20" data-roll-field="dice">
        </label>
        <label class="field-stack">
          <span class="field-label">Modifier</span>
          <input type="number" value="0" data-roll-field="modifier">
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-dice">Roll Dice</button>
      </div>

      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Характеристика</span>
          <select data-roll-char-field="attribute" ${tokenLockedAttr}>${attributeOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Bonus / Penalty</span>
          <input type="number" value="0" data-roll-char-field="modifier" ${tokenLockedAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-char" ${tokenLockedAttr}>Характеристика</button>
      </div>

      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Навык</span>
          <select data-roll-skill-field="skill" ${tokenLockedAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Bonus / Penalty</span>
          <input type="number" value="0" data-roll-skill-field="modifier" ${tokenLockedAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-skill" ${tokenLockedAttr}>Навык</button>
      </div>
    `,
    false,
  );
}

function renderEnglishSkillsBlock(data, disabledAttr) {
  const combatSkillRows = renderOdysseySkillRows(
    data.odyssey,
    getCombatSkillEntries(data.odyssey),
    disabledAttr,
  );
  const appliedSkillRows = renderOdysseySkillRows(
    data.odyssey,
    getAppliedSkillEntries(data.odyssey),
    disabledAttr,
  );

  return renderCollapsibleSection(
    "Skills",
    `
      <div class="field-label">Combat</div>
      <div class="list">${combatSkillRows || '<div class="empty">No combat skills yet.</div>'}</div>
      <div class="field-label">Applied</div>
      <div class="list">${appliedSkillRows || '<div class="empty">No applied skills yet.</div>'}</div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Skill Name</span>
          <input type="text" data-skill-field="new-name" placeholder="New skill" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Value</span>
          <input type="number" min="0" max="10" value="0" data-skill-field="new-value" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Category</span>
          <select data-skill-field="new-category" ${disabledAttr}>
            <option value="${COMBAT_SKILL_CATEGORY}">Combat</option>
            <option value="${APPLIED_SKILL_CATEGORY}" selected>Applied</option>
          </select>
        </label>
        <label class="field-stack checkbox-stack">
          <span class="field-label">Add Strength Bonus?</span>
          <label class="skill-toggle">
            <input type="checkbox" data-skill-field="new-strength-bonus" ${disabledAttr}>
            <span>Enable for attack damage</span>
          </label>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" class="secondary" data-action="add-skill" ${disabledAttr}>Add Skill</button>
      </div>
    `,
    false,
  );
}

function renderEnglishAttackBlock(token, data, tokenLocked) {
  const targetCharacters = getCharacters().filter(
    (item) => item.id !== token.id && item.visible !== false,
  );
  const disabledAttr = tokenLocked || !targetCharacters.length ? "disabled" : "";
  const draft = getAttackDraft(token, data, targetCharacters);
  const skillOptions = buildSkillOptions(getAttackSkillEntries(data.odyssey), draft.skill);
  const selectedTarget = targetCharacters.find((target) => target.id === draft.targetTokenId) ?? null;
  const targetName = selectedTarget ? getCharacterName(selectedTarget) : "No target";
  const isPickingTarget = targetPickState.active && targetPickState.attackerTokenId === token.id;

  return renderCollapsibleSection(
    "Attack",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Attack Skill</span>
          <select data-attack-field="skill" ${disabledAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Target Token</span>
          <select data-attack-field="targetTokenId" ${disabledAttr}>
            ${targetCharacters
              .map(
                (target) =>
                  `<option value="${target.id}" ${target.id === draft.targetTokenId ? "selected" : ""}>${escapeHtml(
                    getCharacterName(target)
                  )}</option>`
              )
              .join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Pick On Map</span>
          <button type="button" class="secondary" data-action="pick-attack-target" ${disabledAttr}>
            ${isPickingTarget ? "Cancel Target Pick" : "Pick Target On Map"}
          </button>
        </label>
        <label class="field-stack">
          <span class="field-label">Target Body Part</span>
          <select data-attack-field="targetPart" ${disabledAttr}>
            ${BODY_ORDER.map(
              (part) =>
                `<option value="${part}" ${part === draft.targetPart ? "selected" : ""}>${part}</option>`
            ).join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon Damage</span>
          <input type="number" value="${draft.weaponDamage}" data-attack-field="weaponDamage" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack Bonus</span>
          <input type="number" value="${draft.attackBonuses}" data-attack-field="attackBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack Penalty</span>
          <input type="number" value="${draft.attackPenalties}" data-attack-field="attackPenalties" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense Bonus</span>
          <input type="number" value="${draft.defenseBonuses}" data-attack-field="defenseBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense Penalty</span>
          <input type="number" value="${draft.defensePenalties}" data-attack-field="defensePenalties" ${disabledAttr}>
        </label>
      </div>
      <div class="muted">${
        targetCharacters.length
          ? "Attack goes from the selected attacker token to the selected target token."
          : "Add at least two visible character tokens to perform an attack."
      }</div>
      <div class="muted">Current target: ${escapeHtml(targetName)}</div>
      <div class="muted">Automatic called-shot penalties: Head -30, arms/legs -15.</div>
      <div class="muted">Strength is added to weapon damage only for attack skills with STR Bonus enabled. ${escapeHtml(PARRY_SKILL_NAME)} is added to defense.</div>
      <div class="row row-gap">
        <button type="button" class="success" data-action="perform-attack" ${disabledAttr}>Attack</button>
      </div>
    `,
    true,
  );
}

function renderEnglishDiceBlock(token, data, tokenLocked) {
  const attributeOptions = ATTRIBUTE_UI_FIELDS
    .map(
      ([key, label]) =>
        `<option value="${escapeHtml(key)}">${escapeHtml(label)} (${data.odyssey.attributes[key] ?? 0})</option>`
    )
    .join("");
  const skillOptions = buildGroupedSkillOptions(data.odyssey);
  const tokenLockedAttr = tokenLocked ? "disabled" : "";

  return renderCollapsibleSection(
    "Dice",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Dice Sides</span>
          <input type="number" min="2" max="1000" value="20" data-roll-field="dice">
        </label>
        <label class="field-stack">
          <span class="field-label">Dice Count</span>
          <input type="number" min="1" max="100" value="1" data-roll-field="count">
        </label>
        <label class="field-stack">
          <span class="field-label">Modifier</span>
          <input type="number" value="0" data-roll-field="modifier">
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-dice">Roll Dice</button>
      </div>

      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Characteristic</span>
          <select data-roll-char-field="attribute" ${tokenLockedAttr}>${attributeOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Bonus / Penalty</span>
          <input type="number" value="0" data-roll-char-field="modifier" ${tokenLockedAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-char" ${tokenLockedAttr}>Roll Characteristic</button>
      </div>

      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Skill</span>
          <select data-roll-skill-field="skill" ${tokenLockedAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Bonus / Penalty</span>
          <input type="number" value="0" data-roll-skill-field="modifier" ${tokenLockedAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-skill" ${tokenLockedAttr}>Roll Skill</button>
      </div>
    `,
    false,
  );
}

function renderPrivateGmDiceBlock() {
  if (!isEditable()) return "";

  const privateLog = gmPrivateEntries.length
    ? gmPrivateEntries
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
        .join("")
    : '<div class="empty">Private GM rolls will stay visible only here.</div>';

  return renderCollapsibleSection(
    "GM Private Dice",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Dice sides</span>
          <input type="number" min="2" max="1000" value="20" data-gm-roll-field="dice">
        </label>
        <label class="field-stack">
          <span class="field-label">Dice Count</span>
          <input type="number" min="1" max="100" value="1" data-gm-roll-field="count">
        </label>
        <label class="field-stack">
          <span class="field-label">Modifier</span>
          <input type="number" value="0" data-gm-roll-field="modifier">
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-gm-private-roll">GM Roll</button>
      </div>
      <div class="list private-roll-log">${privateLog}</div>
    `,
    false,
  );
}

function renderSelectedToken() {
  const panelState = captureSelectedPanelState();
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
  const bodyFieldDisabled = !canEditTokenData(token) ? "disabled" : "";
  const gmOnlyDisabled = !isEditable() ? "disabled" : "";
  const showPartBlock = isEditable();
  const lastRollText = data.lastRoll
    ? escapeHtml(data.lastRoll.summary || "Last roll recorded")
    : "No rolls synced yet";

  ui.selectionHint.textContent = selected ? "Selected on map" : "Showing current focus";

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
          <span class="chip-label">Assigned Player</span>
          <span class="chip-value">${escapeHtml(odyssey.owner.playerName || odyssey.owner.playerId || "Unassigned")}</span>
        </div>
      </div>

      ${
        isEditable()
          ? `
            ${renderOwnerFields({ odyssey }, gmOnlyDisabled)}
            ${renderCharacteristicsBlock({ odyssey }, gmOnlyDisabled)}
            ${renderEnglishSkillsBlock({ odyssey }, gmOnlyDisabled)}
          `
          : ""
      }
      ${renderEnglishAttackBlock(token, { odyssey }, tokenLocked)}
      ${renderEnglishDiceBlock(token, { odyssey }, tokenLocked)}
      ${renderPrivateGmDiceBlock()}
      ${renderCollapsibleSection(
        "Last Roll",
        `<pre class="console-output">${lastRollText}</pre>`,
        false,
      )}
      ${
        showPartBlock
          ? renderCollapsibleSection(
              "Body Parts",
              `
                <div class="body-table-wrap">
                  <table class="body-table">
                    <thead>
                      <tr>
                        <th>Body Part</th>
                        <th>Crit</th>
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
                                )}" data-field="current" data-delta="-1" ${bodyFieldDisabled}>-</button>
                                <input type="number" min="0" max="${part.max}" value="${part.current}" data-action="set-field" data-part="${escapeHtml(
                                  partName
                                )}" data-field="current" ${bodyFieldDisabled}>
                                <button type="button" data-action="change-part" data-part="${escapeHtml(
                                  partName
                                )}" data-field="current" data-delta="1" ${bodyFieldDisabled}>+</button>
                              </div>
                            </td>
                            <td>
                              <input class="compact-input" type="number" min="0" max="99" value="${part.max}" data-action="set-field" data-part="${escapeHtml(
                                partName
                              )}" data-field="max" ${bodyFieldDisabled}>
                            </td>
                            <td>
                              <input class="compact-input" type="number" min="0" max="99" value="${part.armor}" data-action="set-field" data-part="${escapeHtml(
                                partName
                              )}" data-field="armor" ${bodyFieldDisabled}>
                            </td>
                          </tr>
                        `;
                      }).join("")}
                    </tbody>
                  </table>
                </div>
              `,
              true,
            )
          : ""
      }
      ${renderCollapsibleSection(
        "Overlay Preview",
        `<pre class="console-output">${escapeHtml(formatOverlayText(data))}</pre>`,
        false,
      )}
    </div>`;

  restoreSelectedPanelState(panelState);
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
  const [role, id, name, color, items, selection, players] = await Promise.all([
    OBR.player.getRole(),
    OBR.player.getId(),
    OBR.player.getName(),
    OBR.player.getColor(),
    OBR.scene.items.getItems(),
    OBR.player.getSelection(),
    OBR.party.getPlayers(),
  ]);

  playerRole = role;
  playerId = id;
  playerName = name;
  playerColor = color;
  partyPlayers = players ?? [];
  sceneItems = items;
  selectionIds = selection ?? [];

  const selectedCharacterId = selectionIds.find((selectionId) =>
    sceneItems.some((item) => item.id === selectionId && isCharacterToken(item))
  );
  if (selectedCharacterId) {
    activeTokenId = selectedCharacterId;
    const initialized = await initializeCharacterToken(selectedCharacterId);
    if (initialized) {
      sceneItems = await OBR.scene.items.getItems();
    }
  } else if (activeTokenId && !sceneItems.some((item) => item.id === activeTokenId)) {
    activeTokenId = null;
  }

  render();
  await syncTargetHighlight();

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
  await syncTargetHighlight();
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

async function setOwnerPlayer(ownerPlayerId) {
  if (!isEditable()) {
    setStatus("Only the GM can assign token owners.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }

  const selectedPlayer = getSortedPartyPlayers().find((player) => player.id === ownerPlayerId);

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey ??= structuredClone(getTrackerData(token).odyssey);
    next.odyssey.owner ??= { playerId: "", playerName: "" };
    next.odyssey.owner.playerId = selectedPlayer?.id ?? "";
    next.odyssey.owner.playerName = selectedPlayer?.name ?? "";
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
  if (!isEditable()) {
    setStatus("Only the GM can edit Odyssey skills.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.skills[skill] = clamp(Number(value) || 0, 0, 10);
    return next;
  });
  await syncState();
}

async function setOdysseySkillStrengthBonus(skill, enabled) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit Odyssey skills.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.skillStrengthBonuses ??= {};
    next.odyssey.skillStrengthBonuses[skill] = Boolean(enabled);
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
  if (!isEditable()) {
    setStatus("Only the GM can edit characteristics.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.attributes[attribute] = clamp(Number(value) || 0, 0, 15);
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

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);

    if (draft.action === "set-odyssey-skill") {
      if (!isEditable()) return next;
      next.odyssey.skills[draft.skill] = clamp(Number(draft.value) || 0, 0, 10);
      return next;
    }

    if (draft.action === "set-odyssey-attribute") {
      if (!isEditable()) return next;
      next.odyssey.attributes[draft.attribute] = clamp(Number(draft.value) || 0, 0, 15);
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
      if (!canEditTokenData(token)) return next;
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
  if (target.visible === false) {
    setStatus("Hidden tokens cannot be targeted.", "error");
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
  const manualAttackPenalties = Number(getActionFieldValue('[data-attack-field="attackPenalties"]')) || 0;
  const automaticTargetPenalty = getAutomaticTargetPenalty(targetPart);
  const totalAttackPenalties = manualAttackPenalties + automaticTargetPenalty;
  const defenseBonuses = Number(getActionFieldValue('[data-attack-field="defenseBonuses"]')) || 0;
  const defensePenalties = Number(getActionFieldValue('[data-attack-field="defensePenalties"]')) || 0;
  saveAttackDraftValue(attacker.id, "skill", skillName);
  saveAttackDraftValue(attacker.id, "targetTokenId", targetTokenId);
  saveAttackDraftValue(attacker.id, "targetPart", targetPart);
  saveAttackDraftValue(attacker.id, "weaponDamage", getActionFieldValue('[data-attack-field="weaponDamage"]'));
  saveAttackDraftValue(attacker.id, "attackBonuses", getActionFieldValue('[data-attack-field="attackBonuses"]'));
  saveAttackDraftValue(attacker.id, "attackPenalties", getActionFieldValue('[data-attack-field="attackPenalties"]'));
  saveAttackDraftValue(attacker.id, "defenseBonuses", getActionFieldValue('[data-attack-field="defenseBonuses"]'));
  saveAttackDraftValue(attacker.id, "defensePenalties", getActionFieldValue('[data-attack-field="defensePenalties"]'));
  const targetArmor = targetData.body[targetPart]?.armor ?? 0;
  const targetPartState = targetData.body[targetPart] ?? { current: 0, max: 0, armor: 0, minor: 0, serious: 0 };
  const beforeHp = targetPartState.current ?? 0;
  const beforeMinor = targetPartState.minor ?? 0;
  const beforeSerious = targetPartState.serious ?? 0;
  const strengthBonus = getSkillStrengthBonusFlag(attackerOdyssey, skillName)
    ? Math.max((attackerOdyssey.attributes.Strength ?? 0) - 10, 0)
    : 0;
  const finalWeaponDamage = weaponDamage + strengthBonus;
  const targetParry = skillName === MELEE_SKILL_NAME
    ? targetOdyssey.skills[PARRY_SKILL_NAME] ?? 0
    : 0;

  const result = resolveAttack({
    attackSkill: attackerOdyssey.skills[skillName] ?? 0,
    weaponDamage: finalWeaponDamage,
    defenseBonuses,
    defensePenalties,
    attackBonuses,
    attackPenalties: totalAttackPenalties,
    parry: targetParry,
    targetPart,
    targetArmor,
  });
  const projectedPartState =
    result.hit && result.damage
      ? projectPartDamage(targetPartState, result.damage)
      : {
          ...targetPartState,
          critApplied: 0,
        };
  const afterHp = projectedPartState.current ?? beforeHp;
  const afterMinor = projectedPartState.minor ?? beforeMinor;
  const afterSerious = projectedPartState.serious ?? beforeSerious;

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
      next.body[result.targetPart].current = projectedPartState.current;
      next.body[result.targetPart].minor = projectedPartState.minor;
      next.body[result.targetPart].serious = projectedPartState.serious;
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
      weaponDamage: finalWeaponDamage,
      strengthBonus,
      attackBonuses,
      manualAttackPenalties,
      automaticTargetPenalty,
      totalAttackPenalties,
      defenseBonuses,
      defensePenalties,
      targetParry,
      targetArmor,
      result,
      beforeHp,
      afterHp,
      beforeMinor,
      afterMinor,
      beforeSerious,
      afterSerious,
      critApplied: projectedPartState.critApplied ?? 0,
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
  const count = Number(getActionFieldValue('[data-roll-field="count"]')) || 1;
  const modifier = Number(getActionFieldValue('[data-roll-field="modifier"]')) || 0;
  const result = rollDice(dice, modifier, count);
  const diceLabel = `${result.count}d${result.sides}`;

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.lastRoll = {
      eventId: 0,
      actorName: playerName || "Owlbear Player",
      summary: `Rolled ${diceLabel}: ${result.rolls.join(", ")}${modifier ? ` ${modifier >= 0 ? "+" : ""}${modifier}` : ""} = ${result.total}`,
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
  setStatus(`${diceLabel} rolled ${result.total}.`, "success");
}

async function addOdysseySkill() {
  if (!isEditable()) {
    setStatus("Only the GM can add skills.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }

  const name = getActionFieldValue('[data-skill-field="new-name"]').trim();
  const value = clamp(Number(getActionFieldValue('[data-skill-field="new-value"]')) || 0, 0, 10);
  const strengthBonusField = ui.selectedTokenPanel.querySelector(
    '[data-skill-field="new-strength-bonus"]',
  );
  const addStrengthBonus =
    strengthBonusField instanceof HTMLInputElement ? strengthBonusField.checked : false;
  if (!name) {
    setStatus("Enter a skill name first.", "error");
    return;
  }
  const category =
    name === MELEE_SKILL_NAME || name === PARRY_SKILL_NAME
      ? COMBAT_SKILL_CATEGORY
      : getActionFieldValue('[data-skill-field="new-category"]') === COMBAT_SKILL_CATEGORY
        ? COMBAT_SKILL_CATEGORY
        : APPLIED_SKILL_CATEGORY;

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.skills[name] = value;
    next.odyssey.skillCategories ??= {};
    next.odyssey.skillStrengthBonuses ??= {};
    next.odyssey.skillCategories[name] = category;
    next.odyssey.skillStrengthBonuses[name] =
      name === MELEE_SKILL_NAME
        ? true
        : name === PARRY_SKILL_NAME
          ? false
          : Boolean(addStrengthBonus);
    return next;
  });
  await syncState();
  setStatus(`Skill "${name}" saved for ${getCharacterName(token)}.`, "success");
}

async function removeOdysseySkill(skillName) {
  if (!isEditable()) {
    setStatus("Only the GM can remove skills.", "error");
    return;
  }
  if (CORE_COMBAT_SKILLS.includes(skillName)) {
    setStatus("Core combat skills cannot be removed.", "error");
    return;
  }

  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    delete next.odyssey.skills[skillName];
    delete next.odyssey.skillCategories?.[skillName];
    delete next.odyssey.skillStrengthBonuses?.[skillName];
    return next;
  });
  await syncState();
  setStatus(`Skill "${skillName}" removed.`, "success");
}

async function performRollChar() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot roll for this token.", "error");
    return;
  }

  const attribute = getActionFieldValue('[data-roll-char-field="attribute"]') || "Strength";
  const modifier = Number(getActionFieldValue('[data-roll-char-field="modifier"]')) || 0;
  const odyssey = getOdysseyData(token);
  const result = rollCharacterCheck(odyssey.attributes[attribute] ?? 0, modifier);
  const attributeLabel =
    ATTRIBUTE_UI_FIELDS.find(([key]) => key === attribute)?.[1] ?? attribute;
  const summary = `Characteristic ${attributeLabel}: ${result.roll} vs ${result.finalAttribute} (${result.result})`;

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.lastRoll = {
      eventId: 0,
      actorName: getCharacterName(token),
      summary,
      outcome: "roll-char",
      total: result.roll,
      targetPart: "",
      timestamp: new Date().toISOString(),
      source: "owlbear-extension",
    };
    next.history = [next.lastRoll, ...(next.history ?? [])].slice(0, 12);
    return next;
  });

  await pushDebugEntry(
    `${getCharacterName(token)} rolls characteristic`,
    formatRollCharDebug({
      tokenName: getCharacterName(token),
      attributeLabel,
      result,
    }),
    result.result.includes("Failed") ? "info" : "success",
  );
  await syncState();
  setStatus(summary, result.result.includes("Failed") ? "error" : "success");
}

async function performRollSkill() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot roll for this token.", "error");
    return;
  }

  const skillName = getActionFieldValue('[data-roll-skill-field="skill"]');
  if (!skillName) {
    setStatus("Choose a skill first.", "error");
    return;
  }

  const modifier = Number(getActionFieldValue('[data-roll-skill-field="modifier"]')) || 0;
  const odyssey = getOdysseyData(token);
  const result = rollSkillCheck(odyssey.skills[skillName] ?? 0, modifier);
  const summary = `Skill ${skillName}: ${result.totalPrimary} vs ${result.totalSecondary} (${result.result})`;

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.lastRoll = {
      eventId: 0,
      actorName: getCharacterName(token),
      summary,
      outcome: "roll-skill",
      total: result.totalPrimary,
      targetPart: "",
      timestamp: new Date().toISOString(),
      source: "owlbear-extension",
    };
    next.history = [next.lastRoll, ...(next.history ?? [])].slice(0, 12);
    return next;
  });

  await pushDebugEntry(
    `${getCharacterName(token)} checks skill`,
    formatRollSkillDebug({
      tokenName: getCharacterName(token),
      skillName,
      result,
    }),
    result.result === "Check Passed" ? "success" : "info",
  );
  await syncState();
  setStatus(summary, result.result === "Check Passed" ? "success" : "error");
}

async function performPrivateGmRoll() {
  if (!isEditable()) {
    setStatus("Only the GM can use private rolls.", "error");
    return;
  }

  const dice = Number(getActionFieldValue('[data-gm-roll-field="dice"]')) || 20;
  const count = Number(getActionFieldValue('[data-gm-roll-field="count"]')) || 1;
  const modifier = Number(getActionFieldValue('[data-gm-roll-field="modifier"]')) || 0;
  const result = rollDice(dice, modifier, count);
  const diceLabel = `${result.count}d${result.sides}`;

  pushPrivateGmEntry(
    `GM private ${diceLabel}`,
    formatDiceDebug({
      tokenName: "GM private roll",
      result,
    }),
  );
  render();
  setStatus(`Private GM roll: ${diceLabel} = ${result.total}.`, "success");
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

  document.addEventListener(
    "toggle",
    (event) => {
      const target = event.target;
      if (!(target instanceof HTMLDetailsElement)) return;
      if (!target.dataset.sectionKey) return;
      collapsibleSectionState.set(target.dataset.sectionKey, target.open);
    },
    true,
  );

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
    const skill = actionNode.dataset.skill;

    if (action === "select-character" && tokenId) {
      void selectCharacter(tokenId).catch((error) => {
        setStatus(error?.message ?? "Unable to select token.", "error");
      });
    }

    if (action === "focus-token" && activeTokenId) {
      void selectCharacter(activeTokenId).catch((error) => {
        setStatus(error?.message ?? "Unable to focus token.", "error");
      });
    }

    if (action === "pick-attack-target") {
      void startTargetPick().catch((error) => {
        setStatus(error?.message ?? "Unable to start target picking.", "error");
      });
      return;
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
      return;
    }

    if (action === "perform-roll-char") {
      void performRollChar().catch((error) => {
        setStatus(error?.message ?? "Unable to resolve Roll_Char.", "error");
      });
      return;
    }

    if (action === "perform-roll-skill") {
      void performRollSkill().catch((error) => {
        setStatus(error?.message ?? "Unable to resolve Roll_Skill.", "error");
      });
      return;
    }

    if (action === "perform-gm-private-roll") {
      void performPrivateGmRoll().catch((error) => {
        setStatus(error?.message ?? "Unable to perform private GM roll.", "error");
      });
      return;
    }

    if (action === "add-skill") {
      void addOdysseySkill().catch((error) => {
        setStatus(error?.message ?? "Unable to add skill.", "error");
      });
      return;
    }

    if (action === "remove-skill" && skill) {
      void removeOdysseySkill(skill).catch((error) => {
        setStatus(error?.message ?? "Unable to remove skill.", "error");
      });
    }
  });

  document.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement)) return;

    if (target.dataset.attackField && activeTokenId) {
      saveAttackDraftValue(activeTokenId, target.dataset.attackField, target.value);
      if (target.dataset.attackField === "targetTokenId") {
        void syncTargetHighlight().catch((error) => {
          console.warn("[Body HP] Unable to sync target highlight", error);
        });
      }
    }

    if (target.dataset.action === "select-owner-player") {
      void setOwnerPlayer(target.value).catch((error) => {
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

    if (target.dataset.action === "set-skill-strength-bonus") {
      const skill = target.dataset.skill;
      if (!skill || !(target instanceof HTMLInputElement)) return;
      void setOdysseySkillStrengthBonus(skill, target.checked).catch((error) => {
        setStatus(error?.message ?? "Unable to save strength bonus flag.", "error");
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

    if (target.dataset.attackField) {
      saveAttackDraftValue(activeTokenId, target.dataset.attackField, target.value);
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
      void syncTargetHighlight().catch((error) => {
        console.warn("[Body HP] Unable to sync target highlight", error);
      });
    });

    OBR.player.onChange((player) => {
      playerRole = player.role;
      playerId = player.id ?? playerId;
      playerName = player.name ?? playerName;
      playerColor = player.color ?? playerColor;
      selectionIds = player.selection ?? [];
      const selectedCharacterId = selectionIds.find((selectionId) =>
        sceneItems.some((item) => item.id === selectionId && isCharacterToken(item))
      );
      if (selectedCharacterId) {
        activeTokenId = selectedCharacterId;
      }

      void syncState().catch((error) => {
        console.warn("[Body HP] Player state sync failed", error);
        render();
      });
    });

    OBR.party.onChange((players) => {
      partyPlayers = players ?? [];
      render();
    });

    OBR.broadcast.onMessage(DEBUG_BROADCAST_CHANNEL, (event) => {
      const payload = event?.data;
      if (!payload || typeof payload !== "object") return;
      if (payload.type !== "debug-entry") return;

      const nextEntries = mergeDebugEntries([payload.entry], debugEntries);
      debugEntries = nextEntries;
      renderDebugConsole();

      if (playerRole === "GM") {
        void OBR.room.setMetadata({
          [DEBUG_LOG_KEY]: nextEntries,
        }).catch((error) => {
          console.warn("[Body HP] Unable to persist broadcast debug entry", error);
        });
      }
    });

    OBR.room.onMetadataChange((metadata) => {
      debugEntries = mergeDebugEntries(metadata?.[DEBUG_LOG_KEY], debugEntries);
      renderDebugConsole();
    });
  } catch (error) {
    setStatus(error?.message ?? "Extension failed to initialize.", "error");
  }
});
