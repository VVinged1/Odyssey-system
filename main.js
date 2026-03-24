import { Command, buildImage, buildPath } from "@owlbear-rodeo/sdk";
import {
  ABILITIES_SKILL_CATEGORY,
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
  getArmorTotal,
  getAvailableWeapons,
  getBodyTotals,
  getCharacterName,
  getOdysseyData,
  getTargetableBodyParts,
  getTrackerData,
  hasConfiguredSpecial,
  SHIELD_PART_NAME,
  SPECIAL_PART_NAME,
  isCharacterToken,
  isTrackedCharacter,
  removeOverlaysForToken,
  sortCharacters,
  syncTrackedOverlays,
  updateTrackerData,
} from "./shared.js";
import {
  formatAttackOutcomeLabel,
  getAttackOutcomeIcon,
  resolveAttack,
  rollDice,
} from "./odyssey_rules.js";

const DEBUG_LOG_KEY = "com.codex.body-hp/debugLog";
const DEBUG_BROADCAST_CHANNEL = "com.codex.body-hp/debug";
const DEBUG_ENTRY_LIMIT = 50;
const TARGET_PICK_TOOL_ID = "com.codex.body-hp/attack-target-picker";
const TARGET_PICK_MODE_ID = "pick-target";
const TARGET_HIGHLIGHT_KEY = "com.codex.body-hp/local-attack-target";
const LOCAL_SELF_VIEW_KEY = "com.codex.body-hp/local-self-view";
const ATTACK_TARGET_CONTEXT_MENU_ID = "com.codex.body-hp/set-attack-target";
const SHOW_EMBEDDED_PUBLIC_LOG = false;
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
const UNARMED_WEAPON_NAME = "Not Armed";
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
  clearDebugBtn: document.getElementById("clearDebugBtn"),
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
const pendingLocalDebugEntryIds = new Set();
let pendingLocalDebugClear = false;
const collapsibleSectionState = new Map();
const attackFormDrafts = new Map();
const inputAutosaveTimers = new Map();
let selectionPollTimer = null;
let overlayMaintenanceTimer = null;
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
    .slice(0, DEBUG_ENTRY_LIMIT);
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
    .slice(0, DEBUG_ENTRY_LIMIT);
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
  const category = odyssey?.skillCategories?.[skillName];
  if (category === COMBAT_SKILL_CATEGORY) return COMBAT_SKILL_CATEGORY;
  if (category === ABILITIES_SKILL_CATEGORY) return ABILITIES_SKILL_CATEGORY;
  return APPLIED_SKILL_CATEGORY;
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

function getAbilitiesSkillEntries(odyssey) {
  return getSortedSkillEntries(odyssey).filter(
    ([skillName]) => getSkillCategory(odyssey, skillName) === ABILITIES_SKILL_CATEGORY,
  );
}

function getAttackSkillEntries(odyssey) {
  return [...getCombatSkillEntries(odyssey), ...getAbilitiesSkillEntries(odyssey)].filter(
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
  const abilitiesOptions = buildSkillOptions(getAbilitiesSkillEntries(odyssey), selectedValue);
  const appliedOptions = buildSkillOptions(getAppliedSkillEntries(odyssey), selectedValue);

  return [
    combatOptions ? `<optgroup label="Combat">${combatOptions}</optgroup>` : "",
    abilitiesOptions ? `<optgroup label="Abilities">${abilitiesOptions}</optgroup>` : "",
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
  if (field.dataset.manualAttackField) return `manual-attack:${field.dataset.manualAttackField}`;
  if (field.dataset.rollField) return `roll:${field.dataset.rollField}`;
  if (field.dataset.rollCharField) return `roll-char:${field.dataset.rollCharField}`;
  if (field.dataset.rollSkillField) return `roll-skill:${field.dataset.rollSkillField}`;
  if (field.dataset.gmRollField) return `gm-roll:${field.dataset.gmRollField}`;
  if (field.dataset.skillField) return `new-skill:${field.dataset.skillField}`;
  if (field.dataset.weaponField) return `new-weapon:${field.dataset.weaponField}`;

  if (field.dataset.action === "select-owner-player") return "owner";
  if (field.dataset.action === "set-odyssey-skill") return `skill:${field.dataset.skill ?? ""}`;
  if (field.dataset.action === "set-skill-strength-bonus") {
    return `skill-strength:${field.dataset.skill ?? ""}`;
  }
  if (field.dataset.action === "set-odyssey-attribute") {
    return `attribute:${field.dataset.attribute ?? ""}`;
  }
  if (field.dataset.action === "set-weapon-name") {
    return `weapon-name:${field.dataset.weaponIndex ?? ""}`;
  }
  if (field.dataset.action === "set-weapon-damage") {
    return `weapon-damage:${field.dataset.weaponIndex ?? ""}`;
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
    fieldKey.startsWith("manual-attack:") ||
    fieldKey.startsWith("roll:") ||
    fieldKey.startsWith("roll-char:") ||
    fieldKey.startsWith("roll-skill:") ||
    fieldKey.startsWith("gm-roll:") ||
    fieldKey.startsWith("new-skill:") ||
    fieldKey.startsWith("new-weapon:")
  );
}

function captureSelectedPanelState() {
  const renderedTokenId = String(ui.selectedTokenPanel.dataset.tokenId ?? "").trim();
  if (!renderedTokenId || !ui.selectedTokenPanel.childElementCount) return null;

  let focusedKey = "";
  let selectionStart = null;
  let selectionEnd = null;
  const activeField = document.activeElement;
  if (
    (activeField instanceof HTMLInputElement || activeField instanceof HTMLSelectElement) &&
    ui.selectedTokenPanel.contains(activeField)
  ) {
    focusedKey = getTransientFieldKey(activeField);
    if (
      activeField instanceof HTMLInputElement &&
      typeof activeField.selectionStart === "number" &&
      typeof activeField.selectionEnd === "number"
    ) {
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
    tokenId: renderedTokenId,
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

  if (playerRole !== "GM") {
    const firstControllable = getControllableCharacters()[0];
    if (firstControllable) return firstControllable.id;
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

function canViewAttackBlock(token) {
  return canUseToken(token);
}

function canViewBodyHpSummary(token) {
  return canUseToken(token);
}

function canViewDiceBlock(token) {
  return canUseToken(token);
}

function canViewOverlayPreview(token) {
  return canUseToken(token);
}

async function initializeCharacterToken(tokenId) {
  const token = getCharacterById(tokenId);
  if (!token || !isCharacterToken(token)) return false;
  const shouldInitialize = !isTrackedCharacter(token);
  if (!shouldInitialize) return false;

  await updateTrackerData(tokenId, (current) => current);
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
  pendingLocalDebugEntryIds.add(entry.id);
  await OBR.broadcast.sendMessage(
    DEBUG_BROADCAST_CHANNEL,
    { type: "debug-entry", entry },
    { destination: "ALL" },
  );

  if (playerRole === "GM") {
    await OBR.room.setMetadata({
      [DEBUG_LOG_KEY]: nextEntries,
    });
  }
}

async function clearDebugConsole() {
  if (playerRole !== "GM") {
    setStatus("Only the GM can clear the debug console.", "error");
    return;
  }

  debugEntries = [];
  renderDebugConsole();
  pendingLocalDebugClear = true;
  await OBR.broadcast.sendMessage(
    DEBUG_BROADCAST_CHANNEL,
    { type: "debug-clear" },
    { destination: "ALL" },
  );
  await OBR.room.setMetadata({
    [DEBUG_LOG_KEY]: [],
  });
  setStatus("Debug console cleared.", "success");
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
  baseTargetParry,
  targetParry,
  parryMode,
  targetArmor,
  specialArmor,
  specialActive,
  result,
  beforeHp,
  afterHp,
  specialBeforeHp,
  specialAfterHp,
  beforeMinor,
  afterMinor,
  beforeSerious,
  afterSerious,
  critApplied,
  damageAppliedLabel,
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

  const damageRows = [
    ["Attacker", attackerName],
    ["Target", `${targetName} -> ${targetPart}`],
    ["Attack Skill", `${attackSkillName} (${attackSkillValue})`],
    ["Strength Bonus", strengthBonus],
    ["Manual Attack Penalty", manualAttackPenalties],
    ["Auto Target Penalty", automaticTargetPenalty],
    ["Parry Mode", getParryModeLabel(parryMode)],
    ["Base Parry", baseTargetParry],
    ["Effective Parry", targetParry],
    ["Armor", targetArmor],
  ];

  if (specialActive) {
    damageRows.push(["Special Armor", specialArmor]);
    damageRows.push(["Special HP", formatStateTransition(specialBeforeHp, specialAfterHp)]);
  }

  damageRows.push(
    ["Outcome", formatAttackOutcomeLabel(result.outcome)],
    ["Applied Damage", damageAppliedLabel ?? formatAppliedDamageLabel(result.damage, critApplied)],
    ["Damage Diff", result.damage?.damageDiff ?? 0],
    ["Damage Label", result.damage?.label ?? "No damage"],
    ["Applied Min/Sir/Crit", `${result.damage?.minor ?? 0} / ${result.damage?.serious ?? 0} / ${result.damage?.crit ?? 0}`],
    ["Converted Crit", critApplied],
  );

  const damageTable = formatTextTable(
    ["Parameter", "Value"],
    damageRows,
  );

  return [
    `Damage Applied: ${damageAppliedLabel ?? formatAppliedDamageLabel(result.damage, critApplied)}`,
    `Result: ${formatAttackOutcomeLabel(result.outcome)}`,
    "",
    accuracyTable,
    "",
    damageTable,
  ].join("\n");
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

function getNormalizedPartState(part) {
  return {
    current: Number(part?.current) || 0,
    max: Number(part?.max) || 0,
    armor: Number(part?.armor) || 0,
    minor: Number(part?.minor) || 0,
    serious: Number(part?.serious) || 0,
  };
}

function projectDamageWithSpecialProtection({
  specialPart,
  targetPart,
  damage,
  targetPartName,
}) {
  const normalizedSpecial = getNormalizedPartState(specialPart);
  const normalizedTarget = getNormalizedPartState(targetPart);

  if (!damage) {
    return {
      specialProjectedState: normalizedSpecial,
      projectedTargetState: {
        ...normalizedTarget,
        critApplied: 0,
      },
      specialActive: false,
      specialArmor: 0,
      damageAppliedLabel: "No Damage",
    };
  }

  const specialActive =
    hasConfiguredSpecial({ body: { [SPECIAL_PART_NAME]: normalizedSpecial } }) &&
    normalizedSpecial.max > 0 &&
    normalizedSpecial.current > 0;

  if (!specialActive) {
    const projectedTargetState = projectPartDamage(normalizedTarget, damage);
    return {
      specialProjectedState: normalizedSpecial,
      projectedTargetState,
      specialActive: false,
      specialArmor: 0,
      damageAppliedLabel: formatAppliedDamageLabel(damage, projectedTargetState.critApplied ?? 0),
    };
  }

  const specialProjectedState = projectPartDamage(normalizedSpecial, damage);
  const absorbedHp = Math.max(0, normalizedSpecial.current - specialProjectedState.current);
  const totalSpecialCrit = Math.max(0, Number(specialProjectedState.critApplied) || 0);
  const specialCritApplied = Math.min(totalSpecialCrit, absorbedHp);
  const overflowCrit = Math.max(0, totalSpecialCrit - specialCritApplied);
  const overflowDamage = overflowCrit > 0 ? { crit: overflowCrit, serious: 0, minor: 0 } : null;
  const projectedTargetState = overflowDamage
    ? projectPartDamage(normalizedTarget, overflowDamage)
    : {
        ...normalizedTarget,
        critApplied: 0,
      };
  const specialAppliedLabel = formatAppliedDamageLabel(
    {
      serious: damage.serious,
      minor: damage.minor,
    },
    specialCritApplied,
  );
  const targetAppliedLabel = overflowDamage
    ? formatAppliedDamageLabel(overflowDamage, projectedTargetState.critApplied ?? 0)
    : "No Damage";

  let damageAppliedLabel = "No Damage";
  if (specialAppliedLabel !== "No Damage") {
    damageAppliedLabel = `Special ${specialAppliedLabel}`;
  }
  if (targetAppliedLabel !== "No Damage") {
    damageAppliedLabel =
      damageAppliedLabel === "No Damage"
        ? `${targetPartName} ${targetAppliedLabel}`
        : `${damageAppliedLabel}; ${targetPartName} ${targetAppliedLabel}`;
  }

  return {
    specialProjectedState,
    projectedTargetState,
    specialActive: true,
    specialArmor: normalizedSpecial.armor,
    damageAppliedLabel,
  };
}

function formatRawDiceRolls(result) {
  return result.rolls.join(", ");
}

function formatDiceRollsWithModifier(result) {
  const modifier = Number(result.modifier) || 0;
  return result.rolls.map((roll) => (Number(roll) || 0) + modifier).join(", ");
}

function buildDiceRollSummary(diceLabel, result) {
  return `Rolled ${diceLabel}: raw [${formatRawDiceRolls(result)}], with modifier ${formatDiceRollsWithModifier(result)}`;
}

function formatOverlayPreviewText(data) {
  return `Armor Total ${getArmorTotal(data)}\n${formatOverlayText(data)}`;
}

function formatDiceDebug({ tokenName, result }) {
  return formatTextTable(
    ["Parameter", "Value"],
    [
      ["Actor", tokenName],
      ["Dice", `${result.count}d${result.sides}`],
      ["Raw Dice", formatRawDiceRolls(result)],
      ["With Modifier", formatDiceRollsWithModifier(result)],
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
  if (
    targetPart === "L.Arm" ||
    targetPart === "R.Arm" ||
    targetPart === "L.Leg" ||
    targetPart === "R.Leg"
  ) {
    return 15;
  }
  return 0;
}

function getParryDivisor(mode) {
  if (mode === "off") return 0;
  const numeric = Number(mode);
  if (Number.isInteger(numeric) && numeric >= 1 && numeric <= 5) {
    return numeric;
  }
  return 1;
}

function getParryModeLabel(mode) {
  if (mode === "off") return "Ignore Parry";
  const divisor = getParryDivisor(mode);
  return `${divisor} Opponent${divisor === 1 ? "" : "s"}`;
}

function formatAppliedDamageLabel(damage, critApplied = 0) {
  const parts = [];
  const totalCrit = Math.max(0, Number(critApplied) || 0);
  const serious = Math.max(0, Number(damage?.serious) || 0);
  const minor = Math.max(0, Number(damage?.minor) || 0);

  if (totalCrit > 0) {
    parts.push(`${totalCrit} Crit`);
  }
  if (serious > 0) {
    parts.push(`${serious} Serious`);
  }
  if (minor > 0) {
    parts.push(`${minor} Minor`);
  }

  return parts.length ? parts.join(", ") : "No Damage";
}

function formatStateTransition(before, after) {
  if (before == null || after == null) return "-";
  return `${before} -> ${after}`;
}

function getCheckResultIcon(resultLabel) {
  return String(resultLabel).trim() === "Check Passed" ? "✅" : "❌";
}

function formatCheckResultLabel(resultLabel) {
  const normalized = String(resultLabel).trim() || "Check Failed";
  return `${getCheckResultIcon(normalized)} ${normalized}`;
}

function getResolvedCheckResultIcon(resultLabel) {
  const normalized = String(resultLabel).trim();
  if (normalized === "Critical Success") return "🎯";
  if (normalized === "Critical Failure") return "💀";
  if (normalized === "Check Passed") return "✅";
  return "❌";
}

function isResolvedCheckResultSuccess(resultLabel) {
  const normalized = String(resultLabel).trim();
  return normalized === "Check Passed" || normalized === "Critical Success";
}

function formatResolvedCheckResultLabel(resultLabel) {
  const normalized = String(resultLabel).trim() || "Check Failed";
  return `${getResolvedCheckResultIcon(normalized)} ${normalized}`;
}

function formatRollCharDebug({ tokenName, attributeLabel, result }) {
  return [
    `Character: ${tokenName}`,
    `Characteristic: ${attributeLabel}`,
    `${formatResolvedCheckResultLabel(result.result)}`,
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
    `${formatResolvedCheckResultLabel(result.result)}`,
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
  const availableWeapons = getAttackSelectableWeapons(token);
  const defaultWeapon = getDefaultAttackWeapon(token);
  const stored = attackFormDrafts.get(token.id) ?? {};
  const persistedTargetTokenId = String(data.odyssey?.attackDraft?.targetTokenId ?? "").trim();
  const persistedTargetTokenName = String(data.odyssey?.attackDraft?.targetTokenName ?? "").trim();
  const combatSkillNames = getAttackSkillEntries(data.odyssey).map(([skillName]) => skillName);
  const fallbackSkill =
    combatSkillNames[0] ??
    CORE_COMBAT_SKILLS.find((key) => key in data.odyssey.skills) ??
    CORE_COMBAT_SKILLS[0];
  const hasStoredTargetTokenId = Object.prototype.hasOwnProperty.call(stored, "targetTokenId");
  const draftTargetTokenId = persistedTargetTokenId ||
    (hasStoredTargetTokenId ? String(stored.targetTokenId ?? "").trim() : null);
  const draftTargetTokenName =
    persistedTargetTokenName || String(stored.targetTokenName ?? "").trim();
  const targetTokenId =
    draftTargetTokenId && draftTargetTokenId !== token.id
      ? draftTargetTokenId
      : "";
  const resolvedTarget = targetTokenId
    ? targetCharacters.find((target) => target.id === targetTokenId) ?? null
    : null;
  const storedWeaponName = String(stored.weaponName ?? "").trim();
  const selectedWeapon =
    availableWeapons.find((weapon) => weapon.name === storedWeaponName) ?? defaultWeapon;

  return {
    skill: combatSkillNames.includes(stored.skill)
      ? stored.skill
      : fallbackSkill,
    targetTokenId,
    targetTokenName: resolvedTarget ? getCharacterName(resolvedTarget) : draftTargetTokenName,
    targetPart:
      BODY_ORDER.includes(stored.targetPart) && stored.targetPart !== SPECIAL_PART_NAME
        ? stored.targetPart
        : "Torso",
    weaponName: selectedWeapon?.name ?? defaultWeapon.name,
    weaponDamage: stored.weaponDamage ?? String(selectedWeapon?.damage ?? defaultWeapon.damage ?? 0),
    attackBonuses: stored.attackBonuses ?? "0",
    attackPenalties: stored.attackPenalties ?? "0",
    manualAttackBonuses: stored.manualAttackBonuses ?? stored.attackBonuses ?? "0",
    manualAttackPenalties: stored.manualAttackPenalties ?? stored.attackPenalties ?? "0",
    defenseBonuses: stored.defenseBonuses ?? "0",
    defensePenalties: stored.defensePenalties ?? "0",
    manualArmor: stored.manualArmor ?? "0",
    manualParry: stored.manualParry ?? "0",
    parryMode: ["off", "1", "2", "3", "4", "5"].includes(String(stored.parryMode))
      ? String(stored.parryMode)
      : "1",
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

function getSharedAttackDraftField(manualField) {
  if (manualField === "skill") return "skill";
  if (manualField === "weaponName") return "weaponName";
  if (manualField === "weaponDamage") return "weaponDamage";
  if (manualField === "attackBonuses") return "manualAttackBonuses";
  if (manualField === "attackPenalties") return "manualAttackPenalties";
  if (manualField === "manualArmor") return "manualArmor";
  if (manualField === "manualParry") return "manualParry";
  return "";
}

function buildWeaponOptions(weapons, selectedWeaponName = "") {
  return weapons
    .map(
      (weapon) => `<option value="${escapeHtml(weapon.name)}" ${
        weapon.name === selectedWeaponName ? "selected" : ""
      }>${escapeHtml(weapon.name)} (${weapon.damage >= 0 ? "+" : ""}${weapon.damage})</option>`,
    )
    .join("");
}

function getAttackSelectableWeapons(token) {
  const odyssey = getOdysseyData(token);
  const meleeWeapons = odyssey.weapons?.melee ?? [];
  return [{ name: UNARMED_WEAPON_NAME, damage: 0 }, ...meleeWeapons];
}

function getDefaultAttackWeapon(token) {
  const odyssey = getOdysseyData(token);
  const meleeWeapons = odyssey.weapons?.melee ?? [];
  return meleeWeapons[0] ?? { name: UNARMED_WEAPON_NAME, damage: 0 };
}

function getWeaponByName(token, weaponName) {
  const normalizedWeaponName = String(weaponName ?? "").trim();
  const weapons = getAttackSelectableWeapons(token);
  return weapons.find((weapon) => weapon.name === normalizedWeaponName) ?? getDefaultAttackWeapon(token);
}

function syncAttackWeaponInputs(tokenId, weaponName, weaponDamage) {
  if (!tokenId) return;

  saveAttackDraftValue(tokenId, "weaponName", weaponName);
  saveAttackDraftValue(tokenId, "weaponDamage", String(weaponDamage));

  const attackWeaponSelect = ui.selectedTokenPanel.querySelector('[data-attack-field="weaponName"]');
  if (attackWeaponSelect instanceof HTMLSelectElement) {
    const hasOption = Array.from(attackWeaponSelect.options).some(
      (option) => option.value === weaponName,
    );
    if (hasOption) {
      attackWeaponSelect.value = weaponName;
    }
  }

  const manualWeaponSelect = ui.selectedTokenPanel.querySelector('[data-manual-attack-field="weaponName"]');
  if (manualWeaponSelect instanceof HTMLSelectElement) {
    const hasOption = Array.from(manualWeaponSelect.options).some(
      (option) => option.value === weaponName,
    );
    if (hasOption) {
      manualWeaponSelect.value = weaponName;
    }
  }

  ui.selectedTokenPanel
    .querySelectorAll('[data-attack-field="weaponDamage"], [data-manual-attack-field="weaponDamage"]')
    .forEach((field) => {
      if (field instanceof HTMLInputElement) {
        field.value = String(weaponDamage);
      }
    });
}

async function persistAttackTargetToken(tokenId, targetTokenId) {
  if (!tokenId) return;

  const token = getCharacterById(tokenId);
  if (!token) return;

  const normalizedTargetTokenId = String(targetTokenId ?? "").trim();
  const targetToken = normalizedTargetTokenId ? getCharacterById(normalizedTargetTokenId) : null;
  const normalizedTargetTokenName = targetToken ? getCharacterName(targetToken) : "";
  const currentAttackDraft = getOdysseyData(token).attackDraft ?? {};
  const currentTargetTokenId = String(currentAttackDraft.targetTokenId ?? "").trim();
  const currentTargetTokenName = String(currentAttackDraft.targetTokenName ?? "").trim();
  if (
    currentTargetTokenId === normalizedTargetTokenId &&
    currentTargetTokenName === normalizedTargetTokenName
  ) {
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey ??= structuredClone(getTrackerData(token).odyssey);
    next.odyssey.attackDraft ??= { targetTokenId: "", targetTokenName: "" };
    next.odyssey.attackDraft.targetTokenId = normalizedTargetTokenId;
    next.odyssey.attackDraft.targetTokenName = normalizedTargetTokenName;
    return next;
  });
}

async function assignAttackTarget(attacker, target, successMessage) {
  if (!attacker || !isCharacterToken(attacker)) {
    setStatus("Select an attacker token first.", "error");
    return false;
  }

  if (!canUseToken(attacker)) {
    setStatus("You cannot roll for this attacker token.", "error");
    return false;
  }

  if (!target || !isCharacterToken(target)) {
    setStatus("Choose a valid target token.", "error");
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
  saveAttackDraftValue(attacker.id, "targetTokenName", getCharacterName(target));
  await persistAttackTargetToken(attacker.id, target.id);
  activeTokenId = attacker.id;
  render();

  const targetField = ui.selectedTokenPanel.querySelector('[data-attack-field="targetTokenId"]');
  if (targetField instanceof HTMLInputElement || targetField instanceof HTMLSelectElement) {
    targetField.value = target.id;
  }

  await syncTargetHighlight();
  setStatus(successMessage || `Target set to ${getCharacterName(target)}.`, "success");
  return true;
}

function buildCircleCommands(radius, segments = 28) {
  const safeRadius = Math.max(radius, 8);
  const commands = [];

  for (let index = 0; index <= segments; index += 1) {
    const angle = (Math.PI * 2 * index) / segments - Math.PI / 2;
    const x = Math.cos(angle) * safeRadius;
    const y = Math.sin(angle) * safeRadius;

    if (index === 0) {
      commands.push([Command.MOVE, x, y]);
    } else {
      commands.push([Command.LINE, x, y]);
    }
  }

  commands.push([Command.CLOSE]);
  return commands;
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
  const diameter = Math.max(24, Math.min(width, height));
  const radius = diameter / 2;
  const center = bounds?.center ?? targetToken.position;

  return buildPath()
    .name(`Attack Target: ${getCharacterName(targetToken)}`)
    .commands(buildCircleCommands(radius))
    .position(center)
    .rotation(0)
    .attachedTo(targetToken.id)
    .disableAttachmentBehavior(["ROTATION"])
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
  if (!attacker || !isCharacterToken(attacker) || !canViewAttackBlock(attacker)) {
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

function isLocalSelfViewItem(item) {
  return Boolean(item?.metadata?.[LOCAL_SELF_VIEW_KEY]);
}

function buildLocalSelfViewSignature(token) {
  return [
    token.lastModified,
    token.position?.x ?? 0,
    token.position?.y ?? 0,
    token.rotation ?? 0,
    token.scale?.x ?? 1,
    token.scale?.y ?? 1,
    token.zIndex ?? 0,
  ].join("|");
}

async function buildLocalSelfViewItem(token) {
  return buildImage(token.image, token.grid)
    .name(`${getCharacterName(token)} (Local View)`)
    .position(token.position)
    .rotation(token.rotation ?? 0)
    .scale(token.scale ?? { x: 1, y: 1 })
    .zIndex((token.zIndex ?? 0) + 1)
    .visible(true)
    .layer("CHARACTER")
    .locked(true)
    .disableHit(true)
    .metadata({
      [LOCAL_SELF_VIEW_KEY]: {
        tokenId: token.id,
        ownerPlayerId: playerId,
        signature: buildLocalSelfViewSignature(token),
      },
    })
    .build();
}

async function clearLocalSelfViews() {
  const localItems = await OBR.scene.local.getItems();
  const localViewIds = localItems.filter(isLocalSelfViewItem).map((item) => item.id);
  if (localViewIds.length) {
    await OBR.scene.local.deleteItems(localViewIds);
  }
}

async function syncLocalOwnedHiddenTokenViews(items = sceneItems) {
  if (playerRole === "GM" || !playerId) {
    await clearLocalSelfViews();
    return;
  }

  const desiredTokens = items.filter(
    (item) => isCharacterToken(item) && item.visible === false && canUseToken(item),
  );
  const desiredTokenIds = new Set(desiredTokens.map((token) => token.id));
  const localItems = await OBR.scene.local.getItems();
  const existingViews = localItems.filter(isLocalSelfViewItem);
  const staleViewIds = existingViews
    .filter((item) => !desiredTokenIds.has(String(item.metadata?.[LOCAL_SELF_VIEW_KEY]?.tokenId ?? "")))
    .map((item) => item.id);

  if (staleViewIds.length) {
    await OBR.scene.local.deleteItems(staleViewIds);
  }

  const currentLocalItems = staleViewIds.length ? await OBR.scene.local.getItems() : localItems;
  const currentViews = currentLocalItems.filter(isLocalSelfViewItem);
  const viewIdsToReplace = [];
  const itemsToAdd = [];

  for (const token of desiredTokens) {
    const existingView = currentViews.find(
      (item) => item.metadata?.[LOCAL_SELF_VIEW_KEY]?.tokenId === token.id,
    );
    const nextSignature = buildLocalSelfViewSignature(token);
    if (existingView?.metadata?.[LOCAL_SELF_VIEW_KEY]?.signature === nextSignature) {
      continue;
    }
    if (existingView) {
      viewIdsToReplace.push(existingView.id);
    }
    itemsToAdd.push(await buildLocalSelfViewItem(token));
  }

  if (viewIdsToReplace.length) {
    await OBR.scene.local.deleteItems(viewIdsToReplace);
  }
  if (itemsToAdd.length) {
    await OBR.scene.local.addItems(itemsToAdd);
  }
}

function scheduleOverlayMaintenance() {
  if (playerRole !== "GM") return;
  if (overlayMaintenanceTimer) {
    clearTimeout(overlayMaintenanceTimer);
  }
  overlayMaintenanceTimer = window.setTimeout(() => {
    overlayMaintenanceTimer = null;
    void syncTrackedOverlays().catch((error) => {
      console.warn("[Body HP] Overlay maintenance failed", error);
    });
  }, 120);
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
      const clickedTargetId = event.target?.id ?? "";
      const liveItems = clickedTargetId ? await OBR.scene.items.getItems() : [];
      const target =
        (clickedTargetId ? liveItems.find((item) => item.id === clickedTargetId) : null) ?? event.target;

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

      const assigned = await assignAttackTarget(attacker, target);
      if (!assigned) return false;
      await stopTargetPick();
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

async function removeAttackTargetContextMenu() {
  try {
    await OBR.contextMenu.remove(ATTACK_TARGET_CONTEXT_MENU_ID);
  } catch (_error) {
    // Ignore missing menu registrations from older extension builds.
  }
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
  let result = roll <= finalAttribute ? "Check Passed" : "Check Failed";
  let outcome = result === "Check Passed" ? "success" : "failure";

  if (roll === 1) {
    result = "Critical Success";
    outcome = "critical-success";
  } else if (roll === 20) {
    result = "Critical Failure";
    outcome = "critical-failure";
  }

  return {
    roll,
    baseAttribute,
    modifier: Number(modifier) || 0,
    finalAttribute,
    result,
    outcome,
  };
}

function rollSkillCheck(skillValue, modifier = 0) {
  const baseSkill = Number(skillValue) || 0;
  const rollPrimary = Math.floor(Math.random() * 100) + 1;
  const rollSecondary = Math.floor(Math.random() * 100) + 1;
  const totalPrimary = rollPrimary + baseSkill * 10 + (Number(modifier) || 0);
  const totalSecondary = rollSecondary;
  let result = totalPrimary > totalSecondary ? "Check Passed" : "Check Failed";
  let outcome = result === "Check Passed" ? "success" : "failure";

  if (rollPrimary >= 95) {
    result = "Critical Success";
    outcome = "critical-success";
  } else if (rollPrimary <= 5) {
    result = "Critical Failure";
    outcome = "critical-failure";
  }

  return {
    rollPrimary,
    rollSecondary,
    baseSkill,
    modifier: Number(modifier) || 0,
    totalPrimary,
    totalSecondary,
    result,
    outcome,
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
        <input type="text" inputmode="numeric" value="${data.odyssey.attributes[key] ?? 0}" data-action="set-odyssey-attribute" data-attribute="${escapeHtml(key)}" ${disabledAttr}>
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
  if (!canViewDiceBlock(token)) return "";

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
    delete ui.selectedTokenPanel.dataset.tokenId;
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

  ui.selectedTokenPanel.dataset.tokenId = token.id;
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
          <button type="button" data-action="reload-token-visuals" class="secondary" ${tokenLocked ? "disabled" : ""}>Reload Token</button>
        </div>
      </div>

      <div class="summary-strip">
        ${
          canViewBodyHpSummary(token)
            ? `
              <div class="stat-chip">
                <span class="chip-label">Body HP</span>
                <span class="chip-value">${totals.current}/${totals.max}</span>
              </div>
            `
            : ""
        }
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
      ${
        SHOW_EMBEDDED_PUBLIC_LOG
          ? renderCollapsibleSection(
              "Last roll summary",
              `<pre class="console-output">${lastRollText}</pre>`,
              false,
            )
          : ""
      }
      ${renderCollapsibleSection(
        "Part",
        `
          <div class="row row-gap">
            <button type="button" class="secondary" data-action="heal-limbs" ${fieldDisabled}>Heal</button>
          </div>
          <div class="body-table-wrap">
            <table class="body-table">
              <thead>
                <tr>
                  <th>Part</th>
                  <th>Current HP</th>
                  <th>Max HP</th>
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
                          <input type="text" inputmode="numeric" min="0" max="${part.max}" value="${part.current}" data-action="set-field" data-part="${escapeHtml(
                            partName
                          )}" data-field="current" ${fieldDisabled}>
                          <button type="button" data-action="change-part" data-part="${escapeHtml(
                            partName
                          )}" data-field="current" data-delta="1" ${fieldDisabled}>+</button>
                        </div>
                      </td>
                      <td>
                        <input class="compact-input" type="text" inputmode="numeric" min="0" max="99" value="${part.max}" data-action="set-field" data-part="${escapeHtml(
                          partName
                        )}" data-field="max" ${fieldDisabled}>
                      </td>
                      <td>
                        <input class="compact-input" type="text" inputmode="numeric" min="0" max="99" value="${part.armor}" data-action="set-field" data-part="${escapeHtml(
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
      ${
        canViewOverlayPreview(token)
          ? renderCollapsibleSection(
              "Overlay preview",
              `<pre class="console-output">${escapeHtml(formatOverlayPreviewText(data))}</pre>`,
              false,
            )
          : ""
      }
    </div>`;
}

function renderCharacteristicsBlock(data, disabledAttr) {
  const attributeInputs = ATTRIBUTE_UI_FIELDS.map(
    ([key, label]) => `
      <label class="field-stack">
        <span class="field-label">${escapeHtml(label)}</span>
        <input type="text" inputmode="numeric" value="${data.odyssey.attributes[key] ?? 0}" data-action="set-odyssey-attribute" data-attribute="${escapeHtml(key)}" ${disabledAttr}>
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
            <option value="${ABILITIES_SKILL_CATEGORY}">Abilities</option>
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
  const abilitiesSkillRows = renderOdysseySkillRows(
    data.odyssey,
    getAbilitiesSkillEntries(data.odyssey),
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
      <div class="field-label">Abilities</div>
      <div class="list">${abilitiesSkillRows || '<div class="empty">No abilities yet.</div>'}</div>
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
            <option value="${ABILITIES_SKILL_CATEGORY}">Abilities</option>
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

function renderEnglishWeaponsBlock(data, disabledAttr) {
  const meleeWeapons = data.odyssey.weapons?.melee ?? [];
  const weaponRows = meleeWeapons
    .map(
      (weapon, index) => `
        <div class="weapon-row">
          <input
            type="text"
            value="${escapeHtml(weapon.name)}"
            data-action="set-weapon-name"
            data-weapon-index="${index}"
            ${disabledAttr}
          >
          <input
            type="number"
            min="-99"
            max="99"
            value="${weapon.damage}"
            data-action="set-weapon-damage"
            data-weapon-index="${index}"
            ${disabledAttr}
          >
          <button
            type="button"
            class="danger"
            data-action="remove-weapon"
            data-weapon-index="${index}"
            ${disabledAttr}
          >Remove</button>
        </div>
      `,
    )
    .join("");

  return renderCollapsibleSection(
    "Weapons",
    `
      <div class="field-label">Melee Weapons</div>
      <div class="list">${weaponRows || '<div class="empty">No melee weapons yet.</div>'}</div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Weapon Name</span>
          <input type="text" data-weapon-field="new-name" placeholder="New weapon" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon Damage</span>
          <input type="number" min="-99" max="99" value="0" data-weapon-field="new-damage" ${disabledAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" class="secondary" data-action="add-weapon" ${disabledAttr}>Add Weapon</button>
      </div>
    `,
    false,
  );
}

function renderEnglishAttackBlock(token, data, tokenLocked) {
  if (!canViewAttackBlock(token)) return "";

  const targetCharacters = getCharacters().filter(
    (item) => item.id !== token.id && item.visible !== false,
  );
  const disabledAttr = tokenLocked ? "disabled" : "";
  const pickDisabledAttr = tokenLocked || !targetCharacters.length ? "disabled" : "";
  const draft = getAttackDraft(token, data, targetCharacters);
  const skillOptions = buildSkillOptions(getAttackSkillEntries(data.odyssey), draft.skill);
  const weaponOptions = buildWeaponOptions(getAttackSelectableWeapons(token), draft.weaponName);
  const selectedTarget = targetCharacters.find((target) => target.id === draft.targetTokenId) ?? null;
  const targetableBodyParts = getTargetableBodyParts(selectedTarget ? getTrackerData(selectedTarget) : null);
  const targetName = selectedTarget
    ? getCharacterName(selectedTarget)
    : draft.targetTokenName || "No target selected";
  const isPickingTarget = targetPickState.active && targetPickState.attackerTokenId === token.id;
  const attackDisabledAttr = tokenLocked || !draft.targetTokenId ? "disabled" : "";

  return renderCollapsibleSection(
    "Attack",
    `
      <div class="form-grid">
        <input type="hidden" value="${escapeHtml(draft.targetTokenId)}" data-attack-field="targetTokenId">
        <label class="field-stack">
          <span class="field-label">Attack Skill</span>
          <select data-attack-field="skill" ${disabledAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon</span>
          <select data-attack-field="weaponName" ${disabledAttr}>${weaponOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Current Target</span>
          <div class="hint-box">${escapeHtml(targetName)}</div>
        </label>
        <label class="field-stack">
          <span class="field-label">Pick On Map</span>
          <button type="button" class="secondary" data-action="pick-attack-target" ${pickDisabledAttr}>
            ${isPickingTarget ? "Cancel Target Pick" : "Pick Target On Map"}
          </button>
        </label>
        <label class="field-stack">
          <span class="field-label">Target Body Part</span>
          <select data-attack-field="targetPart" ${disabledAttr}>
            ${targetableBodyParts.map(
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
        <label class="field-stack">
          <span class="field-label">Parry Mode</span>
          <select data-attack-field="parryMode" ${disabledAttr}>
            <option value="off" ${draft.parryMode === "off" ? "selected" : ""}>Do Not Count Parry</option>
            ${[1, 2, 3, 4, 5]
              .map(
                (value) =>
                  `<option value="${value}" ${String(value) === draft.parryMode ? "selected" : ""}>${value} Opponent${value === 1 ? "" : "s"}</option>`
              )
              .join("")}
          </select>
        </label>
      </div>
      <div class="muted">${
        targetCharacters.length
          ? "Attack goes from the selected attacker token to the saved target chosen on the map."
          : "No visible target tokens found. You can still keep a saved target or use No Target Attack below."
      }</div>
      <div class="muted">Automatic called-shot penalties: Head -30, arms/legs -15.</div>
      <div class="muted">Strength is added to weapon damage only for attack skills with STR Bonus enabled. ${escapeHtml(PARRY_SKILL_NAME)} is added to defense only for melee attacks.</div>
      <div class="row row-gap">
        <button type="button" class="success" data-action="perform-attack" ${attackDisabledAttr}>Attack</button>
      </div>
    `,
    true,
  );
}

function renderEnglishNoTargetAttackBlock(token, data, tokenLocked) {
  if (!canViewAttackBlock(token)) return "";

  const targetCharacters = getCharacters().filter(
    (item) => item.id !== token.id && item.visible !== false,
  );
  const draft = getAttackDraft(token, data, targetCharacters);
  const skillOptions = buildSkillOptions(getAttackSkillEntries(data.odyssey), draft.skill);
  const weaponOptions = buildWeaponOptions(getAttackSelectableWeapons(token), draft.weaponName);
  const disabledAttr = tokenLocked ? "disabled" : "";

  return renderCollapsibleSection(
    "No Target Attack",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Attack Skill</span>
          <select data-manual-attack-field="skill" ${disabledAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon</span>
          <select data-manual-attack-field="weaponName" ${disabledAttr}>${weaponOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon Damage</span>
          <input type="number" value="${draft.weaponDamage}" data-manual-attack-field="weaponDamage" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack Bonus</span>
          <input type="number" value="${draft.manualAttackBonuses}" data-manual-attack-field="attackBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack Penalty</span>
          <input type="number" value="${draft.manualAttackPenalties}" data-manual-attack-field="attackPenalties" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Armor</span>
          <input type="number" min="0" max="99" value="${draft.manualArmor}" data-manual-attack-field="manualArmor" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Parry</span>
          <input type="number" min="0" max="10" value="${draft.manualParry}" data-manual-attack-field="manualParry" ${disabledAttr}>
        </label>
      </div>
      <div class="muted">Uses the manual attack values below and the defense settings above, but ignores the saved Pick On Map target.</div>
      <div class="muted">Saved target for this token stays unchanged.</div>
      <div class="row row-gap">
        <button type="button" class="success" data-action="perform-manual-attack" ${disabledAttr}>No Target Attack</button>
      </div>
    `,
    false,
  );
}

function renderEnglishDiceBlock(token, data, tokenLocked) {
  if (!canViewDiceBlock(token)) return "";

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
    delete ui.selectedTokenPanel.dataset.tokenId;
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

  ui.selectedTokenPanel.dataset.tokenId = token.id;
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
          <button type="button" data-action="reload-token-visuals" class="secondary" ${tokenLocked ? "disabled" : ""}>Reload Token</button>
        </div>
      </div>

      <div class="summary-strip">
        ${
          canViewBodyHpSummary(token)
            ? `
              <div class="stat-chip">
                <span class="chip-label">Body HP</span>
                <span class="chip-value">${totals.current}/${totals.max}</span>
              </div>
            `
            : ""
        }
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
      ${isEditable() ? renderEnglishWeaponsBlock({ odyssey }, gmOnlyDisabled) : ""}
      ${renderEnglishAttackBlock(token, { odyssey }, tokenLocked)}
      ${renderEnglishNoTargetAttackBlock(token, { odyssey }, tokenLocked)}
      ${renderEnglishDiceBlock(token, { odyssey }, tokenLocked)}
      ${renderPrivateGmDiceBlock()}
      ${
        SHOW_EMBEDDED_PUBLIC_LOG
          ? renderCollapsibleSection(
              "Last Roll",
              `<pre class="console-output">${lastRollText}</pre>`,
              false,
            )
          : ""
      }
      ${
        showPartBlock
          ? renderCollapsibleSection(
              "Body Parts",
              `
                <div class="row row-gap">
                  <button type="button" class="secondary" data-action="heal-limbs" ${bodyFieldDisabled}>Heal</button>
                </div>
                <div class="body-table-wrap">
                  <table class="body-table">
                    <thead>
                      <tr>
                        <th>Body Part</th>
                        <th>Current HP</th>
                        <th>Max HP</th>
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
                                <input type="text" inputmode="numeric" min="0" max="${part.max}" value="${part.current}" data-action="set-field" data-part="${escapeHtml(
                                  partName
                                )}" data-field="current" ${bodyFieldDisabled}>
                                <button type="button" data-action="change-part" data-part="${escapeHtml(
                                  partName
                                )}" data-field="current" data-delta="1" ${bodyFieldDisabled}>+</button>
                              </div>
                            </td>
                            <td>
                              <input class="compact-input" type="text" inputmode="numeric" min="0" max="99" value="${part.max}" data-action="set-field" data-part="${escapeHtml(
                                partName
                              )}" data-field="max" ${bodyFieldDisabled}>
                            </td>
                            <td>
                              <input class="compact-input" type="text" inputmode="numeric" min="0" max="99" value="${part.armor}" data-action="set-field" data-part="${escapeHtml(
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
      ${
        canViewOverlayPreview(token)
          ? renderCollapsibleSection(
              "Overlay Preview",
              `<pre class="console-output">${escapeHtml(formatOverlayPreviewText(data))}</pre>`,
              false,
            )
          : ""
      }
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
  ui.clearDebugBtn?.classList.toggle("hidden", playerRole !== "GM");
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

  await syncLocalOwnedHiddenTokenViews(sceneItems);
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

async function reloadTokenVisuals() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot reload this token.", "error");
    return;
  }

  const currentItems = await OBR.scene.items.getItems();
  const liveToken = currentItems.find((item) => item.id === token.id) ?? token;
  if (liveToken.visible === false) {
    await syncLocalOwnedHiddenTokenViews(currentItems);
  } else {
    await removeOverlaysForToken(token.id, currentItems);
    await ensureOverlayForToken(token.id);
  }

  sceneItems = await OBR.scene.items.getItems();
  await syncLocalOwnedHiddenTokenViews(sceneItems);
  render();
  await syncTargetHighlight();
  setStatus(`${getCharacterName(token)} visuals reloaded.`, "success");
}

async function healLimbs() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can heal this token.", "error");
    return;
  }

  const healedParts = ["L.Arm", "R.Arm", "Torso", "L.Leg", "R.Leg"];
  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    for (const partName of healedParts) {
      const part = next.body?.[partName];
      if (!part) continue;
      part.current = part.max;
      part.minor = 0;
      part.serious = 0;
    }
    return next;
  });
  await ensureOverlayForToken(token.id);
  await syncState();
  setStatus(`${getCharacterName(token)} healed.`, "success");
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
    next.odyssey.attributes[attribute] = clamp(Number(value) || 0, 0, 20);
    return next;
  });
  await syncState();
}

async function addWeapon() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit weapons.", "error");
    return;
  }

  const name = getActionFieldValue('[data-weapon-field="new-name"]').trim() || "New Weapon";
  const damage = clamp(Number(getActionFieldValue('[data-weapon-field="new-damage"]')) || 0, -99, 99);

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.weapons.melee ??= [];
    next.odyssey.weapons.melee.push({ name, damage });
    return next;
  });
  await syncState();
  syncAttackWeaponInputs(token.id, name, damage);
  setStatus(`Weapon "${name}" saved for ${getCharacterName(token)}.`, "success");
}

async function setWeaponDamage(index, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit weapons.", "error");
    return;
  }

  const currentWeaponName = getOdysseyData(token).weapons?.melee?.[index]?.name ?? "Default";
  const nextDamage = clamp(Number(value) || 0, -99, 99);

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.weapons.melee ??= [];
    if (!next.odyssey.weapons.melee[index]) {
      next.odyssey.weapons.melee[index] = { name: "Default", damage: 0 };
    }
    next.odyssey.weapons.melee[index].damage = nextDamage;
    return next;
  });
  await syncState();
  const currentDraft = attackFormDrafts.get(token.id) ?? {};
  if (currentDraft.weaponName === currentWeaponName) {
    syncAttackWeaponInputs(token.id, currentWeaponName, nextDamage);
  }
}

async function setWeaponName(index, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit weapons.", "error");
    return;
  }

  const previousWeapon = getOdysseyData(token).weapons?.melee?.[index] ?? { name: "Default", damage: 0 };
  const nextWeaponName = String(value || "").trim() || "Default";

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.weapons.melee ??= [];
    if (!next.odyssey.weapons.melee[index]) {
      next.odyssey.weapons.melee[index] = { name: "Default", damage: 0 };
    }
    next.odyssey.weapons.melee[index].name = nextWeaponName;
    return next;
  });
  await syncState();
  const currentDraft = attackFormDrafts.get(token.id) ?? {};
  if (currentDraft.weaponName === previousWeapon.name) {
    syncAttackWeaponInputs(token.id, nextWeaponName, previousWeapon.damage);
  }
}

async function removeWeapon(index) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit weapons.", "error");
    return;
  }

  const currentWeapons = getOdysseyData(token).weapons?.melee ?? [];
  const removedWeapon = currentWeapons[index];
  if (!removedWeapon) {
    setStatus("Weapon not found.", "error");
    return;
  }

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.odyssey.weapons.melee = (next.odyssey.weapons.melee ?? []).filter(
      (_weapon, weaponIndex) => weaponIndex !== index,
    );
    return next;
  });
  await syncState();
  const refreshedToken = getCharacterById(token.id);
  const fallbackWeapon = refreshedToken
    ? getDefaultAttackWeapon(refreshedToken)
    : { name: UNARMED_WEAPON_NAME, damage: 0 };
  syncAttackWeaponInputs(token.id, fallbackWeapon.name, fallbackWeapon.damage);
  setStatus(`Weapon "${removedWeapon.name}" removed.`, "success");
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
      next.odyssey.attributes[draft.attribute] = clamp(Number(draft.value) || 0, 0, 20);
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

async function performAttack({ manualDefense = false } = {}) {
  const attacker = getCharacterById(activeTokenId);
  if (!attacker) {
    setStatus("Select an attacker token first.", "error");
    return;
  }
  if (!canUseToken(attacker)) {
    setStatus("You cannot roll for this attacker token.", "error");
    return;
  }

  const savedTargetTokenId = getActionFieldValue('[data-attack-field="targetTokenId"]') || "";
  const targetTokenId = manualDefense ? "" : savedTargetTokenId;
  const target = targetTokenId ? getCharacterById(targetTokenId) : null;
  if (!manualDefense && !targetTokenId) {
    setStatus("Pick a target on the map first.", "error");
    return;
  }
  if (targetTokenId && !target) {
    setStatus("Choose a valid target token.", "error");
    return;
  }
  if (target?.visible === false) {
    setStatus("Hidden tokens cannot be targeted.", "error");
    return;
  }
  if (target && target.id === attacker.id) {
    setStatus("Attacker and target must be different tokens.", "error");
    return;
  }

  const attackerOdyssey = getOdysseyData(attacker);
  const targetData = target ? getTrackerData(target) : null;
  const targetOdyssey = target ? getOdysseyData(target) : null;
  const skillName =
    manualDefense
      ? getActionFieldValue('[data-manual-attack-field="skill"]') || getActionFieldValue('[data-attack-field="skill"]')
      : getActionFieldValue('[data-attack-field="skill"]');
  const weaponName =
    manualDefense
      ? getActionFieldValue('[data-manual-attack-field="weaponName"]') || getActionFieldValue('[data-attack-field="weaponName"]')
      : getActionFieldValue('[data-attack-field="weaponName"]');
  const requestedTargetPart = getActionFieldValue('[data-attack-field="targetPart"]');
  const weaponDamage = Number(
    manualDefense
      ? getActionFieldValue('[data-manual-attack-field="weaponDamage"]') || getActionFieldValue('[data-attack-field="weaponDamage"]')
      : getActionFieldValue('[data-attack-field="weaponDamage"]')
  ) || 0;
  const attackBonuses = Number(
    manualDefense
      ? getActionFieldValue('[data-manual-attack-field="attackBonuses"]') || getActionFieldValue('[data-attack-field="attackBonuses"]')
      : getActionFieldValue('[data-attack-field="attackBonuses"]')
  ) || 0;
  const manualAttackPenalties = Number(
    manualDefense
      ? getActionFieldValue('[data-manual-attack-field="attackPenalties"]') || getActionFieldValue('[data-attack-field="attackPenalties"]')
      : getActionFieldValue('[data-attack-field="attackPenalties"]')
  ) || 0;
  const availableTargetParts = getTargetableBodyParts(targetData);
  const targetPart = availableTargetParts.includes(requestedTargetPart) ? requestedTargetPart : "Torso";
  const automaticTargetPenalty = getAutomaticTargetPenalty(targetPart);
  const totalAttackPenalties = manualAttackPenalties + automaticTargetPenalty;
  const defenseBonuses = Number(getActionFieldValue('[data-attack-field="defenseBonuses"]')) || 0;
  const defensePenalties = Number(getActionFieldValue('[data-attack-field="defensePenalties"]')) || 0;
  const manualArmor = clamp(
    Number(
      manualDefense
        ? getActionFieldValue('[data-manual-attack-field="manualArmor"]') || getActionFieldValue('[data-attack-field="manualArmor"]')
        : getActionFieldValue('[data-attack-field="manualArmor"]')
    ) || 0,
    0,
    99,
  );
  const manualParry = clamp(
    Number(
      manualDefense
        ? getActionFieldValue('[data-manual-attack-field="manualParry"]') || getActionFieldValue('[data-attack-field="manualParry"]')
        : getActionFieldValue('[data-attack-field="manualParry"]')
    ) || 0,
    0,
    10,
  );
  const parryMode = getActionFieldValue('[data-attack-field="parryMode"]') || "1";
  const parryDivisor = getParryDivisor(parryMode);
  saveAttackDraftValue(attacker.id, "skill", skillName);
  saveAttackDraftValue(attacker.id, "weaponName", weaponName);
  if (!manualDefense) {
    saveAttackDraftValue(attacker.id, "targetTokenId", targetTokenId);
    saveAttackDraftValue(attacker.id, "targetTokenName", target ? getCharacterName(target) : "");
    await persistAttackTargetToken(attacker.id, targetTokenId);
  }
  saveAttackDraftValue(attacker.id, "targetPart", targetPart);
  saveAttackDraftValue(attacker.id, "weaponDamage", String(weaponDamage));
  saveAttackDraftValue(
    attacker.id,
    manualDefense ? "manualAttackBonuses" : "attackBonuses",
    manualDefense
      ? getActionFieldValue('[data-manual-attack-field="attackBonuses"]')
      : getActionFieldValue('[data-attack-field="attackBonuses"]'),
  );
  saveAttackDraftValue(
    attacker.id,
    manualDefense ? "manualAttackPenalties" : "attackPenalties",
    manualDefense
      ? getActionFieldValue('[data-manual-attack-field="attackPenalties"]')
      : getActionFieldValue('[data-attack-field="attackPenalties"]'),
  );
  saveAttackDraftValue(attacker.id, "defenseBonuses", getActionFieldValue('[data-attack-field="defenseBonuses"]'));
  saveAttackDraftValue(attacker.id, "defensePenalties", getActionFieldValue('[data-attack-field="defensePenalties"]'));
  saveAttackDraftValue(
    attacker.id,
    "manualArmor",
    manualDefense
      ? getActionFieldValue('[data-manual-attack-field="manualArmor"]')
      : getActionFieldValue('[data-attack-field="manualArmor"]'),
  );
  saveAttackDraftValue(
    attacker.id,
    "manualParry",
    manualDefense
      ? getActionFieldValue('[data-manual-attack-field="manualParry"]')
      : getActionFieldValue('[data-attack-field="manualParry"]'),
  );
  saveAttackDraftValue(attacker.id, "parryMode", parryMode);
  const specialPartState = targetData?.body?.[SPECIAL_PART_NAME] ?? null;
  const specialWasActive =
    Boolean(target) &&
    hasConfiguredSpecial(targetData) &&
    (Number(specialPartState?.max) || 0) > 0 &&
    (Number(specialPartState?.current) || 0) > 0;
  const targetArmor = target
    ? (Number(targetData?.body?.[targetPart]?.armor) || 0) +
      (specialWasActive ? Number(specialPartState?.armor) || 0 : 0)
    : manualArmor;
  const targetPartState = targetData?.body?.[targetPart] ?? { current: 0, max: 0, armor: 0, minor: 0, serious: 0 };
  const beforeHp = target ? (targetPartState.current ?? 0) : null;
  const beforeMinor = target ? (targetPartState.minor ?? 0) : null;
  const beforeSerious = target ? (targetPartState.serious ?? 0) : null;
  const specialBeforeHp = specialWasActive ? Number(specialPartState?.current) || 0 : null;
  const strengthBonus = getSkillStrengthBonusFlag(attackerOdyssey, skillName)
    ? Math.max((attackerOdyssey.attributes.Strength ?? 0) - 10, 0)
    : 0;
  const finalWeaponDamage = weaponDamage + strengthBonus;
  const baseTargetParry =
    skillName === MELEE_SKILL_NAME
      ? target
        ? (targetOdyssey?.skills?.[PARRY_SKILL_NAME] ?? 0)
        : manualParry
      : 0;
  const targetParry =
    parryDivisor <= 0
      ? 0
      : Math.max(Math.floor(baseTargetParry / parryDivisor), 0);

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
  const specialResolution =
    target && result.hit && result.damage
      ? projectDamageWithSpecialProtection({
          specialPart: specialPartState,
          targetPart: targetPartState,
          damage: result.damage,
          targetPartName: targetPart,
        })
      : {
          specialProjectedState: specialPartState ? getNormalizedPartState(specialPartState) : null,
          projectedTargetState: {
            ...(target ? getNormalizedPartState(targetPartState) : {}),
            critApplied: 0,
          },
          specialActive: false,
          specialArmor: 0,
          damageAppliedLabel: "No Damage",
        };
  const projectedPartState = specialResolution.projectedTargetState;
  const projectedSpecialState = specialResolution.specialProjectedState;
  const afterHp = target ? (projectedPartState.current ?? beforeHp) : null;
  const afterMinor = target ? (projectedPartState.minor ?? beforeMinor) : null;
  const afterSerious = target ? (projectedPartState.serious ?? beforeSerious) : null;
  const specialAfterHp = specialWasActive ? (projectedSpecialState?.current ?? specialBeforeHp) : null;
  const resolvedTargetName = target ? getCharacterName(target) : "Manual Defense";
  const resolvedAttackSummary =
    specialResolution.specialActive &&
    result.hit &&
    specialResolution.damageAppliedLabel !== "No Damage"
      ? `${result.summary} Applied: ${specialResolution.damageAppliedLabel}.`
      : result.summary;

  await updateTrackerData(attacker.id, (current) => {
    const next = structuredClone(current);
    next.lastRoll = {
      eventId: 0,
      actorName: playerName || "Owlbear Player",
      summary: `${getCharacterName(attacker)} -> ${resolvedTargetName}: ${resolvedAttackSummary}`,
      outcome: result.outcome,
      total: result.attackTotal,
      targetPart: result.targetPart,
      timestamp: new Date().toISOString(),
      source: "owlbear-extension",
    };
    next.history = [next.lastRoll, ...(next.history ?? [])].slice(0, 12);
    return next;
  });

  if (target) {
    await updateTrackerData(target.id, (current) => {
      const next = structuredClone(current);
      if (specialResolution.specialActive && next.body[SPECIAL_PART_NAME]) {
        next.body[SPECIAL_PART_NAME].current = projectedSpecialState.current;
        next.body[SPECIAL_PART_NAME].minor = projectedSpecialState.minor;
        next.body[SPECIAL_PART_NAME].serious = projectedSpecialState.serious;
      }
      if (result.hit && next.body[result.targetPart]) {
        next.body[result.targetPart].current = projectedPartState.current;
        next.body[result.targetPart].minor = projectedPartState.minor;
        next.body[result.targetPart].serious = projectedPartState.serious;
      }
      next.lastRoll = {
        eventId: 0,
        actorName: getCharacterName(attacker),
        summary: resolvedAttackSummary,
        outcome: result.outcome,
        total: result.attackTotal,
        targetPart: result.targetPart,
        timestamp: new Date().toISOString(),
        source: "owlbear-extension",
      };
      next.history = [next.lastRoll, ...(next.history ?? [])].slice(0, 12);
      return next;
    });
  }

  await ensureOverlayForToken(attacker.id);
  if (target) {
    await ensureOverlayForToken(target.id);
  }
  await pushDebugEntry(
    `${getAttackOutcomeIcon(result.outcome)} ${getCharacterName(attacker)} attacks ${resolvedTargetName}`,
    formatAttackDebug({
      attackerName: getCharacterName(attacker),
      targetName: resolvedTargetName,
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
      baseTargetParry,
      targetParry,
      parryMode,
      targetArmor,
      specialArmor: specialResolution.specialArmor,
      specialActive: specialResolution.specialActive,
      result,
      beforeHp,
      afterHp,
      specialBeforeHp,
      specialAfterHp,
      beforeMinor,
      afterMinor,
      beforeSerious,
      afterSerious,
      critApplied: projectedPartState.critApplied ?? 0,
      damageAppliedLabel: specialResolution.damageAppliedLabel,
    }),
    result.hit ? "success" : "info",
  );
  await syncState();
  setStatus(`${getCharacterName(attacker)} -> ${resolvedTargetName}: ${resolvedAttackSummary}`, result.hit ? "success" : "info");
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
  const summary = buildDiceRollSummary(diceLabel, result);

  await updateTrackerData(token.id, (current) => {
    const next = structuredClone(current);
    next.lastRoll = {
      eventId: 0,
      actorName: playerName || "Owlbear Player",
      summary,
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
  setStatus(summary, "success");
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
        : getActionFieldValue('[data-skill-field="new-category"]') === ABILITIES_SKILL_CATEGORY
          ? ABILITIES_SKILL_CATEGORY
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
  const summary = `${getResolvedCheckResultIcon(result.result)} Characteristic ${attributeLabel}: ${result.roll} vs ${result.finalAttribute} (${result.result})`;

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
    `${getResolvedCheckResultIcon(result.result)} ${getCharacterName(token)} rolls characteristic`,
    formatRollCharDebug({
      tokenName: getCharacterName(token),
      attributeLabel,
      result,
    }),
    isResolvedCheckResultSuccess(result.result) ? "success" : result.result === "Critical Failure" ? "error" : "info",
  );
  await syncState();
  setStatus(summary, isResolvedCheckResultSuccess(result.result) ? "success" : "error");
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
  const summary = `${getResolvedCheckResultIcon(result.result)} Skill ${skillName}: ${result.totalPrimary} vs ${result.totalSecondary} (${result.result})`;

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
    `${getResolvedCheckResultIcon(result.result)} ${getCharacterName(token)} checks skill`,
    formatRollSkillDebug({
      tokenName: getCharacterName(token),
      skillName,
      result,
    }),
    isResolvedCheckResultSuccess(result.result) ? "success" : result.result === "Critical Failure" ? "error" : "info",
  );
  await syncState();
  setStatus(summary, isResolvedCheckResultSuccess(result.result) ? "success" : "error");
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
  const summary = buildDiceRollSummary(diceLabel, result);

  pushPrivateGmEntry(
    `GM private ${diceLabel}`,
    formatDiceDebug({
      tokenName: "GM private roll",
      result,
    }),
  );
  render();
  setStatus(`Private GM roll. ${summary}`, "success");
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

    if (action === "reload-token-visuals") {
      void reloadTokenVisuals().catch((error) => {
        setStatus(error?.message ?? "Unable to reload token visuals.", "error");
      });
      return;
    }

    if (action === "heal-limbs") {
      void healLimbs().catch((error) => {
        setStatus(error?.message ?? "Unable to heal limbs.", "error");
      });
      return;
    }

    if (action === "pick-attack-target") {
      void startTargetPick().catch((error) => {
        setStatus(error?.message ?? "Unable to start target picking.", "error");
      });
      return;
    }

    if (action === "clear-debug-console") {
      void clearDebugConsole().catch((error) => {
        setStatus(error?.message ?? "Unable to clear debug console.", "error");
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
      return;
    }

    if (action === "perform-manual-attack") {
      void performAttack({ manualDefense: true }).catch((error) => {
        setStatus(error?.message ?? "Unable to resolve no target attack.", "error");
      });
      return;
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

    if (action === "add-weapon") {
      void addWeapon().catch((error) => {
        setStatus(error?.message ?? "Unable to add weapon.", "error");
      });
      return;
    }

    if (action === "remove-skill" && skill) {
      void removeOdysseySkill(skill).catch((error) => {
        setStatus(error?.message ?? "Unable to remove skill.", "error");
      });
      return;
    }

    if (action === "remove-weapon") {
      const weaponIndex = Number(actionNode.dataset.weaponIndex ?? -1);
      if (weaponIndex < 0) return;
      void removeWeapon(weaponIndex).catch((error) => {
        setStatus(error?.message ?? "Unable to remove weapon.", "error");
      });
    }
  });

  document.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement)) return;

    if (target.dataset.attackField && activeTokenId) {
      if (target.dataset.attackField === "weaponName") {
        const token = getCharacterById(activeTokenId);
        const selectedWeapon = token ? getWeaponByName(token, target.value) : null;
        if (selectedWeapon) {
          syncAttackWeaponInputs(activeTokenId, selectedWeapon.name, selectedWeapon.damage);
        }
        return;
      }
      saveAttackDraftValue(activeTokenId, target.dataset.attackField, target.value);
      if (target.dataset.attackField === "targetTokenId") {
        const selectedTarget = target.value ? getCharacterById(target.value) : null;
        saveAttackDraftValue(
          activeTokenId,
          "targetTokenName",
          selectedTarget ? getCharacterName(selectedTarget) : "",
        );
        void persistAttackTargetToken(activeTokenId, target.value).catch((error) => {
          console.warn("[Body HP] Unable to persist attack target", error);
        });
        void syncTargetHighlight().catch((error) => {
          console.warn("[Body HP] Unable to sync target highlight", error);
        });
      }
    }

    if (target.dataset.manualAttackField && activeTokenId) {
      if (target.dataset.manualAttackField === "weaponName") {
        const token = getCharacterById(activeTokenId);
        const selectedWeapon = token ? getWeaponByName(token, target.value) : null;
        if (selectedWeapon) {
          syncAttackWeaponInputs(activeTokenId, selectedWeapon.name, selectedWeapon.damage);
        }
        return;
      }
      const draftField = getSharedAttackDraftField(target.dataset.manualAttackField);
      if (draftField) {
        saveAttackDraftValue(activeTokenId, draftField, target.value);
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

    if (target.dataset.manualAttackField) {
      const draftField = getSharedAttackDraftField(target.dataset.manualAttackField);
      if (draftField) {
        saveAttackDraftValue(activeTokenId, draftField, target.value);
      }
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
    await removeAttackTargetContextMenu();
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
      void syncLocalOwnedHiddenTokenViews(items).catch((error) => {
        console.warn("[Body HP] Unable to sync local hidden-token views", error);
      });
      void syncTargetHighlight().catch((error) => {
        console.warn("[Body HP] Unable to sync target highlight", error);
      });
      scheduleOverlayMaintenance();
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

      if (payload.type === "debug-clear") {
        if (pendingLocalDebugClear) {
          pendingLocalDebugClear = false;
          return;
        }
        debugEntries = [];
        renderDebugConsole();
        return;
      }

      if (payload.type !== "debug-entry") return;
      if (pendingLocalDebugEntryIds.has(payload.entry?.id)) {
        pendingLocalDebugEntryIds.delete(payload.entry.id);
        return;
      }

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
      debugEntries = sanitizeDebugEntries(metadata?.[DEBUG_LOG_KEY]);
      renderDebugConsole();
    });
  } catch (error) {
    setStatus(error?.message ?? "Extension failed to initialize.", "error");
  }
});
