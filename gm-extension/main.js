import {
  BODY_ORDER,
  OBR,
  PARRY_SKILL_NAME,
  SPECIAL_PART_NAME,
  clamp,
  ensureOverlayForToken,
  getCharacterName,
  getOdysseyData,
  getTargetableBodyParts,
  getTrackerData,
  hasConfiguredSpecial,
  isCharacterToken,
  sortCharacters,
  updateTrackerData,
} from "../shared.js";
import {
  formatAttackOutcomeLabel,
  getAttackOutcomeIcon,
  resolveAttack,
  rollDice,
} from "../odyssey_rules.js";

const DEBUG_LOG_KEY = "com.codex.body-hp/debugLog";
const DEBUG_BROADCAST_CHANNEL = "com.codex.body-hp/debug";
const DEBUG_ENTRY_LIMIT = 50;
const TARGET_PICK_TOOL_ID = "com.codex.body-hp/gm-target-picker";
const TARGET_PICK_MODE_ID = "pick-gm-target";
const DEFAULT_TARGET_PART = "Torso";
const EXTENSION_ICON_URL = new URL("../icon.svg", window.location.href).href;
const TARGET_PICK_CURSOR = `url("data:image/svg+xml,${encodeURIComponent(
  `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32">
    <circle cx="16" cy="16" r="3.25" fill="#ef4444" stroke="#7f1d1d" stroke-width="1.5"/>
    <path d="M16 2v8M16 22v8M2 16h8M22 16h8" stroke="#ef4444" stroke-width="2.5" stroke-linecap="round"/>
    <path d="M16 6v4M16 22v4M6 16h4M22 16h4" stroke="#7f1d1d" stroke-width="1.2" stroke-linecap="round"/>
  </svg>`,
)}") 16 16, crosshair`;

const ui = {
  roleBadge: document.getElementById("roleBadge"),
  refreshBtn: document.getElementById("refreshBtn"),
  statusBox: document.getElementById("statusBox"),
  gmOnlyNotice: document.getElementById("gmOnlyNotice"),
  gmContent: document.getElementById("gmContent"),
  publicDiceSides: document.getElementById("publicDiceSides"),
  publicDiceCount: document.getElementById("publicDiceCount"),
  publicDiceModifier: document.getElementById("publicDiceModifier"),
  publicDiceBtn: document.getElementById("publicDiceBtn"),
  privateDiceSides: document.getElementById("privateDiceSides"),
  privateDiceCount: document.getElementById("privateDiceCount"),
  privateDiceModifier: document.getElementById("privateDiceModifier"),
  privateDiceBtn: document.getElementById("privateDiceBtn"),
  privateLog: document.getElementById("privateLog"),
  attackSourceName: document.getElementById("attackSourceName"),
  attackSkill: document.getElementById("attackSkill"),
  weaponDamage: document.getElementById("weaponDamage"),
  weaponAccuracy: document.getElementById("weaponAccuracy"),
  attackBonuses: document.getElementById("attackBonuses"),
  attackPenalties: document.getElementById("attackPenalties"),
  defenseBonuses: document.getElementById("defenseBonuses"),
  defensePenalties: document.getElementById("defensePenalties"),
  parryMode: document.getElementById("parryMode"),
  targetName: document.getElementById("targetName"),
  pickTargetBtn: document.getElementById("pickTargetBtn"),
  clearTargetBtn: document.getElementById("clearTargetBtn"),
  targetPart: document.getElementById("targetPart"),
  environmentAttackBtn: document.getElementById("environmentAttackBtn"),
};

let playerRole = "PLAYER";
let playerName = "";
let sceneItems = [];
let characterList = [];
let charactersById = new Map();
let selectedTargetTokenId = "";
let privateEntries = [];

const targetPickState = {
  active: false,
  previousToolId: "",
  previousModeId: undefined,
  toolReady: false,
  restoring: false,
};

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function setStatus(message, kind = "info") {
  ui.statusBox.textContent = message;
  ui.statusBox.className = `status ${kind}`;
}

function setSceneItems(items) {
  sceneItems = Array.isArray(items) ? items : [];
  characterList = sortCharacters(sceneItems.filter(isCharacterToken));
  charactersById = new Map(characterList.map((item) => [item.id, item]));
  syncSelectedTarget();
}

function getCharacters() {
  return characterList.filter((item) => item.visible !== false);
}

function getCharacterById(id) {
  return charactersById.get(id) ?? null;
}

function getSelectedTarget() {
  return selectedTargetTokenId ? getCharacterById(selectedTargetTokenId) : null;
}

function syncSelectedTarget() {
  const target = getSelectedTarget();
  if (!target || target.visible === false || !isCharacterToken(target)) {
    selectedTargetTokenId = "";
  }
}

function formatRawDiceRolls(result) {
  return result.rolls.join(", ");
}

function formatDiceRollsWithModifier(result) {
  const modifier = Number(result.modifier) || 0;
  return result.rolls.map((roll) => (Number(roll) || 0) + modifier).join(", ");
}

function buildDiceRollSummary(diceLabel, result) {
  return `Rolled ${diceLabel}: raw [${formatRawDiceRolls(result)}], sum ${result.subtotal}, with modifier ${formatDiceRollsWithModifier(result)}`;
}

function formatDiceDebug(label, result) {
  return [
    `Actor: ${label}`,
    `Dice: ${result.count}d${result.sides}`,
    `Raw Dice: ${formatRawDiceRolls(result)}`,
    `Dice Sum: ${result.subtotal}`,
    `With Modifier: ${formatDiceRollsWithModifier(result)}`,
  ].join("\n");
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

async function pushSharedLogEntry(title, body, kind = "info") {
  if (playerRole !== "GM") {
    throw new Error("Only the GM can write to the shared Odyssey log.");
  }

  const entry = {
    id: Date.now() * 1000 + Math.floor(Math.random() * 1000),
    title,
    body,
    kind,
    timestamp: new Date().toLocaleTimeString(),
  };

  const metadata = await OBR.room.getMetadata();
  const nextEntries = mergeDebugEntries([entry], metadata?.[DEBUG_LOG_KEY]);
  await OBR.broadcast.sendMessage(
    DEBUG_BROADCAST_CHANNEL,
    { type: "debug-entry", entry },
    { destination: "ALL" },
  );
  await OBR.room.setMetadata({
    [DEBUG_LOG_KEY]: nextEntries,
  });
}

function pushPrivateEntry(title, body) {
  privateEntries = [
    {
      id: Date.now(),
      title,
      body,
      timestamp: new Date().toLocaleTimeString(),
    },
    ...privateEntries,
  ].slice(0, 20);
  renderPrivateEntries();
}

function renderPrivateEntries() {
  if (!privateEntries.length) {
    ui.privateLog.innerHTML = '<div class="empty">Private GM rolls will stay visible only here.</div>';
    return;
  }

  ui.privateLog.innerHTML = privateEntries
    .map(
      (entry) => `
        <div class="debug-entry">
          <div class="debug-head">
            <div class="debug-title">${escapeHtml(entry.title)}</div>
            <div class="muted">${escapeHtml(entry.timestamp)}</div>
          </div>
          <pre class="console-output">${escapeHtml(entry.body)}</pre>
        </div>`,
    )
    .join("");
}

function renderRoleGate() {
  ui.roleBadge.textContent = playerRole;
  const isGm = playerRole === "GM";
  ui.gmOnlyNotice.hidden = isGm;
  ui.gmContent.hidden = !isGm;
}

function renderTargetPartOptions() {
  const target = getSelectedTarget();
  const targetParts = getTargetableBodyParts(target ? getTrackerData(target) : null);
  const currentValue = ui.targetPart.value;
  const nextValue = targetParts.includes(currentValue) ? currentValue : DEFAULT_TARGET_PART;
  ui.targetPart.innerHTML = targetParts
    .map(
      (partName) =>
        `<option value="${escapeHtml(partName)}" ${partName === nextValue ? "selected" : ""}>${escapeHtml(partName)}</option>`,
    )
    .join("");
  if (targetParts.length && !targetParts.includes(ui.targetPart.value)) {
    ui.targetPart.value = nextValue;
  }
}

function renderTargetState() {
  const target = getSelectedTarget();
  const isPicking = targetPickState.active;
  ui.targetName.textContent = target ? getCharacterName(target) : "No target selected";
  ui.pickTargetBtn.textContent = isPicking ? "Cancel Target Pick" : "Pick Target On Map";
  ui.pickTargetBtn.disabled = playerRole !== "GM" || !getCharacters().length;
  ui.clearTargetBtn.disabled = playerRole !== "GM" || !target;
  ui.environmentAttackBtn.disabled = playerRole !== "GM" || !target;
  renderTargetPartOptions();
}

function render() {
  renderRoleGate();
  renderPrivateEntries();
  renderTargetState();
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

  if (totalCrit > 0) parts.push(`${totalCrit} Crit`);
  if (serious > 0) parts.push(`${serious} Serious`);
  if (minor > 0) parts.push(`${minor} Minor`);

  return parts.length ? parts.join(", ") : "No Damage";
}

function formatStateTransition(before, after) {
  if (before == null || after == null) return "-";
  return `${before} -> ${after}`;
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

function projectPartDamage(part, damage) {
  const next = getNormalizedPartState(part);

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
    : { ...normalizedTarget, critApplied: 0 };
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

function formatEnvironmentAttackDebug({
  sourceName,
  targetName,
  targetPart,
  attackSkill,
  weaponDamage,
  weaponAccuracy,
  manualAttackBonuses,
  totalAttackBonuses,
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
  damageAppliedLabel,
}) {
  const lines = [
    `Source: ${sourceName}`,
    `Target: ${targetName} -> ${targetPart}`,
    `Result: ${formatAttackOutcomeLabel(result.outcome)}`,
    `Damage Applied: ${damageAppliedLabel}`,
    "",
    `Accuracy: ${result.attackRoll} + ${attackSkill * 10} + ${totalAttackBonuses} - ${totalAttackPenalties} = ${result.attackTotal}`,
    `Defense: ${result.defenseRoll} + ${targetParry * 10} + ${defenseBonuses} - ${defensePenalties} = ${result.defenseTotal}`,
    `Damage: ${result.attackTotal} + ${weaponDamage} vs ${result.defenseTotal} + ${targetArmor}`,
    "",
    `Weapon Accuracy: ${weaponAccuracy}`,
    `Manual Attack Bonus: ${manualAttackBonuses}`,
    `Manual Attack Penalty: ${manualAttackPenalties}`,
    `Auto Target Penalty: ${automaticTargetPenalty}`,
    `Parry Mode: ${getParryModeLabel(parryMode)}`,
    `Base Parry: ${baseTargetParry}`,
    `Effective Parry: ${targetParry}`,
    `Armor: ${targetArmor}`,
    `Target HP: ${formatStateTransition(beforeHp, afterHp)}`,
  ];

  if (specialActive) {
    lines.push(`Special Armor: ${specialArmor}`);
    lines.push(`Special HP: ${formatStateTransition(specialBeforeHp, specialAfterHp)}`);
  }

  return lines.join("\n");
}

async function performPublicGmRoll() {
  if (playerRole !== "GM") {
    setStatus("Only the GM can use this extension.", "error");
    return;
  }

  const dice = Number(ui.publicDiceSides.value) || 20;
  const count = Number(ui.publicDiceCount.value) || 1;
  const modifier = Number(ui.publicDiceModifier.value) || 0;
  const result = rollDice(dice, modifier, count);
  const diceLabel = `${result.count}d${result.sides}`;
  const summary = buildDiceRollSummary(diceLabel, result);

  await pushSharedLogEntry(
    `GM Dice ${diceLabel}`,
    formatDiceDebug(playerName || "GM Dice", result),
    "success",
  );
  setStatus(summary, "success");
}

function performPrivateGmRoll() {
  if (playerRole !== "GM") {
    setStatus("Only the GM can use this extension.", "error");
    return;
  }

  const dice = Number(ui.privateDiceSides.value) || 20;
  const count = Number(ui.privateDiceCount.value) || 1;
  const modifier = Number(ui.privateDiceModifier.value) || 0;
  const result = rollDice(dice, modifier, count);
  const diceLabel = `${result.count}d${result.sides}`;
  const summary = buildDiceRollSummary(diceLabel, result);

  pushPrivateEntry(`GM Private ${diceLabel}`, formatDiceDebug(playerName || "GM Private Dice", result));
  setStatus(`Private roll. ${summary}`, "success");
}

async function performEnvironmentAttack() {
  if (playerRole !== "GM") {
    setStatus("Only the GM can resolve environment attacks.", "error");
    return;
  }

  const target = getSelectedTarget();
  if (!target) {
    setStatus("Pick a target on the map first.", "error");
    return;
  }
  if (target.visible === false) {
    setStatus("Hidden tokens cannot be targeted.", "error");
    return;
  }

  const targetData = getTrackerData(target);
  const targetOdyssey = getOdysseyData(target);
  const sourceName = ui.attackSourceName.value.trim() || "Environment";
  const attackSkill = clamp(Number(ui.attackSkill.value) || 0, 0, 10);
  const weaponDamage = Number(ui.weaponDamage.value) || 0;
  const weaponAccuracy = Number(ui.weaponAccuracy.value) || 0;
  const manualAttackBonuses = Number(ui.attackBonuses.value) || 0;
  const totalAttackBonuses = manualAttackBonuses + weaponAccuracy;
  const manualAttackPenalties = Number(ui.attackPenalties.value) || 0;
  const requestedTargetPart = ui.targetPart.value || DEFAULT_TARGET_PART;
  const availableTargetParts = getTargetableBodyParts(targetData);
  const targetPart = availableTargetParts.includes(requestedTargetPart)
    ? requestedTargetPart
    : DEFAULT_TARGET_PART;
  const automaticTargetPenalty = getAutomaticTargetPenalty(targetPart);
  const totalAttackPenalties = manualAttackPenalties + automaticTargetPenalty;
  const defenseBonuses = Number(ui.defenseBonuses.value) || 0;
  const defensePenalties = Number(ui.defensePenalties.value) || 0;
  const parryMode = ui.parryMode.value || "1";
  const parryDivisor = getParryDivisor(parryMode);
  const specialPartState = targetData?.body?.[SPECIAL_PART_NAME] ?? null;
  const specialWasActive =
    hasConfiguredSpecial(targetData) &&
    (Number(specialPartState?.max) || 0) > 0 &&
    (Number(specialPartState?.current) || 0) > 0;
  const targetArmor =
    (Number(targetData?.body?.[targetPart]?.armor) || 0) +
    (specialWasActive ? Number(specialPartState?.armor) || 0 : 0);
  const targetPartState =
    targetData?.body?.[targetPart] ?? { current: 0, max: 0, armor: 0, minor: 0, serious: 0 };
  const beforeHp = targetPartState.current ?? 0;
  const specialBeforeHp = specialWasActive ? Number(specialPartState?.current) || 0 : null;
  const baseTargetParry = targetOdyssey?.skills?.[PARRY_SKILL_NAME] ?? 0;
  const targetParry =
    parryDivisor <= 0
      ? 0
      : Math.max(Math.floor((Number(baseTargetParry) || 0) / parryDivisor), 0);

  const result = resolveAttack({
    attackSkill,
    weaponDamage,
    defenseBonuses,
    defensePenalties,
    attackBonuses: totalAttackBonuses,
    attackPenalties: totalAttackPenalties,
    parry: targetParry,
    targetPart,
    targetArmor,
  });

  const specialResolution =
    result.hit && result.damage
      ? projectDamageWithSpecialProtection({
          specialPart: specialPartState,
          targetPart: targetPartState,
          damage: result.damage,
          targetPartName: targetPart,
        })
      : {
          specialProjectedState: specialPartState ? getNormalizedPartState(specialPartState) : null,
          projectedTargetState: {
            ...getNormalizedPartState(targetPartState),
            critApplied: 0,
          },
          specialActive: false,
          specialArmor: 0,
          damageAppliedLabel: "No Damage",
        };

  const projectedPartState = specialResolution.projectedTargetState;
  const projectedSpecialState = specialResolution.specialProjectedState;
  const afterHp = projectedPartState.current ?? beforeHp;
  const specialAfterHp = specialWasActive ? (projectedSpecialState?.current ?? specialBeforeHp) : null;
  const resolvedAttackSummary =
    specialResolution.specialActive &&
    result.hit &&
    specialResolution.damageAppliedLabel !== "No Damage"
      ? `${result.summary} Applied: ${specialResolution.damageAppliedLabel}.`
      : result.summary;

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
      actorName: sourceName,
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

  if (result.hit) {
    await ensureOverlayForToken(target.id);
  }

  await pushSharedLogEntry(
    `${getAttackOutcomeIcon(result.outcome)} ${sourceName} attacks ${getCharacterName(target)}`,
    formatEnvironmentAttackDebug({
      sourceName,
      targetName: getCharacterName(target),
      targetPart,
      attackSkill,
      weaponDamage,
      weaponAccuracy,
      manualAttackBonuses,
      totalAttackBonuses,
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
      damageAppliedLabel: specialResolution.damageAppliedLabel,
    }),
    result.hit ? "success" : result.outcome === "critical-failure" ? "error" : "info",
  );

  setStatus(
    `${sourceName} -> ${getCharacterName(target)}: ${resolvedAttackSummary}`,
    result.hit ? "success" : "info",
  );
}

async function teardownTargetPickerTool() {
  if (!targetPickState.toolReady) return;

  try {
    await OBR.tool.removeMode(TARGET_PICK_MODE_ID);
  } catch (_error) {
    // ignore
  }

  try {
    await OBR.tool.remove(TARGET_PICK_TOOL_ID);
  } catch (_error) {
    // ignore
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
      } catch (_error) {
        // ignore mode restore errors
      }
    }
  } finally {
    targetPickState.restoring = false;
  }
}

async function stopTargetPick(statusMessage = "", statusKind = "info") {
  const wasActive = targetPickState.active;
  targetPickState.active = false;
  renderTargetState();

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
    icons: [{ icon: EXTENSION_ICON_URL, label: "Pick GM Attack Target" }],
    defaultMode: TARGET_PICK_MODE_ID,
    disabled: { roles: ["PLAYER"] },
  });

  await OBR.tool.createMode({
    id: TARGET_PICK_MODE_ID,
    icons: [{ icon: EXTENSION_ICON_URL, label: "Pick GM Attack Target" }],
    disabled: { roles: ["PLAYER"] },
    cursors: [{ cursor: TARGET_PICK_CURSOR }],
    onToolClick: async (_context, event) => {
      if (!targetPickState.active) return false;

      const clickedTargetId = event.target?.id ?? "";
      const liveItems = clickedTargetId ? await OBR.scene.items.getItems() : [];
      const target =
        (clickedTargetId ? liveItems.find((item) => item.id === clickedTargetId) : null) ?? event.target;

      if (!target || !isCharacterToken(target)) {
        setStatus("Click a visible character token to use it as target.", "error");
        return false;
      }
      if (target.visible === false) {
        setStatus("Hidden tokens cannot be targeted.", "error");
        return false;
      }

      selectedTargetTokenId = target.id;
      renderTargetState();
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
  if (playerRole !== "GM") {
    setStatus("Only the GM can pick targets here.", "error");
    return;
  }

  const visibleTargets = getCharacters();
  if (!visibleTargets.length) {
    setStatus("Add at least one visible character token.", "error");
    return;
  }

  if (targetPickState.active) {
    await stopTargetPick("Target picking cancelled.", "info");
    return;
  }

  targetPickState.previousToolId = await OBR.tool.getActiveTool();
  targetPickState.previousModeId = await OBR.tool.getActiveToolMode();
  targetPickState.active = true;

  await ensureTargetPickerTool();
  await OBR.tool.activateTool(TARGET_PICK_TOOL_ID);
  await OBR.tool.activateMode(TARGET_PICK_TOOL_ID, TARGET_PICK_MODE_ID);
  renderTargetState();
  setStatus("Click a visible character token on the map to assign it.", "info");
}

async function refreshState(showStatus = false) {
  const [role, name, items] = await Promise.all([
    OBR.player.getRole(),
    OBR.player.getName(),
    OBR.scene.items.getItems(),
  ]);
  playerRole = role;
  playerName = name ?? "";
  setSceneItems(items);
  render();
  if (showStatus) {
    setStatus(
      playerRole === "GM"
        ? "GM tools refreshed."
        : "This extension is currently available only to the GM.",
      playerRole === "GM" ? "success" : "error",
    );
  }
}

function bindEvents() {
  ui.refreshBtn.addEventListener("click", () => {
    void refreshState(true).catch((error) => {
      setStatus(error?.message ?? "Refresh failed.", "error");
    });
  });

  ui.publicDiceBtn.addEventListener("click", () => {
    void performPublicGmRoll().catch((error) => {
      setStatus(error?.message ?? "Unable to roll GM dice.", "error");
    });
  });

  ui.privateDiceBtn.addEventListener("click", () => {
    try {
      performPrivateGmRoll();
    } catch (error) {
      setStatus(error?.message ?? "Unable to roll private GM dice.", "error");
    }
  });

  ui.pickTargetBtn.addEventListener("click", () => {
    void startTargetPick().catch((error) => {
      setStatus(error?.message ?? "Unable to start target picking.", "error");
    });
  });

  ui.clearTargetBtn.addEventListener("click", () => {
    selectedTargetTokenId = "";
    renderTargetState();
    setStatus("Target cleared.", "info");
  });

  ui.environmentAttackBtn.addEventListener("click", () => {
    void performEnvironmentAttack().catch((error) => {
      setStatus(error?.message ?? "Unable to resolve environment attack.", "error");
    });
  });
}

OBR.onReady(async () => {
  try {
    bindEvents();
    await refreshState(false);

    OBR.scene.items.onChange((items) => {
      setSceneItems(items);
      render();
    });

    OBR.player.onChange((player) => {
      playerRole = player.role;
      playerName = player.name ?? playerName;
      render();
    });

    render();
    setStatus(
      playerRole === "GM"
        ? "Ready. Use GM Dice or pick a target for environment attacks."
        : "This extension is currently available only to the GM.",
      playerRole === "GM" ? "info" : "error",
    );
  } catch (error) {
    console.error("[Odyssey GM Tools] Init failed", error);
    setStatus(error?.message ?? "Failed to initialize GM tools.", "error");
  }
});
