import { BODY_ORDER, clamp } from "./shared.js";

export function rollPercent() {
  return Math.floor(Math.random() * 100) + 1;
}

export function rollDice(sides, modifier = 0, count = 1) {
  const safeSides = clamp(Number(sides) || 0, 2, 1000);
  const safeCount = clamp(Number(count) || 0, 1, 100);
  const rolls = Array.from({ length: safeCount }, () => Math.floor(Math.random() * safeSides) + 1);
  const subtotal = rolls.reduce((sum, roll) => sum + roll, 0);
  return {
    roll: rolls[0] ?? 0,
    rolls,
    count: safeCount,
    sides: safeSides,
    subtotal,
    modifier: Number(modifier) || 0,
    total: subtotal + (Number(modifier) || 0),
  };
}

export function calculateAccuracy(
  attackSkill,
  attackBonuses = 0,
  attackPenalties = 0,
  defenseBonuses = 0,
  defensePenalties = 0,
  parry = 0,
) {
  const attackRoll = rollPercent();
  const defenseRoll = rollPercent();

  const attackTotal =
    attackRoll +
    clamp(Number(attackSkill) || 0, 0, 10) * 10 +
    (Number(attackBonuses) || 0) -
    (Number(attackPenalties) || 0);

  const defenseTotal =
    defenseRoll +
    (Number(defenseBonuses) || 0) -
    (Number(defensePenalties) || 0) +
    clamp(Number(parry) || 0, 0, 10) * 10;

  return {
    attackRoll,
    defenseRoll,
    attackTotal,
    defenseTotal,
  };
}

export function calculateDamage(attackResult, defenseResult, weaponDamage = 0, armor = 0) {
  const totalAttack = (Number(attackResult) || 0) + (Number(weaponDamage) || 0);
  const totalDefense = (Number(defenseResult) || 0) + (Number(armor) || 0);
  const damageDiff = totalAttack - totalDefense;

  let label = "No damage.";
  let crit = 0;
  let serious = 0;
  let minor = 0;

  if (damageDiff > 90) {
    label = "Critical damage: 3 Crit.";
    crit = 3;
  } else if (damageDiff > 60) {
    label = "Critical damage: 2 Crit.";
    crit = 2;
  } else if (damageDiff >= 31) {
    label = "Critical damage: 1 Crit.";
    crit = 1;
  } else if (damageDiff >= 6) {
    label = "Serious hit.";
    serious = 1;
  } else if (damageDiff > 0) {
    label = "Minor damage.";
    minor = 1;
  }

  return {
    totalAttack,
    totalDefense,
    damageDiff,
    label,
    crit,
    serious,
    minor,
  };
}

export function resolveAttack({
  attackSkill = 0,
  weaponDamage = 0,
  defenseBonuses = 0,
  defensePenalties = 0,
  attackBonuses = 0,
  attackPenalties = 0,
  parry = 0,
  targetPart = "Torso",
  targetArmor = 0,
}) {
  const part = BODY_ORDER.includes(targetPart) ? targetPart : "Torso";
  const accuracy = calculateAccuracy(
    attackSkill,
    attackBonuses,
    attackPenalties,
    defenseBonuses,
    defensePenalties,
    parry,
  );

  const criticalSuccess = accuracy.attackRoll >= 95;
  const criticalFailure = accuracy.attackRoll <= 5;
  const hit = criticalSuccess || (!criticalFailure && accuracy.attackTotal > accuracy.defenseTotal);

  let outcome = "failure";
  let damage = null;
  let bodyDelta = 0;

  if (criticalSuccess) {
    outcome = "critical-success";
    const baseDamage = calculateDamage(
      accuracy.attackTotal,
      accuracy.defenseTotal,
      weaponDamage,
      targetArmor,
    );
    const crit = Math.max(baseDamage.crit || 0, 2);
    damage = {
      ...baseDamage,
      label: `Critical hit: ${crit} Crit.`,
      crit,
      serious: 0,
      minor: 0,
    };
    bodyDelta = -crit;
  } else if (criticalFailure) {
    outcome = "critical-failure";
  } else if (hit) {
    outcome = "success";
    damage = calculateDamage(accuracy.attackTotal, accuracy.defenseTotal, weaponDamage, targetArmor);
    bodyDelta = -(damage.crit || 0);
  }

  return {
    ...accuracy,
    targetPart: part,
    targetArmor: Number(targetArmor) || 0,
    weaponDamage: Number(weaponDamage) || 0,
    outcome,
    hit,
    damage,
    bodyDelta,
    summary: buildAttackSummary({
      part,
      outcome,
      damage,
      attackRoll: accuracy.attackRoll,
      attackTotal: accuracy.attackTotal,
      defenseTotal: accuracy.defenseTotal,
    }),
  };
}

function buildAttackSummary({ part, outcome, damage, attackRoll, attackTotal, defenseTotal }) {
  if (outcome === "critical-success") {
    return `Critical success to ${part}. Roll ${attackRoll}; ${attackTotal} vs ${defenseTotal}. ${damage?.label ?? ""}`.trim();
  }
  if (outcome === "critical-failure") {
    return `Critical failure. Roll ${attackRoll}.`;
  }
  if (outcome === "success") {
    return `Hit ${part}. ${attackTotal} vs ${defenseTotal}. ${damage?.label ?? ""}`.trim();
  }
  return `Missed ${part}. ${attackTotal} vs ${defenseTotal}.`;
}
