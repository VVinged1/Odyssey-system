# Checked usable items (1.11.0)

In the GM item editor, Usable items have an optional Use skill text field with character skill suggestions. A new skill name can be entered without adding it to the character. Names are stored as portable labels when possible and retained by token normalization, character templates, pool JSON and account snapshots.

- Empty Use skill: the item applies directly, preserving existing behaviour and making no roll.
- Assigned skill exists on the actor: use the existing opposed V1 skill check with the effective skill level, including passive level modifiers. Base level 0 still counts as a known skill.
- Assigned skill is unknown: after selecting target/body part or inventory armor, choose a characteristic. The existing V1 d20 characteristic check runs first. Only on success does the existing raw d100 vs d100 check run at skill level 0. Failure stops the sequence. A temporary zero-level skill displayed solely because of an item's passive bonus does not count as learned; an active installed implant's granted ability does count as available.

The checks use the item's actor, not the target being healed/repaired/buffed. Existing critical success/failure and tied-roll rules are unchanged. There are no new roll modifiers in this flow.

Cancelling either dialog spends nothing. Invalid/unavailable targets, full-health body parts, undamaged armor, empty stock and the grenade placeholder fail validation before rolling or spending. Once a valid use is confirmed, one unit is spent on either success or check failure. The effect is applied only if every required stage passes. Failed repairs do not split an armor stack; failed stims add no active effect. Self-targeting uses one updated tracker snapshot so applying an effect cannot restore the spent item.

The scene item update resolves the checks and commits the source inventory/history together with the successful target effect. Local tracker writes are queued per actor. Checked outcomes and all stages are saved in the actor's roll history and posted to the public log. A public-log failure does not undo or repeat a completed use; the result remains saved on the character.

Tests use the real V1 roll functions and cover known/missing/temporary skills, base level 0, both critical outcomes, cancellation, each failure stage, full success, self/other targets, repair stacks, stims, placeholder and invalid uses, and portable pool copies. Player browser tests exercise the actual item forms and characteristic picker. This does not introduce new external storage or change attack mechanics.
