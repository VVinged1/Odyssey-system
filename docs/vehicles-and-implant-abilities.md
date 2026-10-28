# Transport, implants and grouped items

## Transport

GM marks an ordinary character token as Vehicle using its Vehicle strip. This keeps its body, inventory, weapons, magazines and ammunition intact. Enter vehicle opens a driver/pilot selector; a player can select a character they own, and GM can select any character. One driver per vehicle, one vehicle per driver; nested vehicles are rejected.

While occupied, displayed characteristics and skills, characteristic/skill checks, attack skills, Parry and range perk reductions use the driver's current effective values. Attacks originate from the vehicle, consume its ammunition and record its history. Item use consumes vehicle stock while using the driver's skill checks. Ownership is not reassigned: the driver gains extension controls while seated and loses that additional access on exit. GM can always exit a vehicle. Missing drivers block rolls until GM clears or replaces the relationship. Both tokens must remain in the same scene.

This is a metadata relationship, not automatic map movement, hiding, or an Owlbear Mount operation. No token positions or Owlbear native movement permissions are changed. Exit preserves damage, items, ammunition and both original character sheets. Edit driver statistics on the driver's own token. Vehicle templates preserve the Vehicle designation but omit the occupant.

## Implant skills and abilities

Resources show an absent skill at zero. An equipped implant can grant an absent skill with its bonus level, including for consumable use requirements. Skill names match without case or surrounding whitespace. Direct character skills retain their level, category, descriptions and perks; additional sources resolve onto that skill. Multiple granted abilities with the same name share one displayed skill; the first implant ability supplies its base level. Numeric bonuses from eligible active sources stack and levels retain the existing 0–10 limit.

Resource and implant passive skill bonuses have Combat/Applied category selectors, including for newly named skills. A character's own category always wins. Categories and skill references survive pool JSON copies and character templates.

Implants support existing Granted skill abilities and two additional kinds:

- Activate bonuses: signed characteristic and skill bonuses create a named active effect. Repeated activation cannot stack the same ability. GM can remove the effect; removing the implant clears effects from it. There is no invented duration or resource cost.
- Attack: name, description, damage, accuracy and assigned skill. Prepare attack selects the ability in existing Attack/No Target Attack forms. Selecting its weapon entry also selects its skill. The existing attack, Parry, armor and damage resolver applies; damage and skill come from the equipped ability. The assigned skill must exist, either directly or through an equipped implant. This ability consumes no ammunition; ammunition-based attacks remain ordinary weapons. Unequipping the implant removes its attack entries.

## Stacks and hover descriptions

Identical unequipped ordinary items combine quantities. Equipped equipment, ammunition, weapons and physical magazines retain their individual mechanics. Same-named items with different properties remain distinct variants within an expandable inventory/pool name group, so wear, bonuses and descriptions are not lost. Quantities are bounded; overflow is retained as another stack.

Description icons and compact buttons show an opacity-faded hover/focus tooltip. Clicking an info icon still opens a description dialog for touch and keyboard use.

## Validation

Pure regressions cover absent implant vs resource skills, category precedence, duplicate skill labels, activation/removal, attack definitions, safe stacks, portable pool categories and driver isolation. Browser regressions exercise actual item/ability editors, activation, assigned attack skills, hover descriptions, inventory and pool groups, driver permissions, driver characteristic rolls, and exit. Existing combat, ammo, pool, templates, concurrency and checked-item suites remain enabled.
