# Items and compact character UI (1.6.0)

All character state remains in token metadata. No Supabase connection or backend is required.

- Skills omit empty groups. The GM pencil edits one skill and its perks. Random check uses the existing V1 skill roll with level zero (d100 vs d100).
- Items contains resources, usables, armor, implants, weapons, ammunition and physical magazines. The typed editor updates the existing V1 weapon/ammunition collections.
- Medkits restore body-part critical HP only. Repair kits restore equipped armor durability only. A target defaults to self; choosing another visible character is supported. Grenades are placeholders.
- Stims add attribute modifiers as active effects. The GM removes the effect chip to remove its modifiers. Resource bonuses apply while carried; equipment bonuses and implant abilities apply while equipped. Skill bonuses modify skill levels, retaining the V1 0–10 range.
- Implant critical capacity increases the equipped body part's maximum HP. It is not separate durability. Equipping never heals; medkits can restore the increased HP pool. Removing the implant restores the base maximum and clamps current HP. An implant can require GM removal.
- Equipped armor has critical durability. Critical damage reaching the part consumes that durability first; overflow reaches body HP. V1 minor/serious promotion and Special protection remain in effect. Broken armor stops providing protection until repaired.
- Armor class counts equipped armor/shields only, excluding implant and base body armor: light 1–70, medium 71–140, heavy 141–260, superheavy 261+. Untrained attribute penalties are light AGI/END −1; medium STR/AGI/END −2; heavy/superheavy STR/AGI −4, END −5. Level 1 in the matching passive armor skill removes these penalties.
- Parry defaults to Do Not Count Parry. Character choices persist on the token; GM environment choices persist in the browser.
- Character rosters are in GM Tools. Action buttons use compact icons with accessible names and hover tooltips.

Validation: npm test (ammunition suite and 16 integration checks); npm run test:ui (mock SDK browser scenarios including owner permissions, healing another player's target, stim removal, implant HP, skills, magazines, templates, HUD, compact buttons and narrow layout); npm run build. Live Owlbear multi-client behavior still requires a room test.

## 1.6.1

Passive bonus editors can add missing skills. Such skills exist while their granting item is active and use V1 level-based rolls. Attack, No Target Attack and GM Environment Attack now use the compact three-card layout. The duplicate inline HUD is removed; the floating HUD is available from the character header.

Ranged weapons choose detachable magazines or a built-in magazine. Detachable weapons store compatible magazine IDs and take capacity/ammunition from the mounted magazine. Built-in weapons store compatible ammo, capacity and rounds directly; no physical magazine is synthesized. Legacy weapons retain their previous magazine state. Switching magazine type requires emptying or detaching loaded rounds first.

Validation: 19 integration checks, ammunition suite and browser tests cover both magazine modes, missing-skill rolls, targeted/manual/environment attacks and layout.

## 1.6.2

Players see only armor training skills at level one. Equipment slot checkboxes list allowed locations; equipping still selects a single body part. Each part permits one armor/shield item and multiple implants. The implant player permission now governs both installation and removal. Templates are GM-only and share the Ownership Import / Export area.

Armor durability is included in body-part HP/Max: maximum durability increases Max, remaining durability increases HP. Damage is retained on the armor when removed and restored to the display when re-equipped. Critical hits consume armor HP first; medkits heal only body HP and repair kits heal only armor HP. Manual HP edits and bridge HP deltas use the same pool. Existing V3 tokens migrate without repeatedly adding durability.

Validation: ammunition suite, 23 integration checks and browser tests, including 100 implants on one limb, armor slot conflicts, migration, template health, healing/repair separation, player permissions and combined HP display.

## 1.7.0

The GM Campaign item pool lives in Owlbear room metadata and persists across scenes and characters. Creating/editing items automatically saves reusable definitions; opening existing scenes also captures their items. Give copies from Items → Campaign item pool. Copies are unequipped and undamaged, and physical magazines are empty. Required ammunition definitions and compatible empty magazines are remapped to the destination when missing. Skill bonuses bind to destination skill labels. JSON import/export moves catalog definitions between rooms.

Catalog entries use independent SHA-256 metadata keys so different definitions from concurrent GMs do not overwrite each other. Removing a definition leaves a tombstone, preventing scene scans from immediately restoring it; explicit creation or import can restore it.

Equipping one unit from a gear stack separates that unit from the spare quantity, preserving its individual damage and bonuses. Creating multiple magazines creates distinct IDs and independent round counts. Numeric attack fields are 72px wide in character and GM attack cards.

Validation: ammunition suite, 23 character integration checks, 7 item-pool/instance checks, and browser scenarios for scene changes, new-character issuance, JSON catalog restoration, independent quantities and compact field widths.

Magazine compatibility follows the selected magazine design (name and capacity), so all identical physical copies fit the weapon while retaining separate IDs and round counts.
