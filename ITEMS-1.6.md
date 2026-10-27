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
