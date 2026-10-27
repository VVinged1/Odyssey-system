# Weapon range

GM: Items → weapon pencil → Ranged → Weapon category. New ranged weapons require a category; existing weapons stay Unclassified until edited. The campaign pool groups ranged weapons by category and JSON/templates retain the settings.

Distances are between token centres, calculated by Owlbear's grid measurement rules. Scene scale is converted to metres (metres, feet, inches, yards and kilometres supported). Unrecognised units never silently become metres. A classified attack requires a valid measurement, or a manual distance in No Target Attack; failure spends no ammunition.

| Category | 0–2 m | >2–15 m | >15–50 m | >50 m |
| --- | ---: | ---: | ---: | ---: |
| Pistol | -15 | 0 | -5 | -10…-50 |
| SMG | -25 | +15 | 0 | -10…-50 |
| Shotgun | -25 | +30 | -15 | -50 |
| Assault rifle / machine gun | -30 | 0 | 0 | 0 |
| Sniper rifle | -50 | -30 | 0 | +10…+50 |

Long range progresses by 10 per 25 m: (50,75] → 10; (75,100] → 20; (100,125] → 30; (125,150] → 40; over 150 → 50. Exactly 50 m is medium range.

Range penalty reduction has separate Clinch / Short / Medium / Long values. GM can configure weapon upgrades in the weapon editor, combined perk relief in the skill editor, and passive relief in resource/implant bonus editors. Perks apply only when using that skill; implants must be equipped. Active effects and passive resources follow the existing availability rules. Applicable sources add together, reduce negative range modifiers to a minimum penalty of zero, and never increase a positive range bonus. Manual modifiers remain additional.

The attack panel and combat log show the final range modifier, base penalty and applied relief. Melee and legacy unclassified weapons retain their previous combat behaviour. Environment Attack has no ranged weapon or origin token and keeps its existing manual modifiers.
