# Odyssey UI integration (feature/odyssey-ui-integration)

Base: TARGET `2737c0dea4de50cb8950f291f9e7991bc8803870`. SOURCE is a read-only visual reference at `d08ef1d606984cf05e4f43404556336c3db15f24`.

## Included

- V2-style 3×3 characteristic cards and skill cards with a full 0–10 scale.
- Assigned players can view and roll their character; GM-only value/skill editors remain separate. Old forms and cards call the same V1 checks.
- Compact weapon/target/action HUD in the console and an optional floating popover, using the existing attack command and target picker.
- Token metadata inventory: generic/consumable/medical/resource/tool items, armor, shields and implants; quantity/description editing; equip/unequip and use-one actions.
- Individual magazines: legacy conversion, detached reserve magazines, fill/empty, mount/swap, compatibility and exclusivity checks.
- Character templates: save, rename, duplicate, remove, apply, JSON import/export, browser-local library and explicit save/load through the Owlbear Assets scene picker.
- GM presentation refresh, restored private-results container and bounded visibility checkbox hit area.

## Compatibility and storage

The tracker still lives at `com.codex.body-hp/data`. New data is under `odyssey.inventory` (version 2). The existing `body`, skills/attributes, owner/identity, weapons/ammunition, history and bridge remain in use. `odyssey_rules.js` and `ammunition.js` are unchanged.

Mounted magazines mirror the V1 weapon `capacity/loaded/loadedAmmoId` fields during normalization. V1 attacks and reloads remain authoritative while mounted; detached magazines keep their own rounds. Removing/converting a weapon detaches its magazine instead of returning the same rounds twice. Removing ammunition is blocked while a loaded magazine uses it. New and legacy weapons gain stable magazine links without spending ammunition. Reading alone does not write the migration to Owlbear; the next normal token mutation persists it.

Armor equip replaces the selected body's armor value and remembers its previous value; it does not add armor a second time. Unequip restores the previous value unless a GM has manually changed armor afterwards. One equipment item per part is supported in this iteration. Use-one only decrements quantity: it does not introduce SOURCE healing, effect/cooldown/encounter rules. Healing and damage continue to use V1 tools.

Templates save healthy body state and configuration, including ammunition/inventory. They exclude the source owner, external identity, target, roll history, bridge cursors and map properties. Applying a template reissues weapon/ammunition/magazine/item IDs and fixes references. Destination owner/identity/history are preserved; saved attack target is cleared. A confirmation lists replaced sections before application. Applying templates is GM-only.

The local library uses browser storage, not an account-wide database. JSON export is the portable backup. Account save packs template JSON into metadata of text items in a scene named **Odyssey Templates**, then calls SDK 3.1.0 `assets.uploadScenes`. Load uses `assets.downloadScenes` and merges matching library payloads. This is an explicit snapshot, not automatic cloud synchronization or updating an existing asset in place. Cancelling the load picker leaves the local library unchanged. The SDK returns no upload-success/cancellation distinction; the UI asks the user to verify the saved asset.

All active main/GM tabs must be refreshed together after deploying this version so an old bundle's sanitizer cannot drop the new inventory fields. No hosting/backend deployment is included in this branch. Supabase is neither installed nor imported.

## Build and checks

```sh
npm ci
npm test
npm run build
npx playwright install chromium
npm run test:ui
```

`npm run build` rebuilds main/background, GM and room-log. Rebuilding GM is required because it also normalizes token metadata. `npm test` runs the existing ammo/attack checks and integration checks for data migration, conservation, template isolation, permissions, storage and V1 roll boundaries. The pre-existing ammo test fixture lacked the current `isEditable/getBodyPartAttackPenalty` stubs; those test boundaries are repaired without changing production attack math.

The Playwright suite bundles the actual main and GM code against `scripts/fixtures/mockSdk.js`. It checks GM/owner/outsider access, characteristic/skill rolls, a real UI attack with magazine synchronization, equipment, template/account payload round-trip, cancellation, floating HUD, private/public GM results and narrow layout. Production builds do not include this SDK fixture. Optional `PLAYWRIGHT_CHANNEL=msedge` selects installed Edge; by default Playwright Chromium is used. Set `ODYSSEY_SCREENSHOT_DIR` to save previews.

## Live Owlbear acceptance remaining

Automated tests mock the SDK boundary. Before deploying, test with GM and player in a disposable real room: permissions, target-tool lifecycle and floating popover; save the Assets library, then load it from another browser/account session and verify metadata retention. Live account storage and cross-client races are not claimed as tested by the fixture.

Manifest URLs still point to TARGET hosting, so opening a feature branch file alone does not serve the branch's code in Owlbear. Use a separate HTTPS preview host with its own manifest URLs to test without changing main hosting.
