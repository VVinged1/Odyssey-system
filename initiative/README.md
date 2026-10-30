# Odyssey Initiative

Install this extension separately from Odyssey System using:

https://vvinged1.github.io/Odyssey-system/initiative/manifest.json?v=1.1.0

- Horizontal token images with full names on hover.
- Add selected tokens, roll d20 + effective Odyssey Reaction, and advance the turn.
- GM can remove a participant with the small × on its card; the map token remains.
- Resize using the bottom-right corner; size is remembered in this browser.
- Background context-menu actions are named Add to Odyssey Initiative and Remove from Odyssey Initiative, with normal HTTP(S) icon URLs.
- Rolls are sent to Odyssey Combat Log with the dice calculation.
- Players see visible participants and cannot edit counts or remove tokens.

Existing participants from the standalone Odyssey Initiative are recognized through the original com.codex.odyssey-initiative/metadata key. No migration or reset of their initiative values is required. The extension reads character data from Odyssey System; it does not use Avallon metadata.

Run npm install and npm run build from the repository root. Bundles are published in initiative/assets/. Source attribution and license are in NOTICE.md and LICENSE.
