# Odyssey Combat Log

Standalone Owlbear Rodeo extension for viewing the shared Odyssey System combat log.

Install manifest after GitHub Pages is active:

- `https://limonety995-maker.github.io/odissey-dicelog/manifest.json`

## What it does

- Opens as a separate Owlbear extension
- Reads the room-wide combat log from shared room metadata
- Listens for live broadcast updates from the main Odyssey System extension
- Shows the latest attacks, rolls, and debug combat tables in a dedicated viewer

## Shared integration

This extension currently listens to the same shared keys used by the main Odyssey System project:

- Metadata key: `com.codex.body-hp/debugLog`
- Broadcast channel: `com.codex.body-hp/debug`

That means it can work immediately with the current main extension without changing the combat logic first.

## Local build

```bash
npm install
npm run build
```
