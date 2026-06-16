# Working in this repo

## Bump the version on every change to the skill

Any change to the skill — `skills/hands-on-deck/SKILL.md`, `deck.py`, the docs it ships, or its behavior — **must** bump the version in the same PR. Two files carry it and they move together:

- `.claude-plugin/plugin.json` → `version`
- `.claude-plugin/marketplace.json` → `plugins[0].version`

(Leave `marketplace.json`'s top-level `metadata.version` alone — that's the marketplace's own version, not the plugin's.)

Follow semver and the existing commit convention:

- **Patch** (`3.0.1` → `3.0.2`) — guidance tweaks, small fixes, doc changes.
- **Minor** (`3.0.0` → `3.1.0`) — new capabilities or commands.
- **Major** — breaking changes to the CLI or patch format.

Commit message format: `hands-on-deck X.Y.Z: <one-line summary of the change>`.
