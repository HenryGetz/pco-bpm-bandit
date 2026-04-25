# PCO BPM Bandit

Tampermonkey userscript for Planning Center plans that builds a MultiTracks BPM/Time Signature table in one click.

## Install

- Direct install URL (Tampermonkey):
  `https://raw.githubusercontent.com/HenryGetz/pco-bpm-bandit/main/pco-multitracks-oneclick.user.js`

Open that link in your browser with Tampermonkey installed and click install.

## Auto-Update Model

The script includes:

- `@updateURL`
- `@downloadURL`

Both point to the raw file on `main`. Tampermonkey checks for newer `@version` values and updates automatically.

## One-Place Update Workflow

1. Edit `pco-multitracks-oneclick.user.js`
2. Bump `@version` (required for client updates)
3. Commit and push to `main`

That is it. All installed users get the update via Tampermonkey background checks.

## Dev Validation

```bash
node --check pco-multitracks-oneclick.user.js
```

