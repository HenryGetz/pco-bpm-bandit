# PCO BPM Bandit

## What This Is

This is a 1-click helper for Planning Center.

On a PCO plan page, it pulls your songs and opens a new page with a table:

- Song
- Version (from MultiTracks)
- Key
- BPM
- Time Signature

This saves worship directors from manually checking every song in MultiTracks.

## 3-Minute Setup (Non-Technical)

1. Install the **Tampermonkey** browser extension.
2. Open this install link:
   `https://raw.githubusercontent.com/HenryGetz/pco-bpm-bandit/main/pco-multitracks-oneclick.user.js`
3. Tampermonkey opens an install screen.
4. Click **Install**.
5. Open any Planning Center plan page.
6. Click the floating **MultiTracks BPM** button.

Done.

## How To Use

1. Go to a PCO plan page.
2. Click **MultiTracks BPM**.
3. Wait a few seconds.
4. A new page opens with the full table.
5. Use **Copy Markdown** if you want to paste the table somewhere.

## Important Notes

- You must be logged in to Planning Center in your browser.
- The tool reads plan data and looks up matching songs on MultiTracks.
- It does **not** edit your PCO plan.

## Auto Updates (Already Built In)

If you installed from the raw GitHub link above, Tampermonkey auto-checks for updates.

When a new version is published, users get it automatically.

## Troubleshooting

- **No button appears**: Refresh the PCO page once.
- **Popup blocked**: Allow popups for Planning Center.
- **Wrong match**: Click the song link in the table to verify MultiTracks result.
