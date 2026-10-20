# fba-prep-sequencer

## Training cards (v2.18)

Photos taken in the app go to a Google Drive folder called **FBA Prep Training** through the Apps Script. One-time setup after pasting the new `Code.gs`:

1. Apps Script editor → pick `setupTraining` in the function dropdown → Run → approve the Drive permission.
2. Deploy → Manage deployments → edit → New version → Deploy.

Without that, cards and steps still sync, but photos stay on the device that took them.
