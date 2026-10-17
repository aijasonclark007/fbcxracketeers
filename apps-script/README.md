# Google Apps Script

Two files that live in the same Apps Script project (the one attached to the registration Google Sheet).
They share one global scope, so keep both in the project.

| File | What it does |
|---|---|
| `Registration.gs` | Web app: saves registrations, decides the waitlist, serves the slot counts, saves payment screenshots |
| `EmailSystem.gs` | Sends the emails when you change **Status** (column F): Approved, Declined, Move Up, Remind, Waitlist |

## Waitlist
- Slots per level are set in `SLOT_CAP` (High 24, Mid 32, Low 24). The website reads them from the script.
- A level is full when its **Approved** (and **Remind**) rows reach the cap.
- A new registration to a full level is saved with Status **Waitlist** and gets a "you're on the waitlist" email automatically.
- When a slot opens, set that row's Status to **Approved**. The player gets the normal approval email (with the payment link) saying a slot opened up.

## Alert email
The organizer alert address is not in the code (this repo is public). In Apps Script: Project Settings > Script properties > add `NOTIFY_EMAIL`.
If it is not set, alerts go to the account that owns the script.

## Updating the live script
1. Paste both files over the old ones and save.
2. Run `setupStatusColumn()` once (adds "Waitlist" to the Status dropdown).
3. Deploy > Manage deployments > pencil icon > Version: **New version** > Deploy. (Same web app URL, so the site keeps working.)
