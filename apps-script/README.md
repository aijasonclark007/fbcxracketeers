# Google Apps Script

Two files that live in the same Apps Script project (the one attached to the registration Google Sheet).
They share one global scope, so keep both in the project.

| File | What it does |
|---|---|
| `Registration.gs` | Web app: saves registrations, decides the waitlist, serves the slot counts, saves payment screenshots |
| `EmailSystem.gs` | Sends the emails when you change **Status** (column F): Approved, Declined, Move Up, Remind, Waitlist |

## Waitlist
- Slots per level are set in `SLOT_CAP` (High 16, Mid 32, Low 24). The website reads them from the script.
- A level is full when its **Approved** (and **Remind**) rows reach the cap.
- A new registration to a full level is saved with Status **Waitlist** and gets a "you're on the waitlist" email automatically.
- When a slot opens, set that row's Status to **Approved**. The player gets the normal approval email (with the payment link) saying a slot opened up.

## Tournament Jersey orders
- The jersey page (`jersey.html`) saves each order to a **TournamentJersey** tab in the same Google Sheet. The tab is created automatically on the first order, or right away if you run `setupJerseySheet()` once from the Apps Script editor.
- Columns: Timestamp, Full Name, Back Print Name (always capitals, up to 12 characters), Size, Email, Status, Emailed, Token, Payment Proof, Paid At.
- Orders close at the end of November 1, 2026 (Philippine time). Change `JERSEY_CLOSES` in `Registration.gs` to move the date, and the date in `jersey.html`.
- The price is `JERSEY_PRICE` in `Registration.gs` (the payment page shows 400 in `payment.html`).
- Set **Status** in the TournamentJersey tab: **Send payment** emails the payment link (the existing payment page with `?k=j`), **Remind** nudges anyone who has not uploaded a screenshot, **Paid** and **Cancelled** send nothing.
- The same name with the same email cannot order twice. You get an email for every new order, like new registrations.
- After pasting both files, deploy a **new version** (see below). No other setup is needed.

## Alert email
The organizer alert address is not in the code (this repo is public). In Apps Script: Project Settings > Script properties > add `NOTIFY_EMAIL`.
If it is not set, alerts go to the account that owns the script.

## Updating the live script
1. Paste both files over the old ones and save.
2. Run `setupStatusColumn()` once (adds "Waitlist" to the Status dropdown).
3. Deploy > Manage deployments > pencil icon > Version: **New version** > Deploy. (Same web app URL, so the site keeps working.)

## Payment alerts
- Whenever a player or jersey buyer uploads a payment screenshot, you get an email (to the same address as new-registration alerts) with the screenshot link and the sheet link, so you can check it and set Status.
