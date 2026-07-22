# CCSC Mail → Google Calendar

Standalone Google Apps Script automation for:

- reading only CCSC confirmed-reservation emails in
  `23mb1095@student.gs.chiba-u.jp`;
- extracting the reservation ID, location, purpose, title, attendee count,
  notes, and start/end time;
- creating a clearly marked confirmed event directly in
  `23mb1095@student.gs.chiba-u.jp`'s primary Google Calendar;
- adding a clickable link back to the source Gmail thread in the event
  description;
- adding `22m1084c@student.gs.chiba-u.jp` as a guest and sending a Google
  Calendar invitation for every newly created event;
- preventing duplicate events and duplicate invitations by Gmail message ID
  and 貸出ID.

## Behavior

The app matches messages from `byo-hp@system.ho.u-chiba.jp` whose subject
contains `【CCSC】予約申請 確認証および貸出証・利用記録`. Provisional
`仮受付` messages are ignored. A one-minute clock trigger checks Gmail because
Apps Script does not provide an incoming-Gmail trigger.

Created events look like:

```text
【CCSC予約確定】大内自主練｜救急蘇生法
```

New events are inserted through the Calendar API, and the returned API event ID
is retained for reliable future updates. The Calendar location contains the simulation room. The description contains
all extracted fields, the print-and-submit instruction, and a source-mail link.
The guest receives a Calendar invitation because event creation uses
`sendInvites: true`. A 30-minute popup reminder is also added for the organizer;
each guest controls their own event-reminder preferences in Google Calendar.

If a matching legacy `【CCSC仮予約】` event exists for the same 貸出ID, the
confirmation upgrades it in place instead of creating a duplicate. Genuine
confirmation corrections update the event and notify the guest; repeated
identical confirmation mail is suppressed using a reservation fingerprint.

The source-mail link opens the Gmail thread in the organizer's mailbox. Guests
can see the URL in the event description but cannot access the organizer's
mailbox. A one-time automatic migration adds the link to recent events created
before this feature was introduced; migration updates use `sendUpdates: none`
to avoid guest update emails.

The app does not currently process rejection or cancellation mail. Historical
provisional events with no matching confirmation are preserved rather than
deleted automatically.

## Local validation

From this directory:

```sh
npm test
```

From the repository root:

```sh
npm run build
```

## Create and push the Apps Script project

First enable the Apps Script API at
<https://script.google.com/home/usersettings>. Turn **Google Apps Script API**
on. If it was just enabled, Google may take a few minutes to apply the change.

Then make sure `clasp` is logged in with
`23mb1095@student.gs.chiba-u.jp` and run:

```sh
cd projects/ccsc-mail-calendar
npm run create
npm run push
npm run open
```

`npm run create` writes the required `.clasp.json`. If creation fails, do not
run `push` yet: resolve the creation error and rerun `npm run create`.

If a cloud Apps Script project already exists, copy `.clasp.json.example` to
`.clasp.json`, replace the placeholder with that project's script ID, then run
`npm run push`.

The generated manifest enables the Advanced Gmail and Calendar services and
requests read-only Gmail access, Calendar access, and permission to manage this
script's trigger. This intentionally avoids the broader full-mailbox scope
required by the built-in `GmailApp` service.

## One-time activation

In the Apps Script editor:

1. Select `runCcscSelfTests` and click **Run**. Confirm the execution succeeds.
2. Select `previewRecentCcscReservations` and click **Run**. Inspect the
   execution log; no Calendar data is changed.
3. Select `setupCcscAutomation` and click **Run**.
4. Review and accept the requested Google permissions.
5. Check the execution log. Matching recent messages are imported immediately,
   and one one-minute trigger is installed.

`setupCcscAutomation` refuses to run if Gmail is authorized as an account other
than `23mb1095@student.gs.chiba-u.jp`.

## Operations

Functions intended for manual use:

| Function | Effect |
| --- | --- |
| `setupCcscAutomation` | Installs/replaces the trigger and immediately processes recent matching mail |
| `previewRecentCcscReservations` | Read-only parser preview |
| `processCcscReservationEmails` | Runs one import cycle manually |
| `reconcileCcscReservationEmails` | Safely rechecks processed confirmations and recreates any missing events |
| `resetCcscAutomation` | Deletes only tracked CCSC events/state, recreates confirmed events on the pinned university calendar, verifies them, and reinstalls one trigger |
| `getCcscAutomationStatus` | Logs trigger/configuration status |
| `backfillCcscEmailLinks` | Idempotently adds source-Gmail links to recent CCSC events without guest update emails |
| `removeCcscAutomation` | Removes only this app's trigger; existing events remain |
| `runCcscSelfTests` | Tests the email parser without Gmail or Calendar changes |

Configuration is in `src/config.ts`. Leave `calendarId` blank to use the
account's primary calendar.
