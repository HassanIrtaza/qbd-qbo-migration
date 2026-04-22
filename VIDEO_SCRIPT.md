# Video Recording Script

A customer-facing demo that shows the **full, real** QBD → QBO flow including
the COM extraction, OAuth to QBO, and live migration. ~3 minutes total.

This script assumes you're recording **on the Windows VM** with real
QuickBooks Desktop running. That's what makes the COM extraction visible —
the "Application Certificate" prompt from QBD is the most credibility-building
moment in the whole video.

---

## Before you hit record

### On the Windows VM

- [ ] QuickBooks Desktop is open with a sandbox / sample company file loaded
  (use *Sample Rock Castle Construction* — ships with every QBD install:
  `File → Open Previous Company → Sample Rock Castle Construction, Inc.`)
- [ ] QuickBooks SDK is installed (one-time, see `DEPLOY_WINDOWS.md`)
- [ ] This app is installed and `.venv` already created (first launch takes
  too long to leave in the video — do one warm-up run before recording)
- [ ] `.env` is configured with a **real Intuit developer sandbox** Client ID
  and Secret, so the QBO OAuth flow is authentic on camera
- [ ] QBO sandbox company is **empty** or newly reset so you can show data
  landing in it live
- [ ] Browser: a single fresh window with **only two tabs** —
  - Tab 1: `http://localhost:5050/demo` (our UI)
  - Tab 2: `https://sandbox.qbo.intuit.com` (already signed in to the sandbox)
- [ ] Previous `qbd_exports\` folder emptied (delete any `.xlsx` there so the
  file counter starts at 0/8 on camera)
- [ ] Token file cleared: delete `tokens.json` if present so Connect triggers
  a real OAuth redirect, not a silent reuse

### Recording setup

- [ ] 1080p or 1440p capture (Loom, OBS, or ScreenFlow)
- [ ] System audio **off** (avoids Windows notification sounds)
- [ ] Voiceover: record separately and dub, or narrate live
- [ ] Windows notifications paused (Focus Assist → Alarms Only)
- [ ] Browser in full-screen (F11) — hides bookmarks bar and tabs that
  would give away the demo URL
- [ ] Font scaling on the VM: 125% is a good balance for readability without
  looking cramped
- [ ] QBD window pre-sized to roughly half the screen on the right — you'll
  pop over to it for the permission prompt

---

## Shot list

Timings are target minimums. Pause for narration as needed.

### 🎬 Shot 1 — Establish QuickBooks Desktop (0:00 – 0:15)

**What's on screen:** QuickBooks Desktop foregrounded, sample company file
open, Home screen visible with the flowchart of invoices/bills/accounts.

**What you do:**
1. Click **Customers → Customer Center** — show a list of real customers.
2. Click **Lists → Chart of Accounts** — show the real COA.

**Narration:**
> "Here's a QuickBooks Desktop company with a year of construction data —
> customers, jobs, vendors, and transactions. Moving this to QuickBooks
> Online is normally a multi-day manual export-and-reimport. Let me show
> you how our tool does it end-to-end."

### 🎬 Shot 2 — Launch the app (0:15 – 0:25)

**What's on screen:** Minimize QBD. Open a PowerShell / CMD window in
`C:\Users\...\qbd-qbo-migration`. Type:

```
run_windows.bat
```

**What you do:** Let it boot. The banner and "Open: http://localhost:5050"
line are the money lines — they appear in ~1 second if the venv is warm.

**Narration:**
> "The connector runs as a small web app on the same machine as QBD —
> that's required because QuickBooks' COM interface is local-only."

### 🎬 Shot 3 — Land on the UI (0:25 – 0:35)

**What's on screen:** Browser at `http://localhost:5050/demo`
*(the `/demo` path is cosmetic — the clean customer URL at `/` is identical).*

**What you do:**
- Pause. Let the viewer absorb the layout.
- Optional slow pan: cursor traces the header → left card (QBD) → arrow →
  right card (QBO) → migration panel below.

**Narration:**
> "Two panels: QuickBooks Desktop on the left as the source, QuickBooks
> Online on the right as the target. Nothing's connected yet."

### 🎬 Shot 4 — Connect to QBD (0:35 – 1:05) **★ the money shot ★**

**What's on screen:** Left card's **Connect to QBD** tab is active.

**What you do:**
1. Leave the company file field blank (it uses whichever company is currently
   open in QBD).
2. Click **Connect to QuickBooks Desktop**.
3. **Alt-Tab to QBD** — the "Application Certificate" dialog will have popped
   up. It says something like:
   > *QBD-QBO Migration is attempting to access your QuickBooks company file*

   Pause on this shot for 3 full seconds. This is what proves it's real.
4. Click the **"Yes, whenever this QuickBooks company file is open"** radio,
   click **Continue**, then **Done** on the summary.
5. Alt-Tab back to the browser. The card has flipped green showing the
   `.QBW` file path.

**Narration:**
> "Click Connect, and QuickBooks Desktop itself asks for permission. That's
> the QB SDK — the app can't even read the company file until the QB user
> grants access. Once granted, we're authorized to extract."

### 🎬 Shot 5 — Extract Now (1:05 – 1:40)

**What's on screen:** Back on the browser, Connected state visible.

**What you do:**
1. Click **Extract Now**.
2. Do nothing — let the log console stream. You'll see lines like:
   ```
   Opening QBXMLRP2 session...
   → OpenConnection2 (AppName='QBD-QBO Migration', Mode=1)
   → BeginSession (company=Rock Castle Construction.QBW)
   ✓ Session ticket acquired
   Extracting Chart of Accounts
   Wrote QBD_ChartOfAccounts.xlsx  (84 rows)
   Extracting Customers + Jobs
   Wrote QBD_Customers.xlsx  (132 rows)
   ...
   ```
3. When all 8 entities are done, switch to the **Upload Files** tab briefly
   — the counter now reads **8 of 8 files ready**, filling the list with
   actual row sizes. Switch back to **Connect to QBD**.

**Narration:**
> "The app sends QBXML requests over COM — AccountQueryRq, CustomerQueryRq,
> InvoiceQueryRq — and parses the responses into Excel. Every entity in
> the company file is extracted in dependency order: accounts first, then
> customers and their sub-customer jobs, vendors, items, open AR/AP, and
> the trial balance."

### 🎬 Shot 6 — Prove it's real data (1:40 – 1:55) *(optional but high-impact)*

**What's on screen:** File Explorer on `C:\Users\...\qbd-qbo-migration\qbd_exports\`

**What you do:**
1. Alt-Tab to File Explorer.
2. Show the 8 new `.xlsx` files with fresh timestamps.
3. Double-click `QBD_Customers.xlsx`. Excel opens.
4. Scroll through the rows — **real names, real addresses, real terms.**
   Close Excel.

**Narration:**
> "Every row in those files came straight out of QBD a moment ago — no
> manual export, no re-typing. Customers, jobs, vendors, every transaction."

### 🎬 Shot 7 — Connect to QuickBooks Online (1:55 – 2:20)

**What's on screen:** Back in the browser.

**What you do:**
1. Click **Connect to QuickBooks** in the right card.
2. Browser redirects to **Intuit's real authorize page** — `appcenter.intuit.com`.
   Pause for a second so the Intuit branding reads on camera.
3. Sign in (have credentials pre-saved in the browser so this is one click).
4. Click **Connect** on Intuit's permissions page.
5. Redirected back to our UI. The right card flips green showing the QBO
   company name.

**Narration:**
> "For QuickBooks Online we use Intuit's standard OAuth 2.0 — same flow
> their first-party apps use. One click to authorize, tokens are stored
> locally, we never see the password."

### 🎬 Shot 8 — Dry run migration (2:20 – 2:45)

**What's on screen:** Migration panel at the bottom.

**What you do:**
1. Leave **Dry Run** ticked.
2. Click **Start Migration**.
3. Watch the 4 phase dots light up in sequence: Extract → Transform → Load
   → Validate.
4. Log streams "Creating 84 accounts in QBO (dry run)..." etc.
5. Results panel on the right fills with counts.

**Narration:**
> "A dry run validates everything without writing to QBO. Accounts mapped,
> customers deduplicated, jobs linked to their parents, terms created.
> No warnings, no errors."

### 🎬 Shot 9 — Live migration (2:45 – 3:15)

**What's on screen:** Same migration panel.

**What you do:**
1. **Uncheck Dry Run.**
2. Click **Start Migration**.
3. Phases light up again, this time with real POSTs to Intuit.
4. When complete, switch to **Tab 2 (QBO sandbox)**.
5. Navigate to **Sales → Customers** in QBO — customers that weren't there
   30 seconds ago are now listed with their jobs as sub-customers.
6. Optional: Navigate to **Accounting → Chart of Accounts** to show the COA
   matched.

**Narration:**
> "Live migration. Every POST goes to QuickBooks Online's REST v3 API with
> proper OAuth, rate-limit handling, and retry. And here's the same data
> in QBO — customers, jobs as sub-customers, chart of accounts, open
> balances. End to end in under three minutes."

### 🎬 Shot 10 — Close (3:15 – 3:30)

**What's on screen:** Back to the migration UI, summary panel visible.

**What you do:** Let the frame hold on the results counts: *Accounts 84,
Customers 132, Jobs 54, Vendors 45...*

**Narration:**
> "That's the tool. Works on any QBD Pro, Premier, or Enterprise edition
> from 2019 on, into any QBO Plus or Advanced company. Ready for your
> migration."

---

## Post-production checklist

- [ ] Blur the QB user's email / company name in the Application Certificate
  shot if it's a real client's file
- [ ] Blur the Realm ID in the QBO status panel if using a real production
  company (sandbox is fine to show)
- [ ] If you dubbed audio, align the "Yes, whenever..." click with the
  narration's "grants access"
- [ ] Add a 5-second title card: "QBD → QBO Migration" with the GitHub URL
- [ ] End card: how to reach you / book a call

---

## Things that will happen during recording and are fine

- **QBD goes slightly unresponsive during extraction** (1–3 seconds, while
  QBXMLRP2 is running queries). This is normal. If anything, it reinforces
  that real work is happening.
- **Token refresh mid-migration**. If the session is long enough, you'll see
  a "401 — refreshing token and retrying" in the log. Leave it in — it
  demonstrates OAuth done right.
- **Trial Balance extraction may emit a warning** on some QBD editions. The
  other 7 files still extract cleanly.

## Things that will happen and are NOT fine (cut or retake)

- **QBD asks for permission a second time** mid-video: you picked "Yes, this
  time only" earlier. Fix in QBD Preferences before recording.
- **OAuth state mismatch error** after redirect: the Flask session cookie
  got stale. Restart the app, retake from Shot 7.
- **"Port 5050 already in use"** in the launcher: a previous run didn't
  exit cleanly. Kill old Python process via Task Manager, restart.

---

## One-page shot cheatsheet

| Time | Shot | Key visual |
|---|---|---|
| 0:00 | Establish QBD | Customer Center + Chart of Accounts |
| 0:15 | Launch | `run_windows.bat` banner |
| 0:25 | UI reveal | Two cards, clean header |
| 0:35 | Connect QBD | **QB Application Certificate prompt** |
| 1:05 | Extract Now | Streaming QBXMLRP2 log + 8/8 files |
| 1:40 | Real data | File Explorer → Excel |
| 1:55 | Connect QBO | Intuit authorize page |
| 2:20 | Dry run | 4 phase dots, results panel |
| 2:45 | Live migration | Data appearing in QBO sandbox |
| 3:15 | Close | Results summary hold |
