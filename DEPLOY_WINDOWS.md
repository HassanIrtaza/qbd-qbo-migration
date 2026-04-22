# Windows Deployment

Running the app **on the same Windows VM as QuickBooks Desktop** — this is the
production setup, and the only way to get real COM extraction (QBXMLRP2 is
local-only).

## Prerequisites

Install once on the VM:

1. **QuickBooks Desktop** (Pro, Premier, Enterprise — 2019 or newer)
2. **QuickBooks SDK** — free from Intuit:
   https://developer.intuit.com/app/developer/qbdesktop/docs/get-started/download-and-install-the-sdk
   (This is what provides the `QBXMLRP2.RequestProcessor` COM class.)
3. **Python 3.9+** — https://www.python.org/downloads/
   During install, tick **"Add python.exe to PATH"**.
4. **Git** (optional, only if you want to `git clone` — otherwise download the
   repo as a ZIP): https://git-scm.com/download/win

## Install the app

Open **PowerShell** or **Command Prompt** on the VM:

```powershell
cd C:\Users\%USERNAME%\Documents
git clone https://github.com/HassanIrtaza/qbd-qbo-migration.git
cd qbd-qbo-migration
```

Or download the ZIP from GitHub and unzip to
`C:\Users\<you>\Documents\qbd-qbo-migration`.

## First launch

1. Open QuickBooks Desktop and open the company file you want to migrate.
   **Leave QBD running.**
2. Double-click **`run_windows.bat`** in File Explorer.
   - First run creates a `.venv` and installs dependencies (~60 seconds).
   - Subsequent runs just boot the server.
3. A browser window won't open automatically — go to
   **http://localhost:5050** in Edge/Chrome.
4. In the Source card, click **"Connect to QuickBooks Desktop"**.
5. **IMPORTANT:** QBD will pop up an "Application Certificate" prompt the
   first time. Choose **"Yes, whenever this QuickBooks company file is open"**
   and click **Continue**. This is a one-time permission grant — QBD
   remembers the app after this.
6. The card flips green showing the company file path. Click **"Extract Now"**.
7. The log console streams QBXMLRP2 calls in real time. Excel files land in
   `qbd_exports\` inside the project folder.

## Configure QBO credentials (for real migrations)

On first run without a `.env` file, the app runs in simulation mode. To hit
a real QuickBooks Online sandbox or production company:

1. Go to https://developer.intuit.com/app/developer/myapps
2. Create a new app → "QuickBooks Online and Payments"
3. Under the app's **Keys & Credentials**, copy the Client ID and Client Secret
   for the **Development** (sandbox) environment.
4. Add `http://localhost:5050/callback` to the app's **Redirect URIs**.
5. In the project folder, copy `.env.example` to `.env` and fill in:
   ```
   QBO_CLIENT_ID=ABcdefg123...
   QBO_CLIENT_SECRET=xyz789...
   QBO_REDIRECT_URI=http://localhost:5050/callback
   QBO_ENVIRONMENT=sandbox
   ```
6. Stop the app (Ctrl+C) and re-run `run_windows.bat`.
7. Click **"Connect to QuickBooks"** in the right card — you'll be redirected
   to Intuit's real authorize page.

## Common issues

**"QBXMLRP2.RequestProcessor class not registered"**
→ The QuickBooks SDK isn't installed. Install it from the Intuit link above,
reboot, try again.

**"The application wants to connect but QuickBooks is not running"**
→ Open QuickBooks Desktop and the company file first, then click Connect.

**QB prompts for app permission every extraction**
→ You said "Yes, this time only" previously. Run Edit → Preferences →
Integrated Applications → Company Preferences, find the app, and set access
to "Allow this application to read and modify this company file".

**pywin32 post-install didn't run**
→ Run as admin: `.venv\Scripts\python.exe Scripts\pywin32_postinstall.py -install`

**Port 5050 is already in use**
→ Set a different port before launching:
```
set PORT=5060
run_windows.bat
```

## Stopping

Close the console window, or press **Ctrl+C** in it. The app does not install
as a Windows service — it only runs while the console window is open. This
is intentional for migration work (you want a human watching).

## Firewall

If you want to reach the UI from your laptop on the same LAN (useful for
presenting while sitting next to the VM), open inbound TCP 5050 in Windows
Defender Firewall and browse to `http://<vm-ip>:5050`. The app binds to
`0.0.0.0` already.
