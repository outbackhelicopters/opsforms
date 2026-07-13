# Setting Up the Backend — Step by Step
### Outback Helicopter Airwork NT — Flight Paperwork App

When a pilot taps **Send**, the app automatically:
- ✉️ Emails the paperwork to your ops inbox (sent from your @outbackhelicopters.com.au address via Office 365)
- 📁 Files the PDF into OneDrive (organised by aircraft and month)

**Time needed:** About 20–30 minutes, done once.

---

## Before you start

Make sure you have:
- Access to your **GitHub account** (where the app code lives)
- Access to your **DigitalOcean account** (where the app is deployed)
- Access to your **Microsoft 365 admin account** (needed to set up the app registration)

---

## Part 1 — Push updated code to GitHub

1. Go to **github.com/outbackhelicopters/opsforms**
2. Click **Add file → Upload files**
3. Drag in the changed files (or use "choose your files")
4. Type a short summary of the change and click **Commit changes**

DigitalOcean starts redeploying automatically — about 2 minutes.

---

## Part 2 — Set up the Microsoft 365 app registration

This creates a "service account" so the app can send emails and file to OneDrive automatically — without anyone needing to be logged in.

You only do this once.

### 2a — Register an app in Azure

1. Go to **portal.azure.com** in your browser
   - Sign in with your Microsoft 365 admin account

2. In the search bar at the top, type **App registrations** and click it

3. Click **+ New registration**
   - **Name:** `Heli App`
   - **Supported account types:** leave on the first option (Single tenant)
   - Leave Redirect URI blank
   - Click **Register**

4. You're now on the app's overview page. Copy and save these two values:
   - **Application (client) ID** — looks like: `a1b2c3d4-1234-5678-abcd-...`
   - **Directory (tenant) ID** — looks like: `e5f6g7h8-5678-abcd-...`

### 2b — Create a client secret

5. In the left sidebar, click **Certificates & secrets**

6. Click **+ New client secret**
   - **Description:** `heli-app`
   - **Expires:** 24 months
   - Click **Add**

7. A secret appears in the table. Copy the **Value** column immediately
   - ⚠️ It disappears permanently after you leave this page

### 2c — Grant permissions

8. In the left sidebar, click **API permissions**

9. Click **+ Add a permission** → click **Microsoft Graph** → click **Application permissions**

10. Add these two permissions (search for each and tick the checkbox):
    - **Mail.Send** — so the app can send email from your address
    - **Files.ReadWrite.All** — so the app can file PDFs to OneDrive

11. Click **Add permissions**

12. You'll see a yellow warning. Click **"Grant admin consent for [your organisation]"** → click **Yes**
    - Both permissions should now show a green tick ✅

> ✅ You now have three things saved: Tenant ID, Client ID, Client Secret.

---

## Part 2d — Set up Twilio for SMS reminders

The Job Calendar texts a pilot 1 hour before their job starts, reminding them to complete
their SWMS and flight briefing. This needs a Twilio account (a paid SMS provider — a few
cents per text).

1. Go to **twilio.com** and sign up (or sign in if you already have an account)

2. On the Twilio Console dashboard, copy and save these two values:
   - **Account SID** — starts with `AC...`
   - **Auth Token** — click "show" to reveal it

3. Buy a phone number that can send SMS:
   - In the left sidebar, **Phone Numbers → Manage → Buy a number**
   - Choose an Australian number (or your own country) with **SMS** capability
   - Buy it, then copy the number in the format `+61...`

> ✅ You now have three things saved: Account SID, Auth Token, and a Twilio phone number.

4. Pilot phone numbers, aircraft and the client list are no longer edited by changing
   `api/config.json` in GitHub. The very first time this app boots with real Microsoft 365
   credentials, it copies whatever is in `api/config.json` into that company's own OneDrive
   (`_system/config.json`) and uses the OneDrive copy from then on — editable with no
   redeploy, and never shared with any other company's deployment. Any format is fine for
   phone numbers (`0412 345 678`, `+61412345678`, etc.) — it's converted automatically.
   Reminders can't be sent to a pilot with a blank phone number.

   Until the admin app's Pilots/Fleet/Clients editors are built (next phase), updates go
   through the API directly, e.g. from a terminal:
   ```
   curl -X PUT https://your-app-name.ondigitalocean.app/api/setup/config \
     -H "Content-Type: application/json" \
     --cookie "adm_sess=<paste from browser dev tools after logging into /admin.html>" \
     -d '{"pilots":[{"name":"Jane Pilot","phone":"0412345678","email":"jane@example.com"}]}'
   ```
   Only signed-in provider/admin accounts (or anyone, before the very first admin account
   is created) can call this.

---

## Part 3 — Add credentials to DigitalOcean

1. Go to **cloud.digitalocean.com**

2. Click **Apps** in the left sidebar → click **outbackheli-app**

3. Click the **api** component → click **Settings** → scroll to **Environment Variables**

4. Click **Edit** and add each variable below. Tick **Encrypt** for each one:

   | Variable Name    | Value                                                          |
   |------------------|----------------------------------------------------------------|
   | `SENDER_EMAIL`   | The @outbackhelicopters.com.au address to send from (and whose OneDrive to file into) |
   | `OPS_EMAIL`      | The email address where completed paperwork should be delivered |
   | `MS_TENANT_ID`   | Directory (tenant) ID from Part 2a                            |
   | `MS_CLIENT_ID`   | Application (client) ID from Part 2a                          |
   | `MS_CLIENT_SECRET` | Client secret Value from Part 2b                            |
   | `TWILIO_ACCOUNT_SID` | Account SID from Part 2d                                    |
   | `TWILIO_AUTH_TOKEN`  | Auth Token from Part 2d                                     |
   | `TWILIO_FROM_NUMBER` | The Twilio phone number from Part 2d, e.g. `+61...`         |

   The `ONEDRIVE_FOLDER` is already set to `Helicopter Paperwork` — you can leave it as is, or change it to whatever folder name you want.

   **For the office admin app**, also add:

   | Variable Name    | Value                                                          |
   |------------------|----------------------------------------------------------------|
   | `SESSION_SECRET` | Any long random string (40+ characters) — signs the office sign-in sessions |

5. Click **Save** — DigitalOcean will redeploy with the new settings

---

## Part 3b — First sign-in to the office admin app

The office team's desktop app lives at `https://your-app.ondigitalocean.app/admin.html`.

1. Open `/admin.html` in a browser. The first time, it runs a **setup wizard**:
   - **Step 1 — System:** checks the Microsoft connection, OneDrive access,
     email settings, and shows what's still missing (Twilio and the device
     token are optional at this stage).
   - **Step 2 — Company:** company names, header lines, ABN, location, brand
     colours and logo — this brands the pilot app, this office app, the PDFs,
     emails and SMS in one go. All editable later in Settings → Branding.
   - **Step 3 — Account:** create the first account (the *provider* login: yours).
2. Go to **Settings → Invite user** to add office staff. Each person gets an
   emailed link to set their own password.
   - **Admin** — everything, including users (for the chief pilot / owner)
   - **Office** — jobs, paperwork, clients (day-to-day scheduling)
3. User accounts are stored in OneDrive under `Helicopter Paperwork/_system/` —
   no extra database needed.

The old `/reports` page and `REPORTS_PASSWORD` still work, but the admin app
replaces them — once everyone's on `/admin` you can remove `REPORTS_PASSWORD`.

---

## Part 4 — Test it works

1. Wait about 2 minutes for the deploy to finish (you'll see "Deployed" turn green)

2. Open a new browser tab and go to:
   ```
   https://your-app-name.ondigitalocean.app/api/health
   ```
   Replace `your-app-name` with your actual DO app URL.
   
   You should see: `{"ok":true,"ts":"2026-..."}`
   If yes — the backend is running ✅

3. On the iPad, complete a test flight all the way through and tap **Send**

4. Within about 30 seconds you should see:
   - An email arrive at your ops inbox, sent from your @outbackhelicopters.com.au address, with the PDF attached
   - The PDF appear in OneDrive under `Helicopter Paperwork / VH-XXX / 2026-06 /`

---

## Troubleshooting

**Email not sending**
→ Check that **Mail.Send** permission was granted in Part 2c (green tick must be showing for both permissions).
→ Make sure `SENDER_EMAIL` is the full email address, e.g. `ops@outbackhelicopters.com.au`

**OneDrive not receiving files**
→ Check that **Files.ReadWrite.All** was granted. The green tick must be visible for both permissions.
→ Make sure admin consent was granted (Part 2c step 12) — not just added, but consented.

**"Not authorised" error in DO logs**
→ The tenant ID, client ID, or client secret may have a typo. Re-enter them carefully — copy/paste is safest.

**SWMS/briefing reminder text not arriving**
→ Check `TWILIO_ACCOUNT_SID`, `TWILIO_AUTH_TOKEN` and `TWILIO_FROM_NUMBER` are all set in DO.
→ Check the pilot has a phone number on file (Part 2d, step 4) — any format works, it's converted automatically.
→ Trial Twilio accounts can only text numbers you've verified in the Twilio Console — upgrade to a paid account to text any number.

**Everything still showing old version on the app**
→ On the iPad: Settings → Safari → Clear History and Website Data → reopen the app.

---

## Part 5 — Device token (locks the API to your iPads)

The server endpoints the iPads use (config, sending, calendar) can be locked
so only your devices and signed-in office users can reach them.

1. Make up a token — any random string of 15+ characters (e.g. from a password generator)
2. On **each iPad**: Settings (PIN) → Security & data → **Device token** → paste it → Save settings
3. Only after every iPad has it: DigitalOcean → Apps → your app → Settings →
   **Environment Variables** → add `DEVICE_TOKEN` = the same string (tick Encrypt) → Save

Order matters: while `DEVICE_TOKEN` is unset on the server nothing is enforced,
so the iPads keep working while you go around entering it. Enforcement starts
the moment the variable is saved.

---

## What happens once it's all set up

- Pilots work completely offline in the field — no internet needed
- When they tap Send, if there's no internet it saves locally and waits
- The moment the iPad gets any connection it sends automatically
- The top bar shows: **orange** = waiting to send, **green** = all sent
- You never need to touch any of this again — it runs automatically

---

## Setting up a second customer

This app is built so every customer runs the exact same code, with nothing
customer-specific meant to live in the repo — see `PLAN-admin-and-commercial.md`
for the full model. Short version: repeat Parts 1–5 above as a brand new
DigitalOcean app pointed at the same GitHub repo, with that customer's own
Microsoft 365 app registration (their data stays in their own tenant) and
their own Twilio subaccount.

**The onboarding flow, end to end:**
1. Stand up their DigitalOcean app + Microsoft 365 + Twilio subaccount (Parts 1–3).
2. Open `/admin.html` yourself once and complete the setup wizard to claim the
   **provider** account (your team's login — exists on every deployment, see
   `PLAN-admin-and-commercial.md`). You can skip past pilots/aircraft here —
   that's the customer's own step next.
3. Go to Settings → Invite user, enter the customer's actual admin (chief
   pilot / owner), role **admin**. This sends them a welcome/setup email with
   a sign-in link — the existing invite system in `auth.js` already does this,
   nothing new to configure.
4. They click the link, set a password, and land straight back in the same
   Company + Pilots & Fleet wizard (branding, logo, pilots, aircraft with
   full weight & balance and accessories, clients) — because the deployment's
   config is still empty, signing in routes them there automatically instead
   of the normal dashboard. They fill in all their own info; nothing from
   OHANT or any other customer is ever visible to them.

**What's already customer-specific and safe to clone as-is:** pilots,
aircraft, W&B, clients and branding — all pulled from that customer's own
OneDrive (empty until their admin fills the wizard in; a brand new iPad
seeds itself from that same server config the first time it boots, so
nothing is hardcoded into the pilot app either).

**What still needs a manual per-customer swap, because these are static
files a browser reads before any app code runs:** `manifest.json`
(`name`/`short_name`/`description`), `icon-192.png`/`icon-512.png`/`logo.png`,
and the `apple-mobile-web-app-title` meta tag in `flight-ops.html` and
`admin.html` — these control the home-screen icon name/artwork on first
install and can't be set from a config file. Swap them before sending a new
customer their install link.

⚠️ One thing to do before cloning for a real second customer: `api/config.json`
in the repo today still holds OHANT's actual pilots, aircraft and 600+
clients — it's only used as a first-boot seed if a deployment's OneDrive is
still empty. Once OHANT's own app has redeployed with this change (their real
data copies into their own OneDrive automatically, nothing lost), replace
`api/config.json`'s `pilots`/`aircraft`/`clients` with empty arrays before
deploying anyone else — OHANT is unaffected either way since their OneDrive
copy is what their app actually reads from that point on.
