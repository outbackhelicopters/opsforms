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
- Access to your **Digital Ocean account** (where the app is deployed)
- Access to your **Microsoft 365 admin account** (needed to set up the app registration)

---

## Part 1 — Push the updated code to GitHub

1. Open **GitHub Desktop** on your Mac
2. You'll see changed files listed on the left
3. In the **Summary** box at the bottom left, type: `Switch to Office 365`
4. Click **Commit to main**
5. Click **Push origin** (top right)

Digital Ocean will start redeploying automatically.

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

## Part 3 — Add credentials to Digital Ocean

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

   The `ONEDRIVE_FOLDER` is already set to `Helicopter Paperwork` — you can leave it as is, or change it to whatever folder name you want.

5. Click **Save** — Digital Ocean will redeploy with the new settings

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

**Everything still showing old version on the app**
→ On the iPad: Settings → Safari → Clear History and Website Data → reopen the app.

---

## What happens once it's all set up

- Pilots work completely offline in the field — no internet needed
- When they tap Send, if there's no internet it saves locally and waits
- The moment the iPad gets any connection it sends automatically
- The top bar shows: **orange** = waiting to send, **green** = all sent
- You never need to touch any of this again — it runs automatically
