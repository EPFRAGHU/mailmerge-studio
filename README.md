# MailMerge Studio ✉️

Also deployable on any server (VPS) via Docker — see **Deploying to a VPS with Coolify** below.

---

## Multi-user: Google Sign-In + Admin panel

Each person signs in with their own Google account and gets a private workspace:
their own SMTP connection (their Gmail + app password), mailing lists, templates,
attachments, send log and daily quota counter. Nobody sees anyone else's data.

**One-time setup — create Google OAuth credentials:**

1. Go to https://console.cloud.google.com → create a project (or reuse one).
2. APIs & Services → OAuth consent screen → External → fill app name + your email → add
   your colleagues' Gmail addresses under **Test users** (or publish the app).
3. APIs & Services → Credentials → Create Credentials → **OAuth client ID** →
   Web application. Under **Authorized redirect URIs** add exactly:
   `https://YOUR-DOMAIN/auth/callback` (and `http://127.0.0.1:5000/auth/callback` for local testing).
4. Copy the **Client ID** and **Client secret**.

**In Coolify → your app → Environment Variables, add:**

- `GOOGLE_CLIENT_ID` — the client ID
- `GOOGLE_CLIENT_SECRET` — the client secret
- `BASE_URL` — your public URL, e.g. `https://mailmerge.yourdomain.com` (used for the OAuth redirect)
- `ADMIN_EMAILS` — your email, e.g. `raghunatha.maharana@gmail.com` (you get the Admin panel)

Then redeploy. The first person to sign in becomes admin automatically if `ADMIN_EMAILS`
is not set. In the Admin panel you can see every user's lifetime send count and
block/unblock accounts.



---


A friendly, private, desktop mail-merge app. Import an **Excel or CSV** contact list,
compose a rich email with merge tags like `{{Name}}`, attach files (same for everyone, or
personalised per recipient), and send from **your own Gmail / Outlook / Yahoo / Zoho /
iCloud / any SMTP** mailbox. No third-party email service ever sees your contacts.

## Quick start (Windows)

1. Install Python from https://python.org if you don't have it (tick "Add to PATH").
2. Double-click **`install_and_run.bat`** — the first run installs everything automatically.
3. Your browser opens at `http://127.0.0.1:5000`.

## How to use

1. **Connect Email** — pick your provider, enter your email + *app password*
   (Gmail: Google Account → Security → 2‑Step Verification → App passwords; other providers
   have similar app-password settings). Save, then send yourself a test email.
2. **Contacts** — drop in your `.xlsx` or `.csv`. The first row must be headers
   (e.g. `Email, Name, Company`). Tick/untick rows to choose recipients.
3. **Compose** — write like in Word (bold, lists, colours, links, images). Click a merge-tag
   pill to insert `{{ColumnName}}` anywhere. Preview with real data from row 1.
4. **Attachments** — upload files everyone gets. For personalised attachments, add a column
   (e.g. `InvoiceFile`) with file names separated by `;` and select that column.
5. **Send & Track** — send to all, only checked rows, or a single test. Pacing controls
   (delay between emails, batch pauses) keep your mailbox safe. Watch live results,
   stop anytime, and export the report.

**Dry run mode** lets you verify everything without sending a single email.

## Notes

- Your settings and password are stored only on this computer (`data/config.json`).
- A full send log lives in `data/send_log.json`.
- Free mailboxes limit how many emails you can send per day (Gmail ≈ 500). Pacing
  defaults are conservative — raise them slowly.
