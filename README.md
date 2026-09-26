# MailMerge Studio ✉️

Also deployable on any server (VPS) via Docker — see **Deploying to a VPS with Coolify** below.

---

## Deploying to a VPS with Coolify (Hostinger)

1. **Push this folder to GitHub** (from VS Code terminal, inside the MailMergeApp folder):

   ```powershell
   git init
   git add .
   git commit -m "MailMerge Studio"
   # create an empty repo on github.com first, then:
   git remote add origin https://github.com/YOUR_USERNAME/mailmerge-studio.git
   git push -u origin main
   ```

2. **In Coolify** (on your Hostinger VPS): New Resource → Dockerfile → pick your
   `mailmerge-studio` repo and main branch.

3. **Set the port** to 5000 in the deployment settings, and add a **persistent volume**
   mapping `/app/data` — this keeps your SMTP config, saved lists, templates and send
   logs across restarts.

4. **Domain & HTTPS**: in Coolify, attach your domain — it handles TLS certificates
   automatically.

5. **Security note (important):** the app currently has no login screen. Anyone who
   discovers the URL could use it. Host it on a domain nobody knows and consider adding
   a password (an `ACCESS_PASSWORD` env var hook is prepared in `.env.example`) before
   sharing it publicly.

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
