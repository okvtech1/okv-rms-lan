# OKV RMS - Apps Script back end: detailed setup guide

You only need ONE file for this part: `Code.gs`.
You do NOT need to create a Google Sheet by hand. The script creates the data sheet itself.
You do NOT need any other file on Apps Script. Every web page lives on GitHub Pages.

What you will do, in order:
1. Create a new Apps Script project and paste `Code.gs`.
2. Put your admin password in one place, and run ONE function: `setup`.
3. Deploy the script as a web app and copy its address (the "deployment URL").
4. Paste that address into the website pages (9 places).
5. Test it.

Time needed: about 15 minutes. You need a Google account (Gmail).

---------------------------------------------------------------
## Part 1 - Create the Apps Script project

1. Open your browser and go to **https://script.google.com**
2. Sign in with the Google account that should OWN the data. Keep using this same account later.
3. Click **+ New project** (top left).
4. A file named `Code.gs` opens with some default text, for example `function myFunction() {}`.
   Click inside the editor, press **Ctrl+A** (Cmd+A on Mac), then **Delete**. The editor is now empty.
5. Open the `Code.gs` file from this package in Notepad (Windows) or TextEdit (Mac). Press **Ctrl+A**, then **Ctrl+C**.
6. Go back to the Apps Script editor, click inside it, press **Ctrl+V**.
7. Click the **disk icon (Save)** or press **Ctrl+S**.
8. At the top left, click **Untitled project** and rename it to **OKV RMS**. Press Enter.

---------------------------------------------------------------
## Part 2 - Set the admin password (one time)

The admin password is used to sign in to the admin page on your website.

1. In the editor, find this near the top, in the section "one-time setup":
   `const INITIAL_ADMIN_PASSWORD = '';`
2. Type your password between the quotes, for example:
   `const INITIAL_ADMIN_PASSWORD = 'MyStrongPass#2026';`
   Rules: at least 8 characters. Write it down somewhere safe.
3. Save (Ctrl+S).

You will change this line back to empty in Part 3, after setup has run. Do not leave the password in the code.

---------------------------------------------------------------
## Part 3 - Run ONE function to create everything

1. At the top of the editor, find the **function dropdown** (it says `myFunction` or `setup`).
   Click it and choose **setup**.
2. Click **Run** (the play icon, ▶).
3. The first time only, Google asks for permission:
   a. Click **Review permissions**.
   b. Choose your Google account.
   c. If you see "Google hasn't verified this app", click **Advanced**, then **Go to OKV RMS (unsafe)**.
   d. Click **Allow**.
   This is normal. The script only works with your own Drive and Sheets.
4. Wait until the bottom panel says **Execution completed**. (The first run takes about 30 seconds.)
5. Look at the bottom panel, under **Execution log**. It shows a line like:
   `Setup complete. Data sheet: https://docs.google.com/spreadsheets/d/XXXX/edit | New activation codes: 1200`
   Click that link to open your new data sheet. It is called **OKV RMS Data** and it is in your Google Drive.
6. Check the sheet. It should have these tabs: **Submissions, Customers, Codes, Alerts, Messages, Payments, Audit, Config**.
   The **Codes** tab has 1,200 activation codes ready to use.

What `setup` did for you automatically:
- created the data sheet (OKV RMS Data) in your Drive, so you did not have to make it;
- created every tab with its headings and colours;
- filled the Config tab with the default settings;
- created 1,200 activation codes;
- saved the admin password (from Part 2) in the script's secure properties.

Safe to repeat: running `setup` again does not delete data or create duplicate codes.

6b. Now clean up. Back in the editor, change the line in Part 2 back to:
    `const INITIAL_ADMIN_PASSWORD = '';`
    and press Ctrl+S. The password is saved; you no longer need it in the code.

---------------------------------------------------------------
## Part 4 - Deploy as a web app (this gives you the deployment URL)

1. In the editor, click **Deploy** (top right) > **New deployment**.
2. Next to "Select type", click the **gear icon** and choose **Web app**.
3. Fill in:
   - Description: `OKV RMS v1`
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Click **Deploy**.
5. If asked, click **Authorize access**, choose your account, and allow (same steps as Part 3).
6. You will see **Web app** and a long address. Click **Copy** next to it.
   It looks like this, and ends with `/exec`:
   `https://script.google.com/macros/s/AKfycbx.......long.......XYZ/exec`
7. Save this address in a text file. This is the **deployment URL**. You will paste it in Part 5.

IMPORTANT: when you change `Code.gs` later, do NOT click New deployment again (that changes the address).
Instead: **Deploy > Manage deployments > pencil icon (Edit) > Version: New version > Deploy**.
The address stays the same.

---------------------------------------------------------------
## Part 5 - Paste the deployment URL into the website

Your website has 9 pages that talk to this back end. In each one there is exactly one line like this:

    var API_URL = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';

Replace only the text inside the quotes with your deployment URL, keeping the quotes:

    var API_URL = 'https://script.google.com/macros/s/AKfycb.......XYZ/exec';

The 9 pages are:

| File | What it does |
|---|---|
| index.html | Home page: free-trial, activation and contact forms |
| pricing.html | Pricing page |
| terms.html | Terms of use |
| privacy.html | Privacy policy |
| cookies.html | Cookie policy |
| licence.html | Licence agreement |
| refund.html | Refund policy |
| disclaimer.html | Disclaimer |
| admin.html | Admin dashboard (sign in with your admin password) |

Quick way to do it (recommended):
1. Open the website folder in Notepad++, VS Code, or any editor that has **Find in Files**.
2. Search for: `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE`
3. Replace it with your deployment URL. Replace in all files. You should see 9 changes.

Do NOT paste it into demo.html or product.html. They do not need it.

---------------------------------------------------------------
## Part 6 - Set the purchase link for the dashboard (optional)

The "Buy now" button in the software points to the address in the Config tab, key `pricing_url`.
The default is already https://www.okvtechnology.com/okvrms, which is the address on the brochure.
To change it: open the **Config** tab, find the row `pricing_url`, and edit the value in column B.

Optional: the admin dashboard link in e-mails.
In the Config tab, add a row:
- Column A: `admin_url`
- Column B: `https://YOUR-USERNAME.github.io/YOUR-REPO/admin.html`

---------------------------------------------------------------
## Part 7 - Test that it works

1. Upload the website to GitHub Pages (see the website README). Wait 2 minutes.
2. Open `https://YOUR-USERNAME.github.io/YOUR-REPO/admin.html`
3. Type the admin password from Part 2 and click Sign in.
   - If you see "This page is not connected yet", the URL in Part 5 is missing or has a typo.
   - If you see "Wrong password", use the exact password from Part 2 (passwords are case-sensitive).
4. Open `index.html`, fill in the free-trial form and submit.
5. Back in the Google Sheet, open the **Submissions** tab. Your test submission should be there.
6. Check your e-mail: an alert e-mail is sent to the OKV address when a submission arrives.

---------------------------------------------------------------
## Troubleshooting

| Problem | Fix |
|---|---|
| "The data sheet is not set up yet" | Run `setup` once from the editor (Part 3). |
| "Authorization required" when running | Click Review permissions and allow (Part 3, step 3). |
| Web page shows "The server did not answer" | The deployment URL is wrong, or the deployment is not set to Who has access: **Anyone**. Redeploy (Part 4). |
| "Wrong password" on admin | Use the exact password from Part 2 (case-sensitive). To set a NEW password: type it in `INITIAL_ADMIN_PASSWORD`, choose **setup**, click Run, then change the line back to `''` and save. |
| Changes to Code.gs not showing | Deploy > Manage deployments > Edit > New version > Deploy (Part 4). |
| Need to see errors | In the editor, click **Executions** (left sidebar) to see each run and its error. |

---------------------------------------------------------------
## Rules that must stay the same

- `LICENSE_SECRET` in `Code.gs` must equal `LICENSE_SECRET` in the installer's `okv_rms.py`.
  Change both together before you ship the installer, or do not change either.
- Keep the deployment URL private to the website. Anyone with it can send forms to your back end, but they cannot read your data or sign in without the admin password.
- Do not share the Google Sheet with customers.

## What Code.gs contains (for reference)

- `setup()` - the one function you run. Creates the data sheet, tabs, settings and codes.
- `setAdminPassword(pw)` - sets the admin password (also run automatically by `setup` when `INITIAL_ADMIN_PASSWORD` is filled in).
- `doGet` - public information for the website and activation checks.
- `doPost` - the website forms (trial, activation request, contact, install request, payment start) and the admin dashboard calls.
