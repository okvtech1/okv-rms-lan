# OKV RMS - website setup guide

## 1. What is in this folder

Everything is in ONE folder (this one). Every page is one self-contained file: styles, scripts and
pictures are inside the file, so each page works on its own, with no other files and no internet links.

| File | What it is | Links to |
|---|---|---|
| index.html | Landing page. Start here. | every page, the demo and the sections on this page |
| product.html | Product page: what OKV RMS does, with screenshots | index, pricing, demo, terms, privacy |
| demo.html | THE demo: the full working system with sample data | index, pricing |
| pricing.html | Price, free trial and buy button | index, product, licence, policy pages |
| licence.html | Software licence agreement | all policy pages |
| terms.html, privacy.html, cookies.html, refund.html, disclaimer.html | Policy pages | all policy pages and index |
| admin.html | OKV Technology admin dashboard (sign in with the admin password) | none (back end only) |
| README_WEBSITE.md | This guide | - |

Every page has a menu (top) and footer that link to the others. Open index.html first.

## 2. Connect the website to the back end (one paste per page)

The back end is Code.gs on Apps Script (see apps_script/SETUP_GUIDE_APPS_SCRIPT.md, Part 4).
It gives you a deployment URL that ends with /exec.

Nine pages have one line to change:

    var API_URL = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';

Pages: index, pricing, terms, privacy, cookies, licence, refund, disclaimer, admin.

1. In your editor, use Find in Files and search for `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE`.
2. Replace it with your /exec URL (keep the quotes). Expect 9 replacements.

Do not change demo.html or product.html. They do not connect to the back end.

## 3. Upload to GitHub Pages (free)

1. Sign in to https://github.com. Click **+** (top right) > **New repository**.
2. Name it, for example `okv-rms-site`. Choose **Public**. Click **Create repository**.
3. Click **uploading an existing file**. Drag ALL the files in this folder into the box.
   Wait for the upload to finish, then click **Commit changes**.
4. Click **Settings** > **Pages** (left menu).
5. Under **Build and deployment**, Source: **Deploy from a branch**. Branch: **main**, folder: **/ (root)**. Click **Save**.
6. Wait 1 to 2 minutes and refresh. GitHub shows your address, for example
   `https://YOUR-USERNAME.github.io/okv-rms-site/`
7. Open it. It should show index.html.

Keep all files at the top level of the repository (not inside another folder). File names are case-sensitive.

## 4. The demo page (demo.html)

The demo is the real OKV RMS screens with sample data. It is always signed in, so there is no login.
Use the buttons at the bottom right to view it as Admin, Manager, Cashier, Waiter or Kitchen.

Some features are locked on purpose. Each locked control shows a gold
"Full version only: buy to unlock" button that opens the pricing page in a new tab:
- Settings > Server: "Keep running in the background when the window is closed"
- Settings > Business logo: "Choose logo"
- Start automatically with Windows (where shown)

Everything else works: sales, products, customers, stock, reports, expenses, staff, import and export,
receipts, kitchen screen, and the manuals.

## 5. Test checklist (after upload)

- [ ] index.html shows its screenshots (kitchen, sell, tables, reports, phone). Scroll down to check them all.
- [ ] Menu links on every page open the right page: About, How it works, Features, Product, Live demo, Pricing, Reseller, FAQ.
- [ ] The "Product tour" button on index opens product.html.
- [ ] Every policy page opens and links back to index.
- [ ] demo.html opens with sample data, and the role buttons switch the view.
- [ ] A gold "Full version only" button opens pricing in a new tab.
- [ ] admin.html: after pasting the URL, sign in with the admin password. Before pasting, it says "not connected yet" (expected).
- [ ] On a phone, pages fit the screen with no sideways scrolling.

## Notes

- Do not put your admin password in any page. You type it on admin.html and the back end checks it.
- To update the site later, upload the changed files again to the repository and commit. It updates within about a minute.
- If a page says "This page is not connected yet", the API_URL line still has the placeholder text.
