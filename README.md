# vTiger CRM Auto Update

A static, front-end-only tool that reads an Excel sheet, finds the cells you highlighted in **red**, and pushes those changes straight into **vTiger CRM** through its public Webservice API — matching records by **Ticket Number**.

No build step and no server of your own. The app runs in your browser; the only server-side piece is a small **Google Apps Script** that keeps the user list and shared settings in a private **Google Sheet** and checks sign-in passwords.

## How it works

The app opens on a **login screen** that asks for your **username** and **password**. These are checked by the [user database](#user-database-google-sheet) (a private Google Sheet), not in the browser. Once the password matches, the Sheet returns your vTiger access key, and the app uses it to sign in via vTiger's own `login` webservice call. Your entry in the Users list also decides your **role** for this session:

- **Admin** — full access, including **Settings** and **Documentation**.
- **Team Member** — access to the ticket-update wizard and **Reconnect** only; Settings and Documentation are hidden.

Once signed in, a status pill in the header shows *Connected* / *Connection failed*, and **Reconnect** lets you retry without logging out. The actual ticket-update workflow is a 4-step wizard:

1. **Upload** — upload the `.xlsx` file that contains your tracked changes and pick the worksheet to scan.
2. **Map Fields** — tell the app which Excel column holds the Ticket Number (the unique identifier), and map every other red-highlighted column to the matching vTiger field API name. Mappings are remembered in your browser for next time.
3. **Review** — see exactly which tickets and fields will change. Uncheck any row or click a highlighted value to exclude just that field.
4. **Update** — click the button and the app looks up each ticket by number, retrieves the current record, applies only the mapped/red fields, and calls vTiger's `update` operation. A results table shows success/failure per ticket.

Cell scanning happens right after you pick a worksheet in step 1: the app inspects every cell's fill color and flags cells with a red-ish background as pending changes.

## Passwords: set at first sign-in

No one — not even an Admin — sets another user's password directly. Instead:

1. An Admin adds a user (in **Settings → Users**, or by typing a row straight into the Sheet's **Users** tab) with just a username, name, vTiger access key, and role — no password.
2. The first time that person signs in, a **Confirm Password** field appears as soon as they leave the username field (the Sheet reports they have no password yet). Whatever they enter becomes their password.
3. Every sign-in after that asks for that same username + password. A wrong password is rejected by the Sheet before the app ever touches vTiger, and **5 wrong attempts lock that username for 15 minutes**.
4. **Change Password** — once signed in, anyone can change their own password from the profile menu (click the avatar → **Change Password**). It asks for the current password, then the new one twice, and saves it to the Sheet.
5. **Forgot password?** — a link under the password field on the login screen. The user enters their username and their own **vTiger access key** (from *My Preferences* in vTiger) to prove the account is theirs, then sets a new password. Wrong access keys count toward the same 5-attempt lockout as wrong passwords.
6. An **Admin** can still **reset** someone's password from Settings → Users → Password column → **Reset** (or by clearing the `passwordHash` and `passwordSalt` cells in the Sheet). The next time they sign in, the Confirm Password field reappears — the admin never sees or sets the new one.
7. **Remember my password on this device** — a toggle on the login screen. Turn it on and both the **username and password** are saved in this browser's `localStorage`, so the login screen shows up already filled in (including after **Log Out** or a [session timeout](#settings-admins-only)). Leave it off on any shared or public computer.

### Password rules

Every new password (first sign-in, Change Password, Forgot password) must have:

- at least **8 characters**, with an **uppercase letter**, a **lowercase letter**, a **number** and a **special character**;
- **no character repeated 3+ times in a row** (`111`, `aaa`);
- **no 3+ sequential characters**, up or down, digits or letters (`123`, `987`, `abc`, `CBA`);
- and it can't be **any of the user's last 3 passwords** (the current one included). An Admin reset keeps this history, so a reset can't be used to go back to an old password.

The password fields show these rules as a live checklist. The Apps Script enforces them again on the server, so they can't be bypassed from the browser. Passwords set before these rules existed keep working until they're next changed.

Passwords are stored in the Sheet only as **salted, peppered, iterated HMAC-SHA256 hashes**. The pepper lives in the Apps Script's Script Properties, not in the Sheet, so even someone who gets a copy of the Sheet can't check password guesses offline.

## Project structure

```
index.html                          Markup / app shell (auto-connect header + 4-step wizard + settings modal)
css/style.css                       Page-layout only (spacing/structure) — component look-and-feel comes from the design system
js/config.js                        Deployment config — the Apps Script web app URL (userApiUrl)
js/app.js                           All application logic (Excel parsing, vTiger API calls, UI wiring)
apps-script/Code.gs                 The user database backend — paste into the Google Sheet's Apps Script editor
assets/                             Logo/favicon (SVG)
Altametrics Growth Design System/   The design system this UI is built on (tokens, components, icons, fonts)
```

The Excel parsing library ([SheetJS/xlsx](https://github.com/SheetJS/sheetjs)) is loaded from a CDN in `index.html`. If you need a fully offline copy, download `xlsx.full.min.js` and reference it locally instead of the CDN `<script>` tag.

## Styling: Altametrics Growth Design System

All buttons, inputs, tables, the settings dialog, pills/tags, the avatar, and the progress/spinner
UI are the **Altametrics Growth Design System**'s own component classes (`.btn`, `.form-control`,
`.pt-table`/`.pt-wrap`, `.fn-modal-*`, `.fn-tag`, `.fn-avtar`, `.fn-progress`, `.fn-spinner`, Phosphor
icons via `<iconify-icon>`), loaded from `Altametrics Growth Design System/styles.css`. `css/style.css`
only adds page-specific layout (the login screen, the step wizard shell, the upload dropzone) — every
value in it is a design-system token (`var(--space-*)`, `var(--radius-*)`, color tokens), never a raw
hex or invented size. See `Altametrics Growth Design System/docs/` for the full component/token
reference if you extend this UI.

## User database (Google Sheet)

Users, their roles and access keys, their (hashed) passwords, and the shared CRM settings all live in one private Google Sheet:
[`1cq0ra2yQLwY3puqF68Sdhb4DZf7EzqnydbXDppwbCOg`](https://docs.google.com/spreadsheets/d/1cq0ra2yQLwY3puqF68Sdhb4DZf7EzqnydbXDppwbCOg/edit). The browser never reads the Sheet; it calls a Google Apps Script web app (`apps-script/Code.gs`) that runs as the Sheet's owner, so the Sheet itself stays unshared.

```
Browser (js/app.js) ──POST {action, username, password | token}──▶ Apps Script web app ──▶ private Google Sheet
                    ◀── {ok, user, token} ─────────────────────────┘   (runs as the owner)
```

### One-time setup

1. **Open the Sheet → Extensions → Apps Script.** Replace the contents of `Code.gs` with `apps-script/Code.gs` from this repo, and save.
2. **Run `setup`.** Pick `setup` in the function dropdown and click **Run**, then approve the permissions prompt (it needs access to the Sheet). This creates the **Users** and **Settings** tabs with their headers, plus the password pepper and token secret in Script Properties.
3. **Add the first Admin.** In the **Users** tab, fill one row: `username` (their vTiger username), `firstName`, `lastName`, `role` = `admin`, and `accessKey` (from *My Preferences* in vTiger). Leave `passwordHash` and `passwordSalt` empty — they'll create a password at their first sign-in.
4. **Fill in the Settings tab** (optional — the app's built-in defaults are used for anything left blank): `url`, `module`, `idField`, `timeoutMinutes`. Admins can change these later from the app's Settings dialog.
5. **Deploy.** Click **Deploy → New deployment → Web app**, set **Execute as: Me** and **Who has access: Anyone**, and click **Deploy**. Copy the **Web app URL** (it ends in `/exec`).
6. **Paste the URL into `js/config.js`** as `userApiUrl`, then publish the app as usual.

"Who has access: Anyone" means anyone can *call* the script. They still need a valid username + password to get anything back, and only Admins can list or change users. The Sheet is never shared.

### Users tab columns

| Column | Meaning |
|---|---|
| `username` | vTiger username; matched case-insensitively |
| `firstName`, `lastName` | Display only (profile menu, Users table) |
| `role` | `admin` or `member` (Team Member) |
| `accessKey` | That user's vTiger access key |
| `passwordHash`, `passwordSalt` | Written by the script. Clear both to reset a password |
| `passwordHistory` | Written by the script: hashes of the last 3 passwords, used to block reuse. Don't edit |
| `updatedAt` | Last time the script changed the row |

Columns can be reordered, but the header names must stay exactly as above. If an older Users tab doesn't have `passwordHistory` yet, the script adds it automatically.

### Updating the script later

Saving the code in the editor does **not** change the live web app. After editing `Code.gs`, go to **Deploy → Manage deployments → (your deployment) → Edit → Version: New version → Deploy**. The URL stays the same, so `js/config.js` doesn't change.

Don't delete the `PASSWORD_PEPPER` Script Property: without it every stored password stops matching, and everyone would need an Admin reset. Changing `TOKEN_SECRET` just signs everyone out.

## Important: CORS

This app calls vTiger's `webservice.php` directly from the browser using `fetch`. Because there is intentionally **no proxy** in front of vTiger, the browser enforces the standard **Cross-Origin Resource Sharing (CORS)** policy:

- If you host this app on the **same domain** as your vTiger CRM, requests work without any extra configuration.
- If you host it elsewhere (e.g. GitHub Pages), your vTiger web server (Apache/Nginx in front of vTiger) must send `Access-Control-Allow-Origin` headers permitting this app's origin, and must respond correctly to any CORS preflight requests. Ask your vTiger administrator to enable this if you see a network/CORS error on the login screen.

Because vTiger's `webservice.php` is normally same-origin only, this is a limitation of using a pure static/front-end app — it is not a bug in this project.

## Detecting "red" highlights

The app treats a cell as changed if its fill color is red-dominant (red channel clearly higher than green and blue), which covers plain red fills as well as common "light red" highlight tones used for tracked changes/conditional formatting. Cell background colors are read from the workbook's styles, so the source file must be a real `.xlsx`/`.xls` with cell fill formatting (not just red text).

## Field mapping tips

- The **Ticket Number** column is the unique identifier used to look up each vTiger record; it does not need to be red-highlighted itself.
- Only columns that contain at least one red cell are shown in the mapping step.
- Field names must match vTiger's internal **API field names**, not the on-screen labels (e.g. use `ticket_title`, not "Title"). After connecting, the app calls vTiger's `describe` operation and offers matching field names as autocomplete suggestions in the mapping step; a field name not recognized by the connected module is outlined in red.
- Mappings are saved per-browser (via `localStorage`) so you don't have to redo them for every upload with the same column headers.
- For reference/owner fields (e.g. "assigned to"), you can put a plain name in Excel (like a vTiger username) — the app looks up the matching Users/Groups record and resolves it to the real vTiger id before sending the update.

## Why a "Success" result can still mean nothing changed

vTiger's `update` Webservice operation returns `success: true` as long as the request itself is well-formed — it does **not** error out just because a field name doesn't exist on the module, or because a reference/owner field was given a value it can't use (e.g. a display name instead of a record id). In both cases vTiger silently ignores that one field rather than rejecting the whole update.

To catch this, the app now:

1. Rejects any mapped field name that isn't in the connected module's field list (from `describe`) *before* sending anything, instead of letting vTiger silently drop it.
2. Attempts to resolve name-like values for reference/owner fields to actual vTiger record ids.
3. Compares the record vTiger returns after `update` against the values you intended to set, and reports the row as **failed** (with the specific field names) if any of them didn't actually persist — instead of trusting the bare `success: true`.

## Settings (Admins only)

Click **Settings** in the header to set the shared **CRM URL**, **ticket module**, **unique identifier field**, and **Session Timeout** once. These are saved to the **Settings** tab of the Google Sheet and apply to everyone, on every device (each browser keeps a cached copy and refreshes it every time the app loads). Use **Reset to Defaults** to restore the built-in values.

**Session Timeout (minutes)** signs a user out automatically after that many minutes with no mouse/keyboard/scroll activity — same effect as clicking **Log Out**, with a toast explaining why. Set it to `0` (or leave it blank) to disable auto sign-out entirely. Changing it while someone is signed in applies immediately, without needing to reconnect.

## Users (Admins only)

Also inside **Settings**, Admins manage the people who can sign in. The table reads from and writes to the **Users** tab of the Google Sheet, so every change applies to everyone immediately. Each entry has a **username**, **first/last name**, that user's own **vTiger access key**, and a **role**:

- **Add / Update User** adds a new user, or overwrites an existing user's name, access key, and role. It never sets their password (see [Passwords](#passwords-set-at-first-sign-in) above). When updating an existing user you can leave Access Key blank to keep their current key.
- Each row also has its own **Role** dropdown for quickly promoting/demoting someone, and a **Show/Hide** toggle to reveal an access key when you need to check it.
- The **Password** column shows **Not set** (they'll create one at their next sign-in) or **Set**, with a **Reset** button to put them through that first-sign-in flow again.
- Users can also change or reset their own password without an Admin (see [Passwords](#passwords-set-at-first-sign-in)).
- **Remove** asks for confirmation. You can't remove your own account, so there's always at least one Admin who can sign in.
- First/last name is used only for display. If left blank, the username is shown instead.
- Anyone who tries to log in with a username *not* on this list is rejected with "Username not found."

Admins can also edit the Users tab directly in Google Sheets, which is handy for bulk changes. The app picks up the changes on the next sign-in, reload, or **Reconnect**.

## Sharing the app with your team

Send teammates the app's URL — that's it. Their username is looked up in the Google Sheet, so there's nothing to import on a new device. To onboard someone:

1. **Admin:** add them in **Settings → Users** (or add a row to the Users tab).
2. **Teammate:** open the app, enter their username, and create their own password on that first sign-in.

Removing someone from the list stops them from signing in, and signs them out on their next reload or Reconnect.

## Login sessions

A successful sign-in returns a **signed session token** from the Apps Script (valid for 12 hours). The app keeps that token and your username in `sessionStorage` (not `localStorage`). Your access key and role are always fetched fresh with the token, so an admin's changes to your role or key take effect the next time you reload or reconnect. This means:

- Reloading the page or navigating within the same browser tab keeps you signed in.
- Closing the tab/browser clears it — you'll need to sign in again next time.
- Click your avatar in the header to open the profile menu, showing your name, username, and role, and to **Log Out** (which clears it immediately).
- If your username is removed from the Users list while you're signed in, **Reconnect** will fail and tell you to contact your admin.

## Security notes

> ⚠️ **Earlier versions of this app (in the previous `CRMAutoUpdate` repository) had the `alta_support` vTiger access key hardcoded in `js/app.js`.** If that repository or app was ever public, rotate that access key in vTiger and put the new one only in the Google Sheet. Access keys belong in the Sheet, never in this repository.

- **Keep the Google Sheet private.** Don't share it as "anyone with the link" — the app never needs that, and it would expose every access key. Share it only with the people who are meant to edit users directly.
- Passwords are checked by the Apps Script, never in the browser, and stored only as salted, peppered, iterated hashes. Repeated wrong passwords lock that username for 15 minutes.
- After a correct password, the browser receives **only that user's own** vTiger access key, because it calls vTiger directly. Only Admins can fetch the full list (including keys), and the script re-checks the caller's role in the Sheet on every Admin request.
- Your login password is **not** your vTiger password, and vTiger never sees it.
- The login session (username + signed token) lives only in `sessionStorage` for the current browser tab. If "Remember my password" is on, the password is also kept in this browser's `localStorage` — only turn that on for a personal device.
- On first load, this version deletes the old local copy of the Users list (with plain-text passwords and access keys) that earlier versions kept in each browser.
- Always review the **Review** step carefully before clicking **Update** — updates are applied immediately to live CRM records.

## Running locally

No build tools required. Set `userApiUrl` in `js/config.js` first (see [User database](#user-database-google-sheet)), then either:

- Open `index.html` directly in a browser, or
- Serve the folder with any static file server, e.g.:

```bash
npx serve .
```

## Deploying

Any static host works (GitHub Pages, Netlify, an internal web server, or the same server that runs vTiger). Just publish the folder as-is — see the [CORS](#important-cors) note above for cross-domain hosting.

## License

Use and adapt freely within your organization.
