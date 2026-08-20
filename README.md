# Availability Picker — Outlook Add-in (personal account version)

This version connects to a **personal Microsoft account** (outlook.com,
hotmail.com, live.com) — not your NCGA work account, and not deployed to
anyone else. Everything here is self-service; no admin approval needed.

## What's different from a corporate deployment

| Corporate | Personal |
|---|---|
| Azure AD app registered under NCGA's tenant | App registered under your own free Microsoft Entra directory (any Microsoft account gets one automatically) |
| Needs admin consent for `Calendars.Read` | You just consent for yourself on first sign-in — no admin step exists for personal accounts |
| Deployed org-wide via M365 admin center | You sideload it once, only into your own Outlook |
| `authority` = your tenant ID | `authority` = `consumers` (already set in `config.js`) |

## Step 1 — Register the Azure AD app (5 minutes, no admin needed)

1. Go to **[portal.azure.com](https://portal.azure.com)** and sign in with
   your **personal** Microsoft account. If you've never used Azure before,
   it'll auto-create a free default directory for you — that's normal and
   free, no credit card needed for this.
2. Search for **"App registrations"** → **New registration**.
3. Name it anything, e.g. `Availability Picker`.
4. Supported account types: choose **"Accounts in any organizational
   directory and personal Microsoft accounts"** (or the personal-only
   option if you see one — either works since `config.js` is pinned to
   `consumers`).
5. Redirect URI: type **Single-page application (SPA)**, value = wherever
   you'll host `taskpane.html` (see Step 3), e.g.
   `https://yourname.github.io/availability-picker/taskpane.html`
6. Register, then copy the **Application (client) ID** from the Overview page.
7. Go to **API permissions → Add a permission → Microsoft Graph →
   Delegated permissions** → add `Calendars.Read`. No admin consent button
   will appear (and none is needed) for personal accounts — you'll just
   approve it yourself the first time you sign in through the add-in.

## Step 2 — Fill in `config.js`

- `clientId` → the Application (client) ID from Step 1
- `redirectUri` → the exact hosted URL of `taskpane.html` (must match Step 1's redirect URI exactly)
- `authority` is already set to `https://login.microsoftonline.com/consumers` — leave it unless you also want to use a work account with the same add-in someday, in which case switch it to `common`.

## Step 3 — Host the files (pick the easiest for you)

Needs HTTPS; a free static host is plenty:
- **GitHub Pages** — easiest for a solo project. Push this folder to a
  GitHub repo, enable Pages on it, and you'll get a URL like
  `https://yourname.github.io/repo-name/`.
- **Azure Static Web Apps** — free tier, works fine too if you'd rather stay in the Microsoft ecosystem.

Once you have the real URL, replace every `https://REPLACE_WITH_YOUR_HOSTED_DOMAIN`
in `manifest.xml` with it, and double check `config.js`'s `redirectUri` matches exactly.

## Step 4 — Sideload it into your own Outlook

1. Go to **outlook.com → Settings (gear) → Get Add-ins → My add-ins → Add a custom add-in → Add from file**.
2. Upload your edited `manifest.xml`.
3. Open a new email — you'll see an **"Availability"** group with a **"Pick Times"** button.
4. Click it, sign in with your personal account, approve the calendar permission prompt (one-time), pick a range, and check that real slots show up.

That's it — no further deployment step, since it's just for you. If you
ever want to hand this to other people individually, each person repeats
Steps 1 (their own app registration, or you share the client ID and they
just sideload the manifest with your hosted URL) and 4 (sideload) on their
own account.

## Notes

- **Business hours and slot length** are adjustable right in the task pane.
- **Tentative events count as busy** by default — change the filter list in
  `taskpane.js`'s `fetchBusyEvents` (`["busy", "tentative", "oof"]`) if you'd
  rather tentative holds show as open.
- **Time zone** is handled automatically via your browser's local time zone.
- This only ever reads **your own** calendar and never sends anything by
  itself — you still review and hit Send yourself.
