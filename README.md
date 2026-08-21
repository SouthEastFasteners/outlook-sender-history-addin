# Sender History — Outlook Add-in

A task pane add-in that shows the last 6 emails from the sender of the
message currently being read, so the user can review history without
leaving the open email.

This scaffold is built as a **web add-in** (Office.js + React + Microsoft
Graph), which runs unchanged across Outlook on Windows (new and classic),
Mac, web, and mobile.

---

## 0. Prerequisites

- Node.js 18+ and npm
- Access to your organization's **Microsoft Entra ID** (Azure AD) tenant
  with permission to register an app (or a tenant admin who can do it for you)
- Outlook desktop, web, or new Outlook to test in
- (Recommended) VS Code

---

## 1. Register the app in Microsoft Entra ID

This is the one-time step that lets the add-in silently sign in your
already-authenticated AD users (SSO) and call Microsoft Graph on their
behalf.

1. Go to **portal.azure.com → Microsoft Entra ID → App registrations → New registration**
2. Name: `Sender History Outlook Add-in`
3. Supported account types: **Accounts in this organizational directory only** (single tenant)
4. Redirect URI: leave blank for now (added in step 1.5)
5. Click **Register** — note the **Application (client) ID** and **Directory (tenant) ID**, you'll need both shortly

### 1.1 Expose an API (required for Office SSO)

1. In the app registration → **Expose an API**
2. Click **Add** next to Application ID URI → accept the default
   `api://<your-domain-or-appid>` (Office SSO requires this exact format;
   you'll paste this into the manifest later)
3. Click **Add a scope**:
   - Scope name: `access_as_user`
   - Who can consent: Admins and users
   - Admin consent display name: `Access Sender History as the signed-in user`
   - State: Enabled
4. Under **Authorized client applications**, add these well-known Office
   client IDs so Outlook is allowed to request your scope silently:
   ```
   ea5a67f6-b6f3-4338-b240-c655ddc3cc8e   (Office on the web)
   d3590ed6-52b3-4102-aeff-aad2292ab01c   (Office desktop client)
   0ec893e0-5785-4de6-99da-4ed124e5296c   (Office desktop client, other)
   57fb890c-0dab-4253-a5e0-7188c88b2bb4   (Outlook on the web)
   d73f4b35-55c9-48c3-97a8-4d9d1a89a7bb   (Outlook desktop client, other)
   ```

### 1.2 API permissions

1. **API permissions → Add a permission → Microsoft Graph → Delegated permissions**
2. Add: `Mail.Read`, `User.Read`, `openid`, `profile`
3. Click **Grant admin consent for [Your Org]** — this removes the
   individual permission popup for every user (recommended for internal
   deployment)

### 1.3 Authentication

1. **Authentication → Add a platform → Single-page application**
2. Add redirect URIs (update the domain once you know your hosting URL —
   for local dev use):
   ```
   https://localhost:3000/taskpane.html
   ```
3. Under **Advanced settings**, enable **Allow public client flows** → Yes

You now have everything needed for the manifest's `WebApplicationInfo`
block (already filled in below — just swap in your IDs).

---

## 2. Install dependencies

```bash
cd outlook-sender-history-addin
npm install
```

## 3. Configure your IDs

Open `src/taskpane/authConfig.ts` and replace the placeholders with the
values from Step 1:

```ts
export const msalConfig = {
  auth: {
    clientId: "YOUR_APPLICATION_CLIENT_ID",
    authority: "https://login.microsoftonline.com/YOUR_TENANT_ID",
  },
};
```

Open `manifest.xml` and replace:
- `YOUR_APPLICATION_CLIENT_ID` (two places)
- `api://YOUR_DOMAIN/YOUR_APPLICATION_CLIENT_ID` (the Application ID URI from step 1.1)
- `https://localhost:3000` with your real hosting domain once deployed

## 4. Run locally

```bash
npm run dev-server
```

This starts a local HTTPS dev server at `https://localhost:3000` using
`office-addin-dev-certs` (installed automatically on first run — accept
the certificate prompt).

Then sideload it into Outlook for testing:

```bash
npm start
```

This uses `office-addin-debugging` to sideload the manifest into your
default Outlook client and launch the task pane automatically.

**Manual sideload alternative** (Outlook on the web):
1. Outlook on the web → **Settings (gear) → Manage add-ins → My add-ins**
2. **Add a custom add-in → Add from file** → select `manifest.xml`

---

## 5. Deploy for real use

### 5.1 Host the static files

Build the production bundle:

```bash
npm run build
```

This outputs static files to `/dist`. Host them anywhere with HTTPS —
Azure Static Web Apps (free tier is generally enough for an internal
tool), an existing internal web server, or GitHub Pages. Update every
`https://localhost:3000` reference in `manifest.xml` to your real hosting
URL once deployed, and add that same URL as a redirect URI back in the
Entra ID app registration (Step 1.3).

### 5.2 Push it to your whole org

Since your users are already in Microsoft 365 / Entra ID, skip manual
sideloading entirely:

1. **Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps**
2. Upload `manifest.xml`
3. Assign to a security group (pilot group first, then expand) or the
   whole org
4. Users see it appear automatically in Outlook — no install action
   needed on their end

---

## Project structure

```
outlook-sender-history-addin/
├── manifest.xml              # Add-in manifest incl. SSO config
├── package.json
├── webpack.config.js
├── tsconfig.json
├── assets/                   # Icons (16/32/80px, referenced in manifest)
└── src/
    └── taskpane/
        ├── taskpane.html
        ├── taskpane.css
        ├── authConfig.ts      # MSAL / Entra ID config — fill in your IDs
        ├── auth.ts            # Nested App Auth token acquisition
        ├── graph.ts           # Microsoft Graph calls (fetch sender history)
        ├── index.tsx          # React entry point
        └── components/
            └── App.tsx         # Main task pane UI
```

## Notes

- **`Mail.Read` scope only reads the current user's own mailbox** — this
  matches "show history in my mailbox," not shared/delegated mailboxes.
  If you later need shared mailbox support, you'll need `Mail.Read.Shared`
  and a broader admin consent conversation.
- **Task pane pinning**: by default the pane closes when the user switches
  messages unless pinned. The manifest below enables the pin icon; consider
  adding a one-time tooltip nudging users to pin it on first use.
- No Microsoft licensing cost for any of this — see cost breakdown from
  our earlier discussion. Only variable cost is hosting (often $0 on a
  free tier).
