# Mail Agent

A local-only starter for organizing a small sample of a Microsoft 365 mailbox. It uses Microsoft sign-in and Microsoft Graph delegated permissions. It does not collect your organisation password, modify messages, send mail, or store tokens on disk.

> This is a starter, not an organisation-approved production integration. Ask your Microsoft 365 administrator to approve the app registration and consent before connecting a work mailbox. The GitHub project should contain no real `.env` file, credentials, tokens, or mailbox data.

## What it does

- Signs in through the Microsoft identity platform using authorization code flow with PKCE.
- Requests `User.Read` and delegated `Mail.Read` only.
- Reads up to 25 recent messages from the signed-in user's mailbox.
- Groups messages using simple subject-based rules and displays subject, sender, date, and read status.
- Keeps the access token in process memory only. Restarting the app clears it.

It does **not** send, move, label, archive, or delete mail. The categories are local suggestions and do not change Outlook.

## Requirements

- Node.js 20 or newer.
- A Microsoft 365 work account.
- Organisation approval for an app registration and delegated Graph `Mail.Read` permission. Tenant policies may require administrator consent.

## Configure Microsoft Entra ID

Ask your administrator to register a single-tenant public client application for local development:

1. Select the **Mobile and desktop applications** platform and register this redirect URI:
   `http://localhost:3000/auth/callback`
2. Do not create a client secret for this local public-client flow.
3. Add Microsoft Graph **delegated** permission `Mail.Read`.
4. Obtain the tenant ID and application (client) ID. Do not grant application-level `Mail.Read` access.
5. Have the administrator review the permissions and provide consent if required by organisation policy.

Never put your email password, access token, refresh token, or client secret in `.env`, source code, prompts, or GitHub.

## Run locally

In PowerShell, from the project directory:

```powershell
Copy-Item .env.example .env
```

Edit `.env` and set `MSAL_CLIENT_ID` and `MSAL_TENANT_ID` to the values approved by your administrator. The client ID and tenant ID identify the app and tenant; they are not your mailbox credentials.

Then run:

```powershell
npm install
npm test
npm start
```

Open `http://localhost:3000` and choose **Sign in with Microsoft**. Enter credentials only on the organisation's Microsoft sign-in page. The server binds to loopback (`127.0.0.1`) and is intended for local development.

## Security notes

- `.env` is ignored by Git; commit only `.env.example`.
- OAuth uses a random state value and PKCE. Tokens are held only in server memory and never logged.
- Graph requests are delegated to the signed-in user and read-only.
- Only subject-based categories are generated; message contents are not downloaded or sent to an AI service.
- Do not expose this development server to the network or deploy it as-is. A production service needs an approved identity/session design, secure token cache, monitoring, data-retention controls, and security review.
- Stop the process to clear the in-memory token. Revoke the app's consent through your organisation's approved process if access should be removed.

## Tests

```powershell
npm test
```
