import "dotenv/config";
import { createHash, randomBytes } from "node:crypto";
import { createServer } from "node:http";
import { PublicClientApplication } from "@azure/msal-node";
import { categorizeMessage, escapeHtml } from "./organize.js";

const port = Number.parseInt(process.env.PORT ?? "3000", 10);
const host = process.env.HOST ?? "127.0.0.1";
const redirectUri = process.env.MSAL_REDIRECT_URI ?? "http://localhost:3000/auth/callback";
const clientId = process.env.MSAL_CLIENT_ID;
const tenantId = process.env.MSAL_TENANT_ID;
const scopes = ["Mail.Read"];
const graphMessagesUrl =
  "https://graph.microsoft.com/v1.0/me/messages?$top=25&$orderby=receivedDateTime%20desc&$select=id,subject,from,receivedDateTime,isRead";
const pendingAuth = new Map();
let accessToken;

if (!Number.isInteger(port) || port < 1 || port > 65535) {
  throw new Error("PORT must be a valid TCP port.");
}

if (!clientId || !tenantId) {
  throw new Error("Set MSAL_CLIENT_ID and MSAL_TENANT_ID in your local .env file.");
}

const redirect = new URL(redirectUri);
if (
  redirect.protocol !== "http:" ||
  redirect.hostname !== "localhost" ||
  redirect.port !== String(port) ||
  redirect.pathname !== "/auth/callback"
) {
  throw new Error("MSAL_REDIRECT_URI must be the configured localhost callback URL.");
}

const msal = new PublicClientApplication({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}`
  }
});

function sendHtml(response, status, content) {
  response.writeHead(status, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store",
    "Content-Security-Policy": "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'",
    "Referrer-Policy": "no-referrer",
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "DENY"
  });
  response.end(content);
}

function page(title, content) {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(title)} - Mail Organizer</title>
  <style>
    body{font:16px/1.5 system-ui,sans-serif;max-width:900px;margin:3rem auto;padding:0 1rem;color:#172b3a}
    h1{color:#10243a}a,button{background:#116d69;color:white;padding:.65rem 1rem;border:0;border-radius:.35rem;text-decoration:none;font:inherit}
    .note{background:#e5f4f2;padding:1rem;border-radius:.4rem}li{margin:.5rem 0}.muted{color:#5f7080}
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  ${content}
</body>
</html>`;
}

function renderMessages(messages) {
  const categories = new Map();

  for (const message of messages) {
    const category = categorizeMessage(message);
    const items = categories.get(category) ?? [];
    items.push(message);
    categories.set(category, items);
  }

  if (messages.length === 0) {
    return "<p>No messages were returned for this page.</p>";
  }

  return [...categories.entries()]
    .map(([category, items]) => {
      const list = items
        .map((message) => {
          const sender = message.from?.emailAddress?.address ?? "Unknown sender";
          const subject = message.subject || "(No subject)";
          const date = message.receivedDateTime
            ? new Date(message.receivedDateTime).toLocaleString()
            : "Date unavailable";
          const unread = message.isRead === false ? " · Unread" : "";

          return `<li><strong>${escapeHtml(subject)}</strong> — ${escapeHtml(sender)}<br><span class="muted">${escapeHtml(date)}${unread}</span></li>`;
        })
        .join("");

      return `<section><h2>${escapeHtml(category)} (${items.length})</h2><ul>${list}</ul></section>`;
    })
    .join("");
}

function createPkcePair() {
  const verifier = randomBytes(32).toString("base64url");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  return { verifier, challenge };
}

function prunePendingAuth() {
  const cutoff = Date.now() - 5 * 60 * 1000;
  for (const [state, request] of pendingAuth) {
    if (request.createdAt < cutoff) pendingAuth.delete(state);
  }
}

async function handleRequest(request, response) {
  const requestUrl = new URL(request.url ?? "/", `http://${host}:${port}`);

  if (requestUrl.pathname === "/" && request.method === "GET") {
    sendHtml(
      response,
      200,
      page(
        "Local mail organizer",
        `<p class="note">This starter requests delegated, read-only mailbox access. It only displays a small categorized message list; it does not send, move, or delete messages.</p>
         <p>Sign in through your organisation’s Microsoft page. Your password is never entered into this app.</p>
         <p><a href="/login">Sign in with Microsoft</a></p>`
      )
    );
    return;
  }

  if (requestUrl.pathname === "/login" && request.method === "GET") {
    prunePendingAuth();
    const state = randomBytes(32).toString("base64url");
    const { verifier, challenge } = createPkcePair();
    pendingAuth.set(state, { verifier, createdAt: Date.now() });

    const authUrl = await msal.getAuthCodeUrl({
      scopes,
      redirectUri,
      state,
      codeChallenge: challenge,
      codeChallengeMethod: "S256",
      prompt: "select_account"
    });

    response.writeHead(302, { Location: authUrl, "Cache-Control": "no-store" });
    response.end();
    return;
  }

  if (requestUrl.pathname === "/auth/callback" && request.method === "GET") {
    const state = requestUrl.searchParams.get("state");
    const code = requestUrl.searchParams.get("code");
    const authError = requestUrl.searchParams.get("error");

    if (authError) {
      sendHtml(response, 400, page("Sign-in was not completed", "<p>Return to the app and try again, or contact your IT administrator.</p>"));
      return;
    }

    const authRequest = state ? pendingAuth.get(state) : undefined;
    if (!authRequest || !code) {
      sendHtml(response, 400, page("Sign-in could not be verified", "<p>The sign-in request expired or could not be verified. Return to the app and try again.</p>"));
      return;
    }

    pendingAuth.delete(state);
    try {
      const result = await msal.acquireTokenByCode({
        code,
        scopes,
        redirectUri,
        codeVerifier: authRequest.verifier
      });

      if (!result?.accessToken) throw new Error("No access token returned.");
      accessToken = result.accessToken;
      sendHtml(
        response,
        200,
        page("Connected", `<p class="note">Mailbox access is active in this app process only.</p><p><a href="/organize">Organize recent messages</a></p>`)
      );
    } catch {
      sendHtml(response, 502, page("Sign-in failed", "<p>Microsoft sign-in did not complete. Check the app registration and redirect URI, then try again.</p>"));
    }
    return;
  }

  if (requestUrl.pathname === "/organize" && request.method === "GET") {
    if (!accessToken) {
      response.writeHead(302, { Location: "/", "Cache-Control": "no-store" });
      response.end();
      return;
    }

    try {
      const graphResponse = await fetch(graphMessagesUrl, {
        headers: { Authorization: `Bearer ${accessToken}` },
        signal: AbortSignal.timeout(15000)
      });

      if (!graphResponse.ok) {
        accessToken = undefined;
        sendHtml(response, 502, page("Mailbox request failed", "<p>Microsoft Graph did not return messages. Sign in again or ask your administrator to check delegated Mail.Read consent.</p>"));
        return;
      }

      const payload = await graphResponse.json();
      sendHtml(
        response,
        200,
        page(
          "Recent messages",
          `<p class="note">Showing up to 25 messages. Categories are simple subject-based suggestions; nothing in your mailbox has been changed.</p>${renderMessages(payload.value ?? [])}<p><a href="/">Home</a></p>`
        )
      );
    } catch {
      sendHtml(response, 502, page("Mailbox request failed", "<p>Could not reach Microsoft Graph. Check your connection and try again.</p>"));
    }
    return;
  }

  sendHtml(response, 404, page("Not found", "<p>The requested page does not exist.</p>"));
}

const server = createServer((request, response) => {
  void handleRequest(request, response).catch(() => {
    if (!response.headersSent) {
      sendHtml(response, 500, page("Unexpected error", "<p>The request could not be completed. Check the local terminal for configuration errors.</p>"));
    } else {
      response.destroy();
    }
  });
});

server.listen(port, host, () => {
  console.log(`Mail organizer listening at http://${host}:${port}`);
  console.log("Read-only mailbox access; tokens are kept in memory and are not logged.");
});

function stop() {
  accessToken = undefined;
  server.close(() => process.exit(0));
}

process.on("SIGINT", stop);
process.on("SIGTERM", stop);
