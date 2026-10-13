# Registering the Microsoft Entra app

Agent-facing, one time per user or machine. This produces an
`OUTLOOK_OAUTH_CLIENT_ID`, which is the only value the server cannot work
without. The full list of environment variables lives in the repository
README; the three that depend on this registration are the client id, the
tenant, and the read-only switch.

## 1. Create the app

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Go to **App registrations** → **New registration**.
3. Name it something the user will recognise later, such as `Outlook Skill`.
4. Under **Supported account types**, pick the one that matches the mailbox:
   - **Personal Microsoft accounts only** for outlook.com, hotmail.com and
     friends. This is the default, and needs no extra configuration.
   - **Accounts in this organizational directory only**, or an option that
     includes both, for a work or school mailbox. Those also need
     `OUTLOOK_OAUTH_TENANT` set to `organizations` or `common`, because the
     server assumes personal accounts otherwise.
5. Select **Register**, then copy the **Application (client) ID** from the
   overview page.

There is no redirect URI and no secret in this design: the device-code flow
signs a person in from a terminal and never needs either.

## 2. Allow the public client flow

**Authentication** → **Advanced settings** → under **Allow public client
flows**, select **Yes**, then **Save**. Without it, device-code sign-in fails
and the error comes from Microsoft, not from here.

## 3. Ask for the delegated permissions

**API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated
permissions**, then add:

| Permission        | Why                                                    |
| ----------------- | ------------------------------------------------------ |
| `Mail.ReadWrite`  | Reading the mailbox and writing drafts                 |
| `Mail.Read`       | The same reading, with no draft tool. One or the other |
| `User.Read`       | Finding out who is signed in                           |
| `email`,`profile` | The OpenID scopes the sign-in already asks for         |

Ask for `Mail.Read` **or** `Mail.ReadWrite`, not both:

- A server that may draft (`Mail.ReadWrite`) writes drafts, replies and
  forwards. It can still not send: the app needs `Mail.Send` for that and this
  project never asks for it.
- A server run with `OUTLOOK_MCP_READONLY=1` asks Microsoft for `Mail.Read` and
  exposes no draft tool at all, so `Mail.ReadWrite` is unnecessary there.

Granting `Mail.ReadWrite` is not consent. Delegated permissions are consented
by the signed-in user at first sign-in; in a work or school tenant where the
admin blocks consent, an admin has to grant them first.

## 4. Hand the client id over

```bash
export OUTLOOK_OAUTH_CLIENT_ID=<the Application (client) ID>
```

Or put it in a `.env` file in the directory the server starts from:

```dotenv
OUTLOOK_OAUTH_CLIENT_ID=your-application-client-id
```

## 5. Sign in once

```bash
node dist/index.js login
```

For an agent that must not block, use the background form and poll, as
`../SKILL.md` describes:

```bash
node dist/index.js login --background --json
```

The session is cached at `~/.outlook-skill/msal-cache.json` and shared by the
CLI and every MCP client on that machine.

## Switching an existing registration

A registration made before drafts existed will carry `Mail.Read`. Add
`Mail.ReadWrite` to it and sign in again: the token then carries the wider
scope, and the draft tools appear. Removing `Mail.ReadWrite` from the app while
running without `OUTLOOK_MCP_READONLY` makes every draft call fail with a
Microsoft consent error rather than a local one.
