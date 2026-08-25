<div align="center">

# Outlook Skill

![](./doc/banner.png)

**A small, secure command-line bridge to your Outlook mailbox.**

Search, list, and read email through Microsoft Graph with device-code sign-in,
silent token refresh, and JSON that is easy for people and agents to use.

</div>

## What it does

- Lists recent Inbox messages, including unread-only views
- Searches the whole mailbox using Microsoft Graph search syntax
- Retrieves a message with its full plain-text body
- Stores the refreshable sign-in cache locally with owner-only permissions
- Keeps email data on stdout as JSON and diagnostics on stderr

> [!IMPORTANT]
> This skill only supports basic Outlook email tasks at the moment as they match my personal needs.
> This can evolve with time, don't hesitate to contribute 🤝

## Setup

### 1. Install

```bash
bun install --frozen-lockfile
cp .env.example .env
```

### 2. Create a Microsoft OAuth client ID

This skill uses Microsoft's device-code flow. It is a public client: **do not
create or add a client secret.**

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Go to **App registrations** → **New registration**.
3. Give it a name such as `Outlook Skill`.
4. Under **Supported account types**, choose **Personal Microsoft accounts
   only**.
5. Select **Register**, then copy the **Application (client) ID** from the
   overview page.
6. Open **Authentication** → **Advanced settings** and enable **Allow public
   client flows**.
7. Open **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions**, then add:
   - `Mail.Read`
   - `User.Read`
   - `email`
   - `profile`

Put the client ID in `.env`:

```dotenv
OUTLOOK_OAUTH_CLIENT_ID=your-application-client-id
```

### 3. Sign in

```bash
bun scripts/outlook.ts login
```

Follow the device-code prompt in your browser. The resulting refreshable
session is saved at `~/.outlook-skill/msal-cache.json`; its directory is mode
`0700` and the cache is mode `0600`.

## Use it

```bash
# Recent Inbox messages
bun scripts/outlook.ts emails list
bun scripts/outlook.ts emails list --unread --limit 25

# Search the mailbox
bun scripts/outlook.ts emails search "from:hello@example.com"
bun scripts/outlook.ts emails search "subject:invoice" --limit 25

# Read a message returned by list or search
bun scripts/outlook.ts emails get "message-id"

# Get its attachments; file content is base64 in contentBytes
bun scripts/outlook.ts emails attachments "message-id"

# Remove the saved session from this machine
bun scripts/outlook.ts logout
```

`list` returns up to 100 messages and `search` up to 1,000. Search accepts
Microsoft Graph message-search syntax, including `from:`, `subject:`, and
`body:`.

### JSON output

Email commands print JSON, so they compose naturally with other tools:

```bash
bun scripts/outlook.ts emails list --limit 5 | jq '.[].subject'
```

```json
[
  {
    "id": "AQMk...",
    "subject": "Project update",
    "receivedDateTime": "2026-08-05T10:30:00Z",
    "isRead": false
  }
]
```

Use an item's `id` with `emails get` when you need its complete text content, or
`emails attachments` for complete attachment data. File attachment payloads are
base64-encoded in `contentBytes`.

## Development

```bash
bun run format:check
bun run test
bun run test:coverage
bunx --package typescript@5.9.3 tsc -p tsconfig.json
```

## For agents

If you are an agent, congratulations you found the right skill for your human's Outlook emails !

Install it from a checked-out repository:

```bash
npx skills add .
```

Or install a hosted copy by replacing the path with its repository URL:

```bash
npx skills add https://github.com/vincentkelleher/outlook-skill
```

The installer discovers `SKILL.md`. After installation, use the Outlook skill
when your human asks you to find, inspect, or summarize messages in their
Outlook mailbox.
