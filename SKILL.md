---
name: outlook-skill
description: List, search, and read Outlook email through Microsoft Graph using cached MSAL authentication. Use when the user asks to find, inspect, or summarize messages in their Outlook mailbox.
compatibility: Requires Bun and OUTLOOK_OAUTH_CLIENT_ID.
---

# Outlook

Install the pinned dependencies once if `node_modules` is missing:

```bash
bun install --frozen-lockfile
```

Commands:

```bash
bun scripts/outlook.ts login
bun scripts/outlook.ts logout
bun scripts/outlook.ts emails list
bun scripts/outlook.ts emails list --unread --limit 25
bun scripts/outlook.ts emails get "message-id"
bun scripts/outlook.ts emails attachments "message-id"
bun scripts/outlook.ts emails search "query" --limit 25
bun scripts/outlook.ts emails search "from:person@example.com" --limit 100
```

Email commands write JSON to stdout. Diagnostics and login instructions use stderr. `list` returns Inbox metadata; `search` covers the mailbox and accepts Microsoft Graph message-search syntax such as `from:`, `subject:`, and `body:`; `get` returns one full body as plain text; `attachments` returns complete attachment data (file content is base64 in `contentBytes`). Search results are capped by Microsoft Graph at 1,000 messages.

If an email command reports no usable session, start device login in the background so the same process keeps polling while the user signs in:

```bash
mkdir -p /tmp/outlook-skill
nohup bun scripts/outlook.ts login >/tmp/outlook-skill/login.log 2>&1 & echo $! >/tmp/outlook-skill/login.pid
until grep -Eq 'https://(www\.)?microsoft\.com/(link|devicelogin)' /tmp/outlook-skill/login.log; do sleep 1; done
cat /tmp/outlook-skill/login.log
```

Show the user the link and code. Do not ask them to rerun login. After they confirm sign-in, wait for that process and retry the requested command:

```bash
pid=$(cat /tmp/outlook-skill/login.pid)
while kill -0 "$pid" 2>/dev/null; do sleep 1; done
bun scripts/outlook.ts emails list
```

MSAL stores its refreshable token cache at `~/.outlook-skill/msal-cache.json` with owner-only permissions. Later commands refresh it silently.
