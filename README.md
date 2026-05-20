# local-office-mcp

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/Node.js-16%2B-green.svg)](https://nodejs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.0%2B-blue.svg)](https://www.typescriptlang.org)
[![MCP](https://img.shields.io/badge/MCP-compatible-purple.svg)](https://modelcontextprotocol.io)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D4.svg)](#requirements)
[![Last Commit](https://img.shields.io/github/last-commit/enoquefcd/local-office-mcp.svg)](https://github.com/enoquefcd/local-office-mcp/commits/main)

**Give your AI assistant native access to Outlook and Teams — entirely on your machine.**

**No Microsoft Graph API. No OAuth. No cloud relay.** Just your local Outlook COM interface and the Teams cache on disk.

```
"Summarize my inbox and flag anything that needs a reply today."
"Find a free 1-hour slot next week for a meeting with 3 people."
"Show me what the team was saying about the deployment last Thursday."
"Search my Teams messages for the database migration plan."
```

---

## How it works

```
┌─────────────────────────────────────────────────┐
│                  AI Assistant                   │
│              (Claude, Cursor, etc.)             │
└──────────────────┬──────────────────────────────┘
                   │ MCP (stdio)
┌──────────────────▼──────────────────────────────┐
│              local-office-mcp                   │
│               Node.js / TypeScript              │
├────────────────────┬────────────────────────────┤
│   Outlook tools    │      Teams tools            │
│                    │                             │
│  PowerShell COM    │  Python script              │
│  Outlook.Application  ccl_chromium_reader       │
│  (live, via IPC)   │  (local IndexedDB cache)    │
└────────────────────┴────────────────────────────┘
         │                        │
  ┌──────▼──────┐         ┌───────▼──────┐
  │   Outlook   │         │    Teams     │
  │  desktop    │         │  IndexedDB   │
  │    app      │         │   on disk    │
  └─────────────┘         └──────────────┘
```

- **Outlook** — spawns PowerShell to talk to `Outlook.Application` via COM. Same interface VBA macros use; no credentials needed.
- **Teams** — copies the local LevelDB cache to a temp folder (bypassing the file lock) and reads it with [`ccl_chromium_reader`](https://github.com/cclgroupltd/ccl_chrome_indexeddb). Teams stays open; data is never written back.

---

## Quick start

```powershell
# 1. clone & build
git clone https://github.com/enoquefcd/local-office-mcp.git
cd local-office-mcp
npm install && npm run build

# 2. Teams support (optional)
pip install git+https://github.com/cclgroupltd/ccl_chrome_indexeddb.git

# 3. wire up Claude Code (Windows)
claude mcp add local-office node -- "C:\path\to\local-office-mcp\dist\index.js"
```

<details>
<summary>Claude Desktop config</summary>

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "local-office": {
      "type": "stdio",
      "command": "node",
      "args": ["C:\\path\\to\\local-office-mcp\\dist\\index.js"]
    }
  }
}
```
</details>

<details>
<summary>Claude Code from WSL</summary>

The server must run as a Windows process (Outlook COM only works on Windows). Point WSL's Claude config at the Windows `node.exe`:

```json
{
  "mcpServers": {
    "local-office": {
      "command": "/mnt/c/Program Files/nodejs/node.exe",
      "args": ["/mnt/c/path/to/local-office-mcp/dist/index.js"]
    }
  }
}
```

> Use the Windows `node.exe` — not the WSL Node binary.
</details>

---

## Requirements

| Requirement | Outlook tools | Teams tools |
|------------|:---:|:---:|
| Windows 10/11 | ✅ | ✅ |
| Microsoft Outlook (desktop) | ✅ | — |
| Node.js 16+ | ✅ | ✅ |
| PowerShell 5+ | ✅ | — |
| Microsoft Teams (installed + opened at least once) | — | ✅ |
| Python 3.9+ | — | ✅ |
| `ccl_chromium_reader` | — | ✅ |

---

## Tools

<details>
<summary><strong>📧 Email</strong> — 12 tools</summary>

| Tool | Description |
|------|-------------|
| `get_inbox_emails` | Retrieve inbox emails |
| `get_sent_emails` | Retrieve sent emails |
| `get_draft_emails` | Retrieve draft emails |
| `get_email_by_id` | Get a specific email by ID |
| `search_inbox_emails` | Search inbox by keyword |
| `search_sent_emails` | Search sent folder by keyword |
| `search_draft_emails` | Search drafts by keyword |
| `mark_email_as_read` | Mark an email as read |
| `summarize_email` | Summarize a single email |
| `summarize_inbox` | Summarize recent inbox with priority grouping |
| `create_draft` | Create a draft (plain text or HTML) |
| `duplicate_email_as_draft` | Duplicate an existing email as a new draft |

</details>

<details>
<summary><strong>📅 Calendar</strong> — 8 tools</summary>

| Tool | Description |
|------|-------------|
| `list_events` | List events in a date range |
| `create_event_with_show_as` | Create event with Free/Busy/OutOfOffice status |
| `set_show_as` | Update Show As on an existing event |
| `update_event` | Edit subject, time, location, description |
| `delete_event` | Delete an event by ID |
| `find_free_slots` | Find open time slots between given hours |
| `get_attendee_status` | Check meeting response statuses |
| `get_calendars` | List available calendars |

</details>

<details>
<summary><strong>💬 Teams</strong> — 5 tools (read-only)</summary>

| Tool | Description |
|------|-------------|
| `teams_get_chats` | List recent chats and group conversations |
| `teams_get_messages` | Get messages from a specific chat |
| `teams_search_messages` | Search messages by keyword across all chats |
| `teams_get_channels` | List Teams channels |
| `teams_list_stores` | Debug: list IndexedDB object stores found |

Teams tools work with **classic Teams 1.x** (`%AppData%\Microsoft\Teams`) and **new Teams 2.x** (`MSTeams_8wekyb3d8bbwe` MSIX install). The right path is auto-detected.

</details>

---

## Limitations

- **Teams is read-only** — the local cache is never written
- **Teams member names** — display names aren't always in the local cache; GUIDs may appear
- **Teams history** — only conversations cached locally (recently visited) are available
- **Outlook** — requires the classic desktop app; does not work with new Outlook or Outlook web
- **Windows only** — COM automation and the Teams cache path are Windows-specific

---

## Project structure

```
local-office-mcp/
├── src/
│   ├── index.ts              # MCP server, tool registry
│   ├── outlook-manager.ts    # Outlook COM via PowerShell
│   ├── teams-manager.ts      # Teams Python subprocess bridge
│   ├── email-summarizer.ts   # Email search/summarization helpers
│   └── draft-generator.ts    # Draft generation helpers
├── scripts/
│   ├── read_teams_idb.py     # Reads Teams IndexedDB (Python)
│   └── requirements.txt      # Python dependencies
├── dist/                     # Compiled JS (generated, not committed)
├── package.json
└── tsconfig.json
```

---

## Contributing

PRs welcome. Keep it simple — no new runtime dependencies unless strictly necessary.

```bash
git checkout -b feat/my-feature
# hack hack hack
git commit -m "feat: add X"
# open a PR
```

Uses [Conventional Commits](https://www.conventionalcommits.org).

---

## License

[MIT](LICENSE)
