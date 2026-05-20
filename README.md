# Windows Office MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/Node.js-16%2B-green.svg)](https://nodejs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.0%2B-blue.svg)](https://www.typescriptlang.org)
[![MCP](https://img.shields.io/badge/MCP-compatible-purple.svg)](https://modelcontextprotocol.io)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D4.svg)](https://www.microsoft.com/windows)

A local [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server that gives AI assistants native access to **Microsoft Outlook** (email + calendar) and **Microsoft Teams** — entirely on your machine, no cloud APIs, no OAuth, no data leaving your device.

- **Outlook** is accessed via COM automation (PowerShell) — same interface VBA macros use
- **Teams** is accessed by reading the local IndexedDB cache — works while Teams is running

> **Privacy-first:** all processing is local. Nothing is sent to external servers beyond your normal AI assistant traffic.

---

## Features

### 📧 Outlook — Email
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
| `summarize_inbox` | Summarize recent inbox |
| `create_draft` | Create a new draft (plain text or HTML) |
| `duplicate_email_as_draft` | Duplicate an email as a new draft |

### 📅 Outlook — Calendar
| Tool | Description |
|------|-------------|
| `list_events` | List events in a date range |
| `create_event_with_show_as` | Create event with Free/Busy/OutOfOffice status |
| `set_show_as` | Update Show As on an existing event |
| `update_event` | Edit subject, time, location, description |
| `delete_event` | Delete an event by ID |
| `find_free_slots` | Find open time slots in a date range |
| `get_attendee_status` | Check meeting response statuses |
| `get_calendars` | List available calendars |

### 💬 Microsoft Teams (read-only)
| Tool | Description |
|------|-------------|
| `teams_get_chats` | List recent chats and group conversations |
| `teams_get_messages` | Get messages from a specific chat |
| `teams_search_messages` | Search messages by keyword across all chats |
| `teams_get_channels` | List Teams channels |
| `teams_list_stores` | Debug: list IndexedDB object stores found |

---

## Requirements

### Outlook tools
- Windows 10 or 11
- Microsoft Outlook installed and signed in
- Node.js 16+
- PowerShell 5+

### Teams tools (additional)
- Microsoft Teams installed and launched at least once
- Python 3.9+
- [`ccl_chromium_reader`](https://github.com/cclgroupltd/ccl_chrome_indexeddb) library

---

## Installation

### 1. Clone and install

```powershell
git clone https://github.com/enoquefcd/windows-outlook-mcp.git
cd windows-outlook-mcp
npm install
npm run build
```

### 2. Install Python dependency (Teams tools only)

```powershell
pip install git+https://github.com/cclgroupltd/ccl_chrome_indexeddb.git
```

### 3. Configure your MCP client

#### Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "windows-office": {
      "type": "stdio",
      "command": "node",
      "args": ["C:\\path\\to\\windows-outlook-mcp\\dist\\index.js"]
    }
  }
}
```

#### Claude Code (native Windows)

```powershell
claude mcp add windows-office node -- "C:\path\to\windows-outlook-mcp\dist\index.js"
```

Or edit `%USERPROFILE%\.claude\settings.json`:

```json
{
  "mcpServers": {
    "windows-office": {
      "command": "node",
      "args": ["C:\\path\\to\\windows-outlook-mcp\\dist\\index.js"]
    }
  }
}
```

#### Claude Code from WSL

The MCP server must run on Windows (Outlook COM only works from a Windows process). Point your WSL Claude config at the Windows Node binary:

```json
{
  "mcpServers": {
    "windows-office": {
      "command": "/mnt/c/Program Files/nodejs/node.exe",
      "args": ["/mnt/c/path/to/windows-outlook-mcp/dist/index.js"]
    }
  }
}
```

> Use the Windows `node.exe` — not the WSL Node binary. The server spawns PowerShell processes that only work from a Windows process.

---

## Usage examples

```
Summarize my inbox and flag anything urgent.

Schedule a vacation from Dec 24 to Jan 2, marked as Out of Office.

Find a free 1-hour slot next week between 9 AM and 6 PM.

Show me recent messages in my ARCH Team chat.

Search my Teams messages for "deployment plan".
```

---

## How it works

### Outlook (COM automation)

The server spawns PowerShell processes that interact with Outlook via its COM object model (`Outlook.Application`). This is the same interface used by VBA macros and the Windows Scripting Host — no credentials, no tokens, just local IPC with the running Outlook process.

### Teams (local IndexedDB cache)

Teams stores all chat data in a Chrome IndexedDB (LevelDB) database on disk. The server:

1. Copies the LevelDB folder to a temp directory (to avoid the file lock held by the running Teams process)
2. Reads the copy using [`ccl_chromium_reader`](https://github.com/cclgroupltd/ccl_chrome_indexeddb), a well-known forensic analysis library
3. Parses the `conversation-manager` and `replychain-manager` databases to extract chats and messages
4. Cleans up the temp copy

Teams does **not** need to be closed. The copy-then-read approach bypasses the lock.

Both classic Teams (1.x, `%AppData%\Microsoft\Teams`) and new Teams 2.x (MSIX/Store, `%LocalAppData%\Packages\MSTeams_8wekyb3d8bbwe`) are auto-detected, with new Teams taking priority.

---

## Project structure

```
windows-outlook-mcp/
├── src/
│   ├── index.ts              # MCP server entry point, tool registry
│   ├── outlook-manager.ts    # Outlook COM automation
│   ├── teams-manager.ts      # Teams Python subprocess bridge
│   ├── email-summarizer.ts   # Email search/summarization helpers
│   └── draft-generator.ts    # Draft generation helpers
├── scripts/
│   ├── read_teams_idb.py     # Python script: reads Teams IndexedDB
│   └── requirements.txt      # Python dependencies
├── dist/                     # Compiled JavaScript (generated)
├── package.json
├── tsconfig.json
└── README.md
```

---

## Limitations

| Feature | Limitation |
|---------|------------|
| Teams | Read-only — the local cache is never written back |
| Teams members | Display names not always available in local cache (UUIDs may appear) |
| Teams channels | Only channels cached locally (recently visited) are visible |
| Outlook | Requires Outlook desktop app — does not work with Outlook web or new Outlook |
| Platform | Windows only |

---

## Contributing

PRs welcome. Keep the architecture simple — no new runtime dependencies unless strictly necessary.

1. Fork the repo
2. Create a feature branch: `git checkout -b feat/my-feature`
3. Commit using [Conventional Commits](https://www.conventionalcommits.org): `feat: add X`
4. Open a pull request

---

## License

MIT — see [LICENSE](LICENSE).
