# Outlook MCP Server

This is a unified MCP (Model Context Protocol) server for comprehensive Microsoft Outlook integration. It operates the local Outlook client on Windows via COM and PowerShell, providing both **email management** and **calendar management** features. Its main advantage is fast deployment on Windows without complex security authentication.



## Installation

### 0. System Requirements
- Windows 10/11
- Microsoft Outlook installed and configured
- Node.js 16.0 or higher
- PowerShell 5.0 or higher

### 1. Install dependencies
```powershell
cd path\to\windows-outlook-mcp
npm install
```

### 2. Compile TypeScript
```powershell
npm run build
```

### 3. Configure your MCP client

#### Claude Desktop

Add to your Claude Desktop configuration file (`%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "type": "stdio",
      "command": "node",
      "args": ["path\\to\\windows-outlook-mcp\\dist\\index.js"],
      "env": {}
    }
  }
}
```

#### Claude Code (native Windows)

Add to `%USERPROFILE%\.claude\settings.json` (or run `claude mcp add`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["C:\\Users\\<your-username>\\windows-outlook-mcp\\dist\\index.js"]
    }
  }
}
```

#### Claude Code from WSL (Windows Subsystem for Linux)

Claude Code runs inside WSL but the MCP server must run on Windows (it needs the Outlook COM interface). Bridge them by pointing the WSL Claude config at the Windows Node binary and the Windows dist path.

Add to `~/.claude/settings.json` inside WSL:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "/mnt/c/Program Files/nodejs/node.exe",
      "args": ["C:\\Users\\<your-username>\\windows-outlook-mcp\\dist\\index.js"]
    }
  }
}
```

> **Note:** Use the Windows `node.exe` via `/mnt/c/...` — do **not** use the WSL Node binary. The server spawns PowerShell COM calls that only work from a Windows process.

## Usage Examples

#### Create an out-of-office event
```
Create a vacation event from 2025-12-24 to 2025-12-31 marked as OutOfOffice.
```

#### Find free time slots
```
Find free 30-minute slots in my calendar for next week between 9 AM and 5 PM.
```



## Development Notes

Project structure:
```
outlook/
├── src/
│   ├── index.ts              # Main server file
│   ├── outlook-manager.ts    # Outlook interface manager
│   ├── email-summarizer.ts   # Email summarization functionality
│   └── draft-generator.ts    # Draft generation functionality
├── dist/                     # Compiled output directory
├── package.json
├── tsconfig.json
└── README.md
```

To extend functionality, modify the relevant TypeScript files and recompile.

### Building from WSL

You can edit source files in WSL and build using the Windows npm:

```bash
cd /mnt/c/Users/<your-username>/windows-outlook-mcp
"/mnt/c/Program Files/nodejs/npm.cmd" run build
```

Or open a PowerShell terminal on the Windows side and run `npm run build` there.

## Available Tools & Features

### 📧 Email Management
- `get_inbox_emails` - Retrieve a list of inbox emails
- `get_sent_emails` - Retrieve a list of sent emails
- `get_draft_emails` - Retrieve a list of draft emails
- `get_email_by_id` - Get details of a specific email by ID
- `search_inbox_emails` - Search inbox emails by keyword
- `search_sent_emails` - Search sent emails by keyword
- `search_draft_emails` - Search draft emails by keyword
- `mark_email_as_read` - Mark an email as read

### 📝 Email Summarization
- `summarize_email` - Intelligently summarize a single email with priority detection
- `summarize_inbox` - Batch summarize inbox emails with priority grouping

### ✍️ Draft Management
- `create_draft` - Create a new email draft with recipients, subject, and body (supports HTML via `isHtml: true`)
- `duplicate_email_as_draft` - Duplicate an existing email as a draft (preserving formatting)

### 📅 Calendar Management
- `list_events` - List calendar events within a specified date range
- `create_event_with_show_as` - Create a calendar event with specific Show As status (Free/Busy/OutOfOffice/etc.)
- `set_show_as` - Set Show As status for an existing calendar event
- `update_event` - Update an existing calendar event (time, location, description, etc.)
- `delete_event` - Delete a calendar event by its ID
- `find_free_slots` - Find available time slots in the calendar with customizable work hours
- `get_attendee_status` - Check the response status of meeting attendees
- `get_calendars` - List all available calendars