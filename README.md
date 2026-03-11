# Google Sheets MCP Server

[![npm version](https://img.shields.io/npm/v/@xmqywxkris/google-sheets-mcp.svg)](https://www.npmjs.com/package/@xmqywxkris/google-sheets-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A production-ready [Model Context Protocol](https://modelcontextprotocol.io) server that gives AI assistants full access to Google Sheets. Read, write, format, and manage spreadsheets directly from Claude, GPT, or any MCP-compatible client.

## What It Does

This MCP server wraps the Google Sheets API v4 and exposes 14 tools that let AI assistants:

- Create and manage spreadsheets
- Read and write cell data
- Append rows, clear ranges, search across sheets
- Apply formatting (bold, colors, number formats, alignment)
- Get sheet overviews with sample data

## Available Tools (14)

| Tool | Description |
|------|-------------|
| `list_spreadsheets` | List recent spreadsheets from Google Drive |
| `create_spreadsheet` | Create a new spreadsheet with custom sheet names |
| `get_spreadsheet` | Get spreadsheet metadata (title, sheets, locale, timezone) |
| `list_sheets` | List all sheets/tabs in a spreadsheet |
| `add_sheet` | Add a new sheet/tab |
| `delete_sheet` | Delete a sheet/tab by ID |
| `read_range` | Read cells from a range (e.g., `Sheet1!A1:D10`) |
| `write_range` | Write a 2D array of values to a range |
| `append_rows` | Append rows to the end of a sheet |
| `clear_range` | Clear all values in a range |
| `get_values` | Get all values from a sheet |
| `format_cells` | Apply formatting (bold, italic, colors, number format, alignment) |
| `search_sheets` | Search for text across all sheets (case-insensitive) |
| `sheet_overview` | Get row/column counts, headers, and sample data |

## Quick Start

### Install

```bash
npm install -g @xmqywxkris/google-sheets-mcp
```

Or use directly with npx:

```bash
npx -y @xmqywxkris/google-sheets-mcp
```

### Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "npx",
      "args": ["-y", "@xmqywxkris/google-sheets-mcp"],
      "env": {
        "GOOGLE_SERVICE_ACCOUNT_KEY": "/path/to/service-account.json"
      }
    }
  }
}
```

### Claude Code

```bash
claude mcp add google-sheets -- npx -y google-sheets-mcp

# Set the environment variable
export GOOGLE_SERVICE_ACCOUNT_KEY="/path/to/service-account.json"
```

## Authentication

### Option 1: Service Account (Recommended)

Full read/write access. Best for automation and production use.

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project (or select an existing one)
3. Enable the **Google Sheets API** and **Google Drive API**
4. Go to **Credentials** > **Create Credentials** > **Service Account**
5. Create the service account and download the JSON key file
6. Set the environment variable:

```bash
export GOOGLE_SERVICE_ACCOUNT_KEY="/path/to/service-account.json"
```

7. **Important**: Share your spreadsheets with the service account email (found in the JSON file as `client_email`) to grant access.

### Option 2: API Key (Read-Only)

Simpler setup, but only works with publicly shared spreadsheets.

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Enable the **Google Sheets API**
3. Go to **Credentials** > **Create Credentials** > **API Key**
4. Set the environment variable:

```bash
export GOOGLE_SHEETS_API_KEY="your-api-key"
```

## Configuration

| Variable | Required | Description |
|----------|----------|-------------|
| `GOOGLE_SERVICE_ACCOUNT_KEY` | One of these | Path to service account JSON key file (full access) |
| `GOOGLE_SHEETS_API_KEY` | is required | Google API key (read-only, public sheets only) |

## Usage Examples

**Data analysis:**
> "Read the sales data from my spreadsheet and summarize Q1 results"

**Report generation:**
> "Create a new spreadsheet called 'Monthly KPIs', add headers, and format them bold with a blue background"

**Data entry:**
> "Append these 50 customer records to the Customers sheet"

**Search and audit:**
> "Search all sheets for cells containing 'ERROR' or 'N/A'"

**Formatting:**
> "Bold the header row, set currency format on column C, and color negative values red"

**Sheet management:**
> "Add a 'Summary' tab and populate it with totals from other sheets"

## Development

```bash
git clone https://github.com/xmqywx/google-sheets-mcp.git
cd google-sheets-mcp
npm install
npm run dev
```

Build for production:

```bash
npm run build
npm start
```

## Cursor Plugin

Install directly in Cursor IDE:

1. Open Cursor → Settings → Plugins
2. Search "google-sheets-mcp"
3. Click Install
4. Configure your Google credentials in the MCP server settings

Or manually: clone this repo and open it in Cursor — the plugin auto-discovers.

## License

MIT
