# I Built an MCP Server That Lets Claude Control My Google Sheets

Every week I was doing the same tedious dance.

Copy sales numbers from one sheet. Paste into the report sheet. Format the header row. Run a VLOOKUP. Forget the exact formula, Google it. Apply conditional formatting. Send a Slack message saying "report's ready."

Forty minutes of my life, every single week, for data that could have moved itself.

I'm a developer. I automate things for a living. Yet here I was, manually babysitting spreadsheets like it was 2008.

So I built a fix.

---

## What Is MCP?

MCP (Model Context Protocol) is Anthropic's open standard that lets AI models like Claude talk directly to external tools and APIs. Instead of copy-pasting data into a chat window, you give Claude a set of *tools* it can call — and it figures out how to use them to complete your request.

Think of it as giving Claude hands instead of just a voice.

I'd already built an MCP server for Shopify. Google Sheets was the obvious next target.

---

## The Google Sheets MCP Server: 14 Tools

The package is called `@xmqywxkris/google-sheets-mcp` on npm. It wraps the official Google Sheets API with 14 tools Claude can call directly:

| Tool | What It Does |
|------|-------------|
| `read_sheet` | Read all data from a sheet tab |
| `read_range` | Read a specific cell range (e.g., `A1:D20`) |
| `write_range` | Write data to a specific range |
| `append_rows` | Append rows to the end of a sheet |
| `clear_range` | Clear a cell range |
| `create_spreadsheet` | Create a brand-new spreadsheet |
| `create_sheet` | Add a new tab to an existing spreadsheet |
| `delete_sheet` | Remove a sheet tab |
| `get_spreadsheet_info` | Get metadata, sheet names, dimensions |
| `format_range` | Apply bold, colors, borders, number formats |
| `merge_cells` | Merge a range of cells |
| `freeze_rows_columns` | Freeze header rows/columns |
| `search_sheet` | Search for a value across the entire sheet |
| `copy_sheet` | Duplicate a sheet tab within a spreadsheet |

Full spreadsheet control. No Python scripts. No App Script. Just Claude.

---

## Real Usage: What This Actually Looks Like

Here are four scenarios I use regularly now.

### 1. Pulling Data for a Quick Summary

**Me:** "Read my sales sheet and tell me which product had the highest revenue last month."

**Claude:** *(calls `read_sheet` on my Sales tab, scans the data)*
"Looking at your data, 'Widget Pro' led with $12,400 in February — about 34% of total revenue. Runner-up was 'Widget Lite' at $8,100."

No formulas. No pivot tables. Just an answer.

---

### 2. Updating a Report Automatically

**Me:** "Append these new signups to my CRM sheet: Alice Wong, alice@example.com, Enterprise plan, March 11. Bob Lee, bob@example.com, Starter plan, March 11."

**Claude:** *(calls `append_rows` with the structured data)*
"Done. I've added both rows to your CRM sheet. Alice is in row 47 and Bob is in row 48."

---

### 3. Formatting a Sheet for Presentation

**Me:** "Make the header row in my Q1 report bold with a dark blue background and white text, then freeze it so it stays visible when scrolling."

**Claude:** *(calls `format_range` then `freeze_rows_columns`)*
"Done. Row 1 is now formatted and frozen. Your report is presentation-ready."

That used to take me five minutes of clicking around the toolbar.

---

### 4. Finding Data Across a Large Sheet

**Me:** "Search my inventory sheet for any row containing 'discontinued' and tell me what products and quantities are affected."

**Claude:** *(calls `search_sheet`)*
"Found 3 rows flagged 'discontinued': SKU-441 (12 units), SKU-889 (0 units), SKU-1203 (47 units). Total: 59 units of discontinued stock."

---

## Setup

### Step 1: Install

```bash
npm install -g @xmqywxkris/google-sheets-mcp
```

### Step 2: Google Cloud Service Account

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a project (or use an existing one)
3. Enable the **Google Sheets API**
4. Create a **Service Account** under IAM & Admin
5. Generate a JSON key and download it
6. Share your Google Sheets with the service account email (treat it like a collaborator)

### Step 3: Configure Claude Desktop

Add this to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "google-sheets-mcp",
      "env": {
        "GOOGLE_SERVICE_ACCOUNT_KEY_PATH": "/absolute/path/to/your/service-account-key.json"
      }
    }
  }
}
```

Restart Claude Desktop. That's it.

---

## Why TypeScript + Official Client

The server is built in TypeScript with the `@modelcontextprotocol/sdk` and Google's official `googleapis` client. No third-party wrappers, no unofficial endpoints — just the standard Google Sheets API v4 with proper OAuth scoping via service account.

This means:
- Type-safe tool definitions Claude can reason about clearly
- Rate limiting handled by the official client
- Works with any Google Workspace or personal Google account
- Zero mystery dependencies

---

## What's Next

I'm planning to add batch operations (update multiple ranges in one call) and a `copy_to_spreadsheet` tool for cross-file operations. If you have a use case I haven't covered, open an issue on GitHub.

---

## Links

- **npm:** `npm install @xmqywxkris/google-sheets-mcp`
- **GitHub:** [github.com/xmqywx/google-sheets-mcp](https://github.com/xmqywx/google-sheets-mcp)
- **Need a custom MCP server or automation?** I'm available on [Upwork](https://www.upwork.com) — search for "MCP server developer" or message me directly. I build AI integrations, automation pipelines, and custom tooling for teams who are tired of doing things manually.

---

If you're still copying data between spreadsheets by hand, you don't have to be.

#mcp #googlesheets #ai #claude #automation #typescript
