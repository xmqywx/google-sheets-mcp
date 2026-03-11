# Reddit Post

**Subreddits:** r/ClaudeAI, r/MachineLearning

---

**Title:** I built an MCP server that lets Claude read and write Google Sheets

---

**Body:**

Hey everyone,

I kept running into the same friction: I'd be mid-conversation with Claude, analyzing data or drafting reports, and I'd have to stop, export a CSV, paste it in, then manually copy results back into my spreadsheet. It got old fast.

So I built a Google Sheets MCP server that connects Claude directly to your spreadsheets.

**What it does:**

It exposes 14 tools to Claude via the Model Context Protocol, covering most of what you'd do in Sheets day-to-day:

- Read and write cell ranges
- Create, rename, and delete sheets
- Batch update rows (great for bulk data entry)
- Format cells and ranges
- Find and replace across a sheet
- List all sheets in a spreadsheet

Once it's set up, you can ask Claude things like "summarize the Q1 sales data in my spreadsheet" or "add these contacts to my leads sheet" and it just works.

**Setup:**

1. Install via npm:
   ```bash
   npm install -g google-sheets-mcp
   ```

2. Create a Google Cloud project and enable the Sheets API (the README walks through this step by step — it takes about 5 minutes).

3. Add this to your Claude Desktop config (`~/Library/Application Support/Claude/claude_desktop_config.json` on Mac):

   ```json
   {
     "mcpServers": {
       "google-sheets": {
         "command": "google-sheets-mcp",
         "env": {
           "GOOGLE_CLIENT_ID": "your_client_id",
           "GOOGLE_CLIENT_SECRET": "your_client_secret"
         }
       }
     }
   }
   ```

4. Restart Claude Desktop. First run will open an OAuth2 browser flow to authorize access to your Google account.

**Why I built this instead of using an existing solution:**

There isn't one, at least not a maintained npm package with full Sheets API coverage. There are some LangChain integrations and one-off scripts floating around, but nothing drop-in for Claude Desktop.

The project is open source. If you run into issues, find a bug, or want a tool added, open an issue or PR — I'm actively maintaining it.

GitHub: [github.com/your-handle/google-sheets-mcp]

Feedback very welcome, especially from anyone integrating it into heavier workflows.
