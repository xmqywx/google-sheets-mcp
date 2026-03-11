#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { SheetsClient } from "./sheets-client.js";

const server = new McpServer({
  name: "google-sheets-mcp",
  version: "1.0.0",
});

let client: SheetsClient;

function getClient(): SheetsClient {
  if (!client) {
    client = new SheetsClient();
  }
  return client;
}

// ══════════════════════════════════════════════════════════
// Spreadsheet Tools
// ══════════════════════════════════════════════════════════

server.tool(
  "list_spreadsheets",
  "List recent spreadsheets from Google Drive (requires service account)",
  {
    maxResults: z.number().min(1).max(100).default(20).describe("Maximum number of spreadsheets to return"),
  },
  async ({ maxResults }) => {
    try {
      const files = await getClient().listSpreadsheets(maxResults);
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(files, null, 2),
          },
        ],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

server.tool(
  "create_spreadsheet",
  "Create a new Google Spreadsheet",
  {
    title: z.string().describe("Title for the new spreadsheet"),
    sheetNames: z.array(z.string()).optional().describe("Names for initial sheets/tabs (default: ['Sheet1'])"),
  },
  async ({ title, sheetNames }) => {
    try {
      const spreadsheet = await getClient().createSpreadsheet(title, sheetNames);
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              {
                spreadsheetId: spreadsheet.spreadsheetId,
                title: spreadsheet.properties?.title,
                url: spreadsheet.spreadsheetUrl,
                sheets: spreadsheet.sheets?.map((s) => s.properties?.title),
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

server.tool(
  "get_spreadsheet",
  "Get spreadsheet metadata including title, sheets, and locale",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
  },
  async ({ spreadsheetId }) => {
    try {
      const spreadsheet = await getClient().getSpreadsheet(spreadsheetId);
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              {
                spreadsheetId: spreadsheet.spreadsheetId,
                title: spreadsheet.properties?.title,
                locale: spreadsheet.properties?.locale,
                timeZone: spreadsheet.properties?.timeZone,
                url: spreadsheet.spreadsheetUrl,
                sheets: spreadsheet.sheets?.map((s) => ({
                  sheetId: s.properties?.sheetId,
                  title: s.properties?.title,
                  index: s.properties?.index,
                  rowCount: s.properties?.gridProperties?.rowCount,
                  columnCount: s.properties?.gridProperties?.columnCount,
                })),
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

// ══════════════════════════════════════════════════════════
// Sheet/Tab Tools
// ══════════════════════════════════════════════════════════

server.tool(
  "list_sheets",
  "List all sheets/tabs in a spreadsheet",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
  },
  async ({ spreadsheetId }) => {
    try {
      const sheets = await getClient().listSheets(spreadsheetId);
      return {
        content: [{ type: "text" as const, text: JSON.stringify(sheets, null, 2) }],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

server.tool(
  "add_sheet",
  "Add a new sheet/tab to a spreadsheet",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
    title: z.string().describe("Name for the new sheet/tab"),
  },
  async ({ spreadsheetId, title }) => {
    try {
      const sheet = await getClient().addSheet(spreadsheetId, title);
      return {
        content: [{ type: "text" as const, text: JSON.stringify(sheet, null, 2) }],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

server.tool(
  "delete_sheet",
  "Delete a sheet/tab from a spreadsheet",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
    sheetId: z.number().describe("The numeric ID of the sheet to delete (from list_sheets)"),
  },
  async ({ spreadsheetId, sheetId }) => {
    try {
      await getClient().deleteSheet(spreadsheetId, sheetId);
      return {
        content: [{ type: "text" as const, text: `Sheet ${sheetId} deleted successfully.` }],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

// ══════════════════════════════════════════════════════════
// Data Tools
// ══════════════════════════════════════════════════════════

server.tool(
  "read_range",
  "Read cells from a range (e.g., 'Sheet1!A1:D10')",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
    range: z.string().describe("The A1 notation range to read (e.g., 'Sheet1!A1:D10')"),
  },
  async ({ spreadsheetId, range }) => {
    try {
      const values = await getClient().readRange(spreadsheetId, range);
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify({ range, rowCount: values.length, values }, null, 2),
          },
        ],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

server.tool(
  "write_range",
  "Write data to a range in a spreadsheet",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
    range: z.string().describe("The A1 notation range to write to (e.g., 'Sheet1!A1:D3')"),
    values: z
      .array(z.array(z.string()))
      .describe("2D array of values to write (rows of cells)"),
  },
  async ({ spreadsheetId, range, values }) => {
    try {
      const result = await getClient().writeRange(spreadsheetId, range, values);
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              {
                updatedRange: result.updatedRange,
                updatedRows: result.updatedRows,
                updatedColumns: result.updatedColumns,
                updatedCells: result.updatedCells,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

server.tool(
  "append_rows",
  "Append rows to the end of a sheet",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
    range: z.string().describe("The sheet name or range to append to (e.g., 'Sheet1')"),
    values: z
      .array(z.array(z.string()))
      .describe("2D array of row data to append"),
  },
  async ({ spreadsheetId, range, values }) => {
    try {
      const result = await getClient().appendRows(spreadsheetId, range, values);
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              {
                updatedRange: result.updates?.updatedRange,
                updatedRows: result.updates?.updatedRows,
                updatedCells: result.updates?.updatedCells,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

server.tool(
  "clear_range",
  "Clear all values from a range of cells",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
    range: z.string().describe("The A1 notation range to clear (e.g., 'Sheet1!A1:D10')"),
  },
  async ({ spreadsheetId, range }) => {
    try {
      await getClient().clearRange(spreadsheetId, range);
      return {
        content: [{ type: "text" as const, text: `Range ${range} cleared successfully.` }],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

server.tool(
  "get_values",
  "Get all values from a sheet",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
    sheetName: z.string().describe("The name of the sheet/tab (e.g., 'Sheet1')"),
  },
  async ({ spreadsheetId, sheetName }) => {
    try {
      const values = await getClient().getValues(spreadsheetId, sheetName);
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify({ sheet: sheetName, rowCount: values.length, values }, null, 2),
          },
        ],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

// ══════════════════════════════════════════════════════════
// Formatting Tools
// ══════════════════════════════════════════════════════════

server.tool(
  "format_cells",
  "Apply formatting to a range of cells (bold, color, number format, alignment)",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
    sheetId: z.number().describe("The numeric sheet ID (from list_sheets)"),
    startRow: z.number().describe("Start row index (0-based)"),
    endRow: z.number().describe("End row index (exclusive, 0-based)"),
    startCol: z.number().describe("Start column index (0-based)"),
    endCol: z.number().describe("End column index (exclusive, 0-based)"),
    bold: z.boolean().optional().describe("Set text bold"),
    italic: z.boolean().optional().describe("Set text italic"),
    fontSize: z.number().optional().describe("Font size in points"),
    foregroundColor: z
      .object({ red: z.number().min(0).max(1), green: z.number().min(0).max(1), blue: z.number().min(0).max(1) })
      .optional()
      .describe("Text color as RGB (0-1 range)"),
    backgroundColor: z
      .object({ red: z.number().min(0).max(1), green: z.number().min(0).max(1), blue: z.number().min(0).max(1) })
      .optional()
      .describe("Background color as RGB (0-1 range)"),
    numberFormat: z
      .object({
        type: z.string().describe("Format type: TEXT, NUMBER, PERCENT, CURRENCY, DATE, TIME, SCIENTIFIC"),
        pattern: z.string().describe("Format pattern (e.g., '#,##0.00', 'yyyy-mm-dd')"),
      })
      .optional()
      .describe("Number format"),
    horizontalAlignment: z
      .enum(["LEFT", "CENTER", "RIGHT"])
      .optional()
      .describe("Horizontal alignment"),
  },
  async ({
    spreadsheetId,
    sheetId,
    startRow,
    endRow,
    startCol,
    endCol,
    bold,
    italic,
    fontSize,
    foregroundColor,
    backgroundColor,
    numberFormat,
    horizontalAlignment,
  }) => {
    try {
      await getClient().formatCells(spreadsheetId, sheetId, startRow, endRow, startCol, endCol, {
        bold,
        italic,
        fontSize,
        foregroundColor,
        backgroundColor,
        numberFormat,
        horizontalAlignment,
      });
      return {
        content: [{ type: "text" as const, text: "Formatting applied successfully." }],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

// ══════════════════════════════════════════════════════════
// Utility Tools
// ══════════════════════════════════════════════════════════

server.tool(
  "search_sheets",
  "Search for text across all sheets in a spreadsheet",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
    query: z.string().describe("Text to search for (case-insensitive)"),
  },
  async ({ spreadsheetId, query }) => {
    try {
      const results = await getClient().searchSheets(spreadsheetId, query);
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify({ query, matchCount: results.length, matches: results }, null, 2),
          },
        ],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

server.tool(
  "sheet_overview",
  "Get an overview of a sheet with row/column counts and sample data",
  {
    spreadsheetId: z.string().describe("The ID of the spreadsheet"),
    sheetName: z.string().describe("The name of the sheet/tab"),
    sampleRows: z.number().min(1).max(50).default(5).describe("Number of sample data rows to include"),
  },
  async ({ spreadsheetId, sheetName, sampleRows }) => {
    try {
      const overview = await getClient().sheetOverview(spreadsheetId, sheetName, sampleRows);
      return {
        content: [{ type: "text" as const, text: JSON.stringify(overview, null, 2) }],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: "text" as const, text: `Error: ${message}` }], isError: true };
    }
  }
);

// ══════════════════════════════════════════════════════════
// Start Server
// ══════════════════════════════════════════════════════════

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Google Sheets MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
