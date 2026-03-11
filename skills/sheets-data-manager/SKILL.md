---
name: sheets-data-manager
description: Use this skill when the user wants to read, write, or manage data in Google Sheets. Handles cell operations, range queries, data appending, and spreadsheet creation.
---

# Sheets Data Manager

## When to Use
- User asks to read/write data from/to Google Sheets
- User wants to create a new spreadsheet
- User needs to append rows or update ranges
- User wants to search for data across sheets

## Steps
1. Ensure the google-sheets MCP server is connected
2. For reading: use read_range with the spreadsheet ID and range
3. For writing: use write_range with values as 2D array
4. For appending: use append_rows (auto-detects last row)
5. For new spreadsheets: use create_spreadsheet with a title
6. For searching: use search_sheets with a query string

## Available MCP Tools
- read_range: Read cells from a range
- write_range: Write values to a range
- append_rows: Add rows to the end of data
- clear_range: Clear cell contents
- get_values: Get all values from a sheet
- create_spreadsheet: Create new spreadsheet
- list_spreadsheets: List accessible spreadsheets
- get_spreadsheet: Get spreadsheet metadata
- search_sheets: Search for values across sheets
- sheet_overview: Get summary of a spreadsheet
