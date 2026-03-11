---
name: sheets-formatter
description: Use this skill when the user wants to format cells, add styling, or change the appearance of Google Sheets data. Handles cell formatting like colors, fonts, and borders.
---

# Sheets Formatter

## When to Use
- User asks to format cells or add colors
- User wants headers styled or data highlighted
- User needs conditional-style formatting suggestions

## Steps
1. First read the current data to understand the layout
2. Use format_cells with the target range and format options
3. Common patterns:
   - Header row: bold, background color, center-aligned
   - Number columns: number format with decimals
   - Currency: currency format with symbol
   - Dates: date format
4. Apply formatting after all data is written

## Available MCP Tools
- format_cells: Apply formatting (colors, fonts, borders, alignment)
- read_range: Check current data layout before formatting
