import { google } from "googleapis";
import { readFileSync } from "fs";
export class SheetsClient {
    sheets;
    drive;
    authenticated;
    constructor() {
        const auth = this.getAuth();
        this.sheets = google.sheets({ version: "v4", auth });
        this.drive = google.drive({ version: "v3", auth });
        this.authenticated = !!process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
    }
    getAuth() {
        // Option 1: Service account JSON file (full read/write access)
        if (process.env.GOOGLE_SERVICE_ACCOUNT_KEY) {
            const keyPath = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
            const keyFile = JSON.parse(readFileSync(keyPath, "utf-8"));
            const auth = new google.auth.GoogleAuth({
                credentials: keyFile,
                scopes: [
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive.readonly",
                ],
            });
            return auth;
        }
        // Option 2: API key (read-only, public sheets only)
        if (process.env.GOOGLE_SHEETS_API_KEY) {
            return process.env.GOOGLE_SHEETS_API_KEY;
        }
        throw new Error("No Google authentication configured. Set GOOGLE_SERVICE_ACCOUNT_KEY (path to service account JSON) or GOOGLE_SHEETS_API_KEY.");
    }
    // ── Spreadsheet Operations ──
    async listSpreadsheets(maxResults = 20) {
        if (!this.authenticated) {
            throw new Error("Listing spreadsheets requires service account authentication.");
        }
        const res = await this.drive.files.list({
            q: "mimeType='application/vnd.google-apps.spreadsheet'",
            pageSize: maxResults,
            fields: "files(id, name, createdTime, modifiedTime, webViewLink)",
            orderBy: "modifiedTime desc",
        });
        return res.data.files || [];
    }
    async createSpreadsheet(title, sheetNames) {
        if (!this.authenticated) {
            throw new Error("Creating spreadsheets requires service account authentication.");
        }
        const sheets = sheetNames
            ? sheetNames.map((name, i) => ({ title: name, index: i }))
            : [{ title: "Sheet1", index: 0 }];
        const res = await this.sheets.spreadsheets.create({
            requestBody: {
                properties: { title },
                sheets: sheets.map((s) => ({ properties: s })),
            },
        });
        return res.data;
    }
    async getSpreadsheet(spreadsheetId) {
        const res = await this.sheets.spreadsheets.get({ spreadsheetId });
        return res.data;
    }
    // ── Sheet/Tab Operations ──
    async listSheets(spreadsheetId) {
        const spreadsheet = await this.getSpreadsheet(spreadsheetId);
        return (spreadsheet.sheets || []).map((s) => s.properties).filter(Boolean);
    }
    async addSheet(spreadsheetId, title) {
        if (!this.authenticated) {
            throw new Error("Adding sheets requires service account authentication.");
        }
        const res = await this.sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
                requests: [{ addSheet: { properties: { title } } }],
            },
        });
        return res.data.replies[0].addSheet.properties;
    }
    async deleteSheet(spreadsheetId, sheetId) {
        if (!this.authenticated) {
            throw new Error("Deleting sheets requires service account authentication.");
        }
        await this.sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
                requests: [{ deleteSheet: { sheetId } }],
            },
        });
    }
    // ── Data Operations ──
    async readRange(spreadsheetId, range) {
        const res = await this.sheets.spreadsheets.values.get({
            spreadsheetId,
            range,
            valueRenderOption: "FORMATTED_VALUE",
        });
        return (res.data.values || []);
    }
    async writeRange(spreadsheetId, range, values) {
        if (!this.authenticated) {
            throw new Error("Writing data requires service account authentication.");
        }
        const res = await this.sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption: "USER_ENTERED",
            requestBody: { values },
        });
        return res.data;
    }
    async appendRows(spreadsheetId, range, values) {
        if (!this.authenticated) {
            throw new Error("Appending data requires service account authentication.");
        }
        const res = await this.sheets.spreadsheets.values.append({
            spreadsheetId,
            range,
            valueInputOption: "USER_ENTERED",
            insertDataOption: "INSERT_ROWS",
            requestBody: { values },
        });
        return res.data;
    }
    async clearRange(spreadsheetId, range) {
        if (!this.authenticated) {
            throw new Error("Clearing data requires service account authentication.");
        }
        await this.sheets.spreadsheets.values.clear({
            spreadsheetId,
            range,
            requestBody: {},
        });
    }
    async getValues(spreadsheetId, sheetName) {
        return this.readRange(spreadsheetId, sheetName);
    }
    // ── Formatting ──
    async formatCells(spreadsheetId, sheetId, startRow, endRow, startCol, endCol, format) {
        if (!this.authenticated) {
            throw new Error("Formatting requires service account authentication.");
        }
        const cellFormat = {};
        const fields = [];
        if (format.bold !== undefined || format.italic !== undefined || format.fontSize !== undefined) {
            cellFormat.textFormat = {};
            if (format.bold !== undefined) {
                cellFormat.textFormat.bold = format.bold;
                fields.push("userEnteredFormat.textFormat.bold");
            }
            if (format.italic !== undefined) {
                cellFormat.textFormat.italic = format.italic;
                fields.push("userEnteredFormat.textFormat.italic");
            }
            if (format.fontSize !== undefined) {
                cellFormat.textFormat.fontSize = format.fontSize;
                fields.push("userEnteredFormat.textFormat.fontSize");
            }
        }
        if (format.foregroundColor) {
            cellFormat.textFormat = cellFormat.textFormat || {};
            cellFormat.textFormat.foregroundColorStyle = {
                rgbColor: {
                    red: format.foregroundColor.red || 0,
                    green: format.foregroundColor.green || 0,
                    blue: format.foregroundColor.blue || 0,
                },
            };
            fields.push("userEnteredFormat.textFormat.foregroundColorStyle");
        }
        if (format.backgroundColor) {
            cellFormat.backgroundColorStyle = {
                rgbColor: {
                    red: format.backgroundColor.red || 0,
                    green: format.backgroundColor.green || 0,
                    blue: format.backgroundColor.blue || 0,
                },
            };
            fields.push("userEnteredFormat.backgroundColorStyle");
        }
        if (format.numberFormat) {
            cellFormat.numberFormat = {
                type: format.numberFormat.type,
                pattern: format.numberFormat.pattern,
            };
            fields.push("userEnteredFormat.numberFormat");
        }
        if (format.horizontalAlignment) {
            cellFormat.horizontalAlignment = format.horizontalAlignment;
            fields.push("userEnteredFormat.horizontalAlignment");
        }
        await this.sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
                requests: [
                    {
                        repeatCell: {
                            range: {
                                sheetId,
                                startRowIndex: startRow,
                                endRowIndex: endRow,
                                startColumnIndex: startCol,
                                endColumnIndex: endCol,
                            },
                            cell: { userEnteredFormat: cellFormat },
                            fields: fields.join(","),
                        },
                    },
                ],
            },
        });
    }
    // ── Utility ──
    async searchSheets(spreadsheetId, query) {
        const spreadsheet = await this.getSpreadsheet(spreadsheetId);
        const results = [];
        const lowerQuery = query.toLowerCase();
        for (const sheet of spreadsheet.sheets || []) {
            const name = sheet.properties?.title;
            if (!name)
                continue;
            try {
                const values = await this.readRange(spreadsheetId, name);
                for (let r = 0; r < values.length; r++) {
                    for (let c = 0; c < values[r].length; c++) {
                        const cell = values[r][c];
                        if (cell && cell.toString().toLowerCase().includes(lowerQuery)) {
                            results.push({ sheet: name, row: r + 1, col: c + 1, value: cell });
                        }
                    }
                }
            }
            catch {
                // Skip sheets that can't be read
            }
        }
        return results;
    }
    async sheetOverview(spreadsheetId, sheetName, sampleRows = 5) {
        const values = await this.readRange(spreadsheetId, sheetName);
        const totalRows = values.length;
        const totalCols = values.reduce((max, row) => Math.max(max, row.length), 0);
        const headers = values.length > 0 ? values[0] : [];
        const sampleData = values.slice(1, 1 + sampleRows);
        return {
            title: sheetName,
            totalRows,
            totalCols,
            headers,
            sampleData,
        };
    }
}
