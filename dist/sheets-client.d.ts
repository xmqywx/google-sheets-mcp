import { sheets_v4, drive_v3 } from "googleapis";
export declare class SheetsClient {
    private sheets;
    private drive;
    private authenticated;
    constructor();
    private getAuth;
    listSpreadsheets(maxResults?: number): Promise<drive_v3.Schema$File[]>;
    createSpreadsheet(title: string, sheetNames?: string[]): Promise<sheets_v4.Schema$Spreadsheet>;
    getSpreadsheet(spreadsheetId: string): Promise<sheets_v4.Schema$Spreadsheet>;
    listSheets(spreadsheetId: string): Promise<sheets_v4.Schema$SheetProperties[]>;
    addSheet(spreadsheetId: string, title: string): Promise<sheets_v4.Schema$SheetProperties>;
    deleteSheet(spreadsheetId: string, sheetId: number): Promise<void>;
    readRange(spreadsheetId: string, range: string): Promise<string[][]>;
    writeRange(spreadsheetId: string, range: string, values: string[][]): Promise<sheets_v4.Schema$UpdateValuesResponse>;
    appendRows(spreadsheetId: string, range: string, values: string[][]): Promise<sheets_v4.Schema$AppendValuesResponse>;
    clearRange(spreadsheetId: string, range: string): Promise<void>;
    getValues(spreadsheetId: string, sheetName: string): Promise<string[][]>;
    formatCells(spreadsheetId: string, sheetId: number, startRow: number, endRow: number, startCol: number, endCol: number, format: {
        bold?: boolean;
        italic?: boolean;
        fontSize?: number;
        foregroundColor?: {
            red?: number;
            green?: number;
            blue?: number;
        };
        backgroundColor?: {
            red?: number;
            green?: number;
            blue?: number;
        };
        numberFormat?: {
            type: string;
            pattern: string;
        };
        horizontalAlignment?: string;
    }): Promise<void>;
    searchSheets(spreadsheetId: string, query: string): Promise<{
        sheet: string;
        row: number;
        col: number;
        value: string;
    }[]>;
    sheetOverview(spreadsheetId: string, sheetName: string, sampleRows?: number): Promise<{
        title: string;
        totalRows: number;
        totalCols: number;
        headers: string[];
        sampleData: string[][];
    }>;
}
