import {Injectable} from "@nestjs/common";
import {GoogleAuthService} from "./google-auth-service";
import {JWT} from "google-auth-library";
import {google, sheets_v4} from "googleapis";
import {GaxiosPromise} from "googleapis-common";

@Injectable()
export class GoogleSheetConnectorService{

    private readonly _jwtClient: JWT;
    private _loadedSpreadsheet: sheets_v4.Schema$Spreadsheet;
    private loadedSheet: sheets_v4.Schema$Sheet;
    constructor(private _authService: GoogleAuthService) {

        this._jwtClient = this._authService.getClient();
    }

    getLoadedSheet(): sheets_v4.Schema$Sheet {
        return this.loadedSheet;
    }

    async loadSheet(spreadsheetId: string, index: number): Promise<sheets_v4.Schema$Sheet> {
        if (!this._loadedSpreadsheet || this._loadedSpreadsheet.spreadsheetId !== spreadsheetId) {
            await this.loadSpreadSheet(spreadsheetId);
        }

        this.loadedSheet = await this.getSheetByLoadedSpreadSheetIndex(index);
        return this.loadedSheet;
    }

    getLoadedSpreadSheet(): sheets_v4.Schema$Spreadsheet {
        return this._loadedSpreadsheet;
    }

    async loadSpreadSheet(spreadsheetId: string): Promise<sheets_v4.Schema$Spreadsheet> {
        const sheets = this.getGoogleSheetConnect()
        const res = await sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            includeGridData: true,
        });

        this._loadedSpreadsheet = res.data;

        return res.data;
    }

    /**
     * Read a range of cells from a spreadsheet
     *
     * @param range
     *
     * @returns {Promise<any>}
     */
    async readRangeFromLoadedSheet(range: string): Promise<any[][]> {
        return this.readRangeFromLoadedSpreadSheet(this.getLoadedSheet().properties.title, range);
    }

    /**
     * Read a range of cells from a spreadsheet
     *
     * @param sheet
     * @param range
     *
     * @returns {Promise<any>}
     */
    async readRangeFromLoadedSpreadSheet(sheet, range: string): Promise<any[][]> {

        const spreadsheet = this.getLoadedSpreadSheet();
        const sheetIndex = spreadsheet.sheets.findIndex((s) => s.properties.title === sheet);
        const sheetData = spreadsheet.sheets[sheetIndex].data[0].rowData;
        const rangeData = range.split('!')[1];
        return this.getRangeData(sheetData, rangeData);
    }

    private getCharNumber(char: string): number {
        let result = 0;
        for (let i = 0; i < char.length; i++) {
            result = result * 26 + char.charCodeAt(i) - 64;
        }
        return result;
    }

    private getRangeCell(range: string): { row: number, column: number } {
        const row = parseInt(range.match(/\d+/)[0]) - 1;
        const column = this.getCharNumber(range.match(/[A-Z]+/)[0]) - 1;
        return {row, column};
    }
    private getRangeData(sheetData: sheets_v4.Schema$RowData[], range: string): any[][] {
        const rangeData = range.split(':');
        const start = this.getRangeCell(rangeData[0]);
        const end = this.getRangeCell(rangeData[1]);
        const data = [];
        for (let i = start.row; i <= end.row; i++) {
            const row = [];
            for (let j = start.column; j <= end.column; j++) {
                const value = sheetData[i].values[j].formattedValue;
                row.push(value);
            }
            data.push(row);
        }
        return data;
    }

    /**
     * Read a range of cells from a spreadsheet
     *
     * @param spreadsheetId
     * @param range
     *
     * @returns {Promise<any>}
     */
    async readRange(spreadsheetId: string, range: string): Promise<any[][]> {

        const sheets = this.getGoogleSheetConnect()
        const res = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range,
        });

        return res.data.values;
    }

    /**
     * Read a cell from a spreadsheet
     *
     * @param spreadsheetId
     * @param cell
     *
     * @returns {Promise<any>}
     */
    async readCell(spreadsheetId: string,cell: string): Promise<any> {

        const range = await this.readRange(spreadsheetId, cell);

        return range[0][0];
    }

    /**
     * Write a range of cells to a spreadsheet
     *
     * @param spreadsheetId
     * @param range
     * @param values
     *
     * @returns {GaxiosPromise<sheets_v4.Schema$UpdateValuesResponse>}
     */
    async writeRange(spreadsheetId: string,
                     range: string,
                     values: any[][],
                     valueInputOption = 'USER_ENTERED'): GaxiosPromise<sheets_v4.Schema$UpdateValuesResponse> {

        const sheets = this.getGoogleSheetConnect()

        return await sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption,
            requestBody: {
                values
            }
        });
    }

    /**
     * Write a cell to a spreadsheet
     *
     * @param spreadsheetId
     * @param cell
     * @param value
     *
     * @returns {GaxiosPromise<sheets_v4.Schema$UpdateValuesResponse>}
     */
    async writeCell(spreadsheetId: string,
                    cell: string,
                    value: any): GaxiosPromise<sheets_v4.Schema$UpdateValuesResponse> {

        return await this.writeRange(spreadsheetId, cell, [[value]]);
    }

    /**
     * Read a spreadsheet
     *
     * @param spreadsheetId
     *
     * @returns {Promise<sheets_v4.Schema$Sheet[]>}
     */
    async readAllSheet(spreadsheetId: string): Promise<sheets_v4.Schema$Sheet[]> {

        const sheets = this.getGoogleSheetConnect()
        const res = await sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            includeGridData: true,
        });

        return res.data.sheets;
    }

    /**
     * Add new row to a spreadsheet
     *
     * @param spreadsheetId
     * @param range
     * @param values
     *
     * @returns {GaxiosPromise<sheets_v4.Schema$AppendValuesResponse>}
     */
    async addRow(spreadsheetId: string,
                 range: string,
                 values: any[][],
                 valueInputOption = 'USER_ENTERED'): GaxiosPromise<sheets_v4.Schema$AppendValuesResponse> {

        const sheets = this.getGoogleSheetConnect()

        return await sheets.spreadsheets.values.append({
            spreadsheetId,
            range,
            valueInputOption,
            requestBody: {
                values
            }
        });
    }

    private async getSheetByLoadedSpreadSheetIndex(index: number): Promise<sheets_v4.Schema$Sheet> {
        return this._loadedSpreadsheet.sheets[index];
    }

    private async getSheetBySpreadSheetIndex(spreedsheatId: string, index: number): Promise<sheets_v4.Schema$Sheet> {
        return this.readAllSheet(spreedsheatId).then((sheets) => sheets[index]);
    }
    /**
     * Create a new spreadsheet
     *
     * @param title
     *
     * @returns Promise<string>
     */
    async createSpreadsheet(title: string): Promise<string> {

        const sheets = this.getGoogleSheetConnect()
        const res = await sheets.spreadsheets.create({
            requestBody: {
                properties: {
                    title: title
                }
            }
        });

        return res.data.spreadsheetId;
    }

    private async getSheetMergeMetadata(
        sheet: sheets_v4.Schema$Sheet,
    ) {
        // Find all rows groups locations from merged cells
        return sheet.merges.map((merge) => ({
            sheetId: merge.sheetId,
            start: merge.startRowIndex,
            end: merge.endRowIndex,
        }));
    }

    /**
     * Insert empty dimension
     *
     * @param spreadsheetId
     * @param sheetId
     * @param dimension
     * @param startIndex
     * @param endIndex
     */
    private async insertDimension(
        spreadsheetId: string,
        sheetId: number,
        dimension: string,
        startIndex: number,
        endIndex: number,
    ) {
        // Initialize request parameters.
        const request: sheets_v4.Schema$Request = {
            insertDimension: {
                range: {
                    sheetId,
                    dimension,
                    startIndex,
                    endIndex,
                },
                inheritFromBefore: false,
            },
        };

        // Update spreadsheet.
        return this.getGoogleSheetConnect()
            .spreadsheets.batchUpdate({
                spreadsheetId: spreadsheetId,
                requestBody: {
                    requests: [request],
                },
            });
    }

    /**
     * Insert empty row
     *
     * @param spreadsheetId,
     * @param sheetId
     * @param startIndex
     * @param endIndex
     */
    private async insertRow(
        spreadsheetId: string,
        sheetId: number,
        startIndex: number,
        endIndex: number,
    ) {
        return this.insertDimension(spreadsheetId, sheetId, 'ROWS', startIndex, endIndex);
    }

    /**
     * Merge cells in sheet
     *
     * @param spreadsheetId
     * @param sheetId
     * @param startRowIndex
     * @param endRowIndex
     * @param startColumnIndex
     * @param endColumnIndex
     */
    private async mergeCells(
        spreadsheetId: string,
        sheetId: number,
        startRowIndex: number,
        endRowIndex: number,
        startColumnIndex: number,
        endColumnIndex: number,
    ) {
        // Initialize request parameters.
        const request: sheets_v4.Schema$Request = {
            mergeCells: {
                range: {
                    sheetId: sheetId,
                    startRowIndex: startRowIndex,
                    endRowIndex: endRowIndex,
                    startColumnIndex: startColumnIndex,
                    endColumnIndex: endColumnIndex,
                },
            },
        };

        return await this.getGoogleSheetConnect()
            .spreadsheets.batchUpdate({
                spreadsheetId: spreadsheetId,
                requestBody: {
                    requests: [request],
                },
            });
    }

    async appendRow(spreedsheatId, sheet, rows, valueInputOption = 'USER_ENTERED', insertDataOption = 'INSERT_ROWS') {
        // Initialize request parameters
        const request = {
            spreadsheetId: spreedsheatId,
            range: `${sheet.properties.title}`,
            resource: {
                values: rows,
            },
            valueInputOption,
            insertDataOption,
        };

        return await this.getGoogleSheetConnect().spreadsheets.values.append(request);
    }

    /**
     * Get column value from a spreadsheet
     * @param spreedsheatId
     * @param sheet
     * @param columnKey
     * @private
     */
    private async getColumnValues(spreedsheatId, sheet, columnKey: string): Promise<string[]> {
        const values = await this.getGoogleSheetConnect()
            .spreadsheets.values.batchGet({
                spreadsheetId: spreedsheatId,
                ranges: [`${sheet.properties.title}!${columnKey}`],
            });

        return values.data.valueRanges[0].values.map((value: string[]) => value[0]);
    }

    /**
     * Get Google Sheet Connection
     *
     * @returns sheets_v4.Sheets
     */
    public getGoogleSheetConnect(): sheets_v4.Sheets {

        return google.sheets({version: 'v4', auth: this._jwtClient});
    }
}
