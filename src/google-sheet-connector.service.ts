import {Injectable} from "@nestjs/common";
import {GoogleAuthService} from "./google-auth-service";
import {JWT} from "google-auth-library";
import {google, sheets_v4} from "googleapis";
import {GaxiosPromise} from "googleapis-common";

@Injectable()
export class GoogleSheetConnectorService{

    private readonly _jwtClient: JWT;
    constructor(private _authService: GoogleAuthService) {

        this._jwtClient = this._authService.getClient();
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

        const range = this.readRange(spreadsheetId, cell);

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
                     values: any[][]): GaxiosPromise<sheets_v4.Schema$UpdateValuesResponse> {

        const sheets = this.getGoogleSheetConnect()

        return await sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption: 'USER_ENTER',
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
                 values: any[][]): GaxiosPromise<sheets_v4.Schema$AppendValuesResponse> {

        const sheets = this.getGoogleSheetConnect()

        return await sheets.spreadsheets.values.append({
            spreadsheetId,
            range,
            valueInputOption: 'USER_ENTER',
            requestBody: {
                values
            }
        });
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

    /**
     * Get Google Sheet Connection
     *
     * @returns sheets_v4.Sheets
     */
    public getGoogleSheetConnect(): sheets_v4.Sheets {

        return google.sheets({version: 'v4', auth: this._jwtClient});
    }
}
