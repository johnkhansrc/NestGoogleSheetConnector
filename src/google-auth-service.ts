import {Inject, Injectable} from "@nestjs/common";
import GoogleSheetConnectorDto from "./dto/google-sheet-connector.dto";
import {JWT} from 'google-auth-library';
import {google} from 'googleapis';

@Injectable()
export class GoogleAuthService {

    private readonly SCOPE_SPREADSHEETS: string;
    constructor(@Inject('GOOGLE_SHEET_CONNECTOR') private readonly _credentials: GoogleSheetConnectorDto) {

        this.SCOPE_SPREADSHEETS = 'https://www.googleapis.com/auth/spreadsheets';
    }

    /**
     * Get the Google JWT client
     *
     * @returns {JWT}
     */
    public getClient() {

        const {
            client_email: email,
            private_key: key,
        } = this._credentials;

        return new google.auth.JWT(
            {
                email,
                key,
                scopes: [this.SCOPE_SPREADSHEETS],
            }
        );
    }
}
