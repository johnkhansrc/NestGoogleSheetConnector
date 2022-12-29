import {DynamicModule, Module} from "@nestjs/common";
import GoogleSheetConnectorDto from "./dto/google-sheet-connector.dto";
import {GoogleSheetConnectorService} from "./google-sheet-connector.service";
import {GoogleAuthService} from "./google-auth-service";

@Module({})
export class GoogleSheetModule {
    static register(options: GoogleSheetConnectorDto): DynamicModule {
        return {
            module: GoogleSheetModule,
            providers: [
                {
                    provide: 'GOOGLE_SHEET_CONNECTOR',
                    useValue: options,
                },
                GoogleSheetConnectorService,
                GoogleAuthService
            ],
            imports: [],
            exports: [GoogleSheetConnectorService, GoogleAuthService],
        }
    }
}
