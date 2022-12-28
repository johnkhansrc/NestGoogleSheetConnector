import {Inject, Injectable} from "@nestjs/common";
import GoogleSheetConnectorDto from "./dto/google-sheet-connector.dto";

@Injectable()
export class GoogleSheetConnectorService{
    constructor(@Inject('GOOGLE_SHEET_CONNECTOR') private readonly options: GoogleSheetConnectorDto) {
    }

    test() {
        console.log('______________TEST______________');
        console.log(this.options);
        console.log('______________TEST______________');
        return 'test ' + this.options.foo + ' ' + this.options.bar;
    };
}
