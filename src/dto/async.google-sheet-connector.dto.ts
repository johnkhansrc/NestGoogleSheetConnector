import { ModuleMetadata } from '@nestjs/common';
import { FactoryProvider } from '@nestjs/common/interfaces/modules/provider.interface';
import GoogleSheetConnectorDto from "./google-sheet-connector.dto";

type AsyncGoogleSheetConnectorDto = Pick<ModuleMetadata, 'imports'> &
    Pick<FactoryProvider<GoogleSheetConnectorDto>, 'useFactory' | 'inject'>;

export default AsyncGoogleSheetConnectorDto;
