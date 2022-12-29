# NestGoogleSheetConnector
A Google Sheet integration for NestJs.


# Installation

To use this module you will need:
- create a Google Cloud service account
- install the module
- provide the module with the credentials of the google cloud service account
- provide permissions to your service account on the spreadsheets you want to manipulate

## Create a Google Cloud service account

1. Access the  [Google APIs Console](https://console.developers.google.com/)  **while logged into your Google account**.
2. Create a **new project** and give it a name.
[Create a new project](https://robocorp.com/docs/static/build/development-guide/google-sheets/interacting-with-google-sheets/console-create-project.png)
3. Click on  `ENABLE APIS AND SERVICES`.
[Enable Google Sheets API](https://robocorp.com/docs/static/build/development-guide/google-sheets/interacting-with-google-sheets/enable-google-sheets-api.png)
4. Find and enable the `Google Sheet API`.
5. Create new credentials to the `Google Sheets API`. Select `Other UI` from the dropdown and select `Application Data`. Then click on the `What credentials do I need?` button.![Create credentials](https://robocorp.com/docs/static/build/development-guide/google-sheets/interacting-with-google-sheets/create-credentials.png)
6. On the next screen, choose a name for your service account, assign it a role of `Project`->`Editor`, and click `Continue`.![Create credentials step 2](https://robocorp.com/docs/static/build/development-guide/google-sheets/interacting-with-google-sheets/create-credentials-step2.png)
7. The credentials JSON file will be downloaded by your browser.

   > The credentials file allows anyone to access your cloud resources, so you should store it securely.  [More information from Google](https://cloud.google.com/iam/docs/understanding-service-accounts#managing_service_account_keys).

8.  Find the downloaded file and rename it to  `service_account.json`.

## Install the module
```
npm install --save xxxxxxxxx
```

## Provide the module with the credentials of the google cloud service account
```ts
import { GoogleSheetModule } from 'nest-google-sheet-connector';

@Module({  
  imports: [  
    GoogleSheetModule.register(credentials), // credentials is a JSON object downloaded from Google Cloud Platform  
  ],  
  controllers: [AppController],  
  providers: [AppService],  
})  
export class AppModule {}
```

## Provide permissions to your service account on the spreadsheets you want to manipulate
1.  Create or select an existing Google Sheet.
2.  Open the  `service_account.json`  file and find the  `client_email`  property.
3.  Click on the  `Share`  button in the top right, and add the email address of the service account as an editor.![Add user as an editor](https://robocorp.com/docs/static/build/development-guide/google-sheets/interacting-with-google-sheets/spreadsheet-add-user-as-editor.png)

>If you want only to allow the account read access to the spreadsheet, assign it the `Viewer` role instead.

4. Take note of the ID of the Google Sheet document, which is contained in its URL, after the  `/d`  element. So, for example, if the URL of your document is  `https://docs.google.com/spreadsheets/d/1234567890123abcf/edit#gid=0`, the ID will be  `1234567890123abcf`.

# Documentation
## Inject the service into a constructor
```ts
import { GoogleSheetConnectorService } from '../../lib/nest-google-sheet-connector';

@Injectable()  
export class AppService {  
  constructor(private googleSheetConnectorService: GoogleSheetConnectorService) {}  
}
```
This service contain google sheet providers methods
## Features
['feature'](ressources/img.png)
## Troubleshooting
Google Cloud requires an SSL certificate for API requests.
Here is how to certify your local environment under NestJs:

1. Get your certificate
 ```
 openssl req -x509 -out localhost.crt -keyout localhost.key \
  -newkey rsa:2048 -nodes -sha256 \
  -subj '/CN=localhost' -extensions EXT -config <( \
   printf "[dn]\nCN=localhost\n[req]\ndistinguished_name = dn\n[EXT]\nsubjectAltName=DNS:localhost\nkeyUsage=digitalSignature\nextendedKeyUsage=serverAuth")
 ```
2. Modify the NestFactory in the src/main.ts file
 ```ts
import { NestFactory } from '@nestjs/core';  
import { AppModule } from './app.module';  
  
async function bootstrap() {  
  const fs = require('fs');  
 const keyFile = fs.readFileSync(__dirname + '/../localhost.key');  
 const certFile = fs.readFileSync(__dirname + '/../localhost.crt');  
 const app = await NestFactory.create(AppModule, {  
    httpsOptions: {  
      key: keyFile,  
	  cert: certFile,  
    },  
  });  
 await app.listen(3000);  
}  
bootstrap();
 ```
