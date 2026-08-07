import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import fs from 'fs';
import dotenv from 'dotenv';

dotenv.config({ path: '.dev.vars' });

async function run() {
  const serviceAccountEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL!;
  const privateKey = process.env.GOOGLE_PRIVATE_KEY!.replace(/\\n/g, '\n');
  let spreadsheetId = process.env.SPREADSHEET_ID!;
  const match = spreadsheetId.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) spreadsheetId = match[1];

  const auth = new JWT({
    email: serviceAccountEmail,
    key: privateKey,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const doc = new GoogleSpreadsheet(spreadsheetId, auth);
  await doc.loadInfo();
  
  console.log("=== Sheets ===");
  for (const sheetId in doc.sheetsById) {
    const sheet = doc.sheetsById[sheetId];
    console.log(`Sheet: ${sheet.title}`);
    try {
      await sheet.loadHeaderRow();
      console.log(`Headers:`, sheet.headerValues);
      const rows = await sheet.getRows();
      if (rows.length > 0) {
         console.log(`Row 0 _rawData:`, rows[0]._rawData);
         // console.log(`Row 0 aValue:`, rows[0].get(sheet.headerValues[1]));
      }
    } catch (e: any) {
      console.log(`Error reading sheet ${sheet.title}:`, e.message);
    }
    console.log("-------------");
  }
}

run().catch(console.error);
