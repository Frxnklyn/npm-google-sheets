import { readFile } from "node:fs/promises";
import {
  GoogleSheetConnection,
  GoogleSheetsDataType,
} from "../dist/index.js";

const credentialsPath = process.env.GOOGLE_SERVICE_ACCOUNT_FILE;
const spreadsheetUrl = process.env.GOOGLE_SHEETS_URL;

if (!credentialsPath) {
  throw new Error("GOOGLE_SERVICE_ACCOUNT_FILE muss auf die Service-Account-JSON-Datei zeigen.");
}
if (!spreadsheetUrl) {
  throw new Error("GOOGLE_SHEETS_URL muss den Link zum Spreadsheet enthalten.");
}

const credentials = JSON.parse(await readFile(credentialsPath, "utf8"));
const connection = new GoogleSheetConnection(spreadsheetUrl, credentials);
const workbook = new GoogleSheetsDataType(connection);

await workbook.dataRead();

console.log("Spreadsheet:", workbook.getName());
console.log(
  "Sheets:",
  workbook.getSheets().map((sheet) => ({
    name: sheet.getName(),
    index: sheet.getIndex(),
  })),
);
