import assert from "node:assert/strict";
import test from "node:test";
import {
  extractSpreadsheetId,
  GoogleSheetConnection,
  GoogleSheetsDataType,
  GoogleSheetsFilter,
} from "../dist/index.js";

const SHEET_ID = "1ABC_def-123456789";
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit?gid=0`;
const TEST_CREDENTIALS = {
  type: "service_account",
  client_email: "test@example.iam.gserviceaccount.com",
  private_key: "not-used-by-the-mocked-client",
};

function fixtureSheet() {
  return {
    properties: { title: "Produkte", index: 0, sheetId: 10 },
    data: [{
      startRow: 0,
      startColumn: 0,
      rowData: [
        { values: [
          { effectiveValue: { stringValue: "Name" } },
          { effectiveValue: { stringValue: "Preis" } },
          { effectiveValue: { stringValue: "Brutto" } },
        ] },
        { values: [
          { effectiveValue: { stringValue: "Tastatur" } },
          { effectiveValue: { numberValue: 49.9 } },
          { effectiveValue: { numberValue: 59.381 }, userEnteredValue: { formulaValue: "=B2*1.19" } },
        ] },
        { values: [
          { effectiveValue: { stringValue: "Maus" } },
          { effectiveValue: { numberValue: 19.9 } },
          { effectiveValue: { numberValue: 23.681 }, userEnteredValue: { formulaValue: "=B3*1.19" } },
        ] },
      ],
    }],
  };
}

function createConnection() {
  const reads = [];
  const valueReads = [];
  const updates = [];
  const clears = [];
  const batchUpdates = [];
  const client = {
    spreadsheets: {
      get: async (request) => {
        reads.push(request);
        return {
          data: {
            properties: { title: "Shop" },
            sheets: [fixtureSheet()],
          },
        };
      },
      batchUpdate: async (request) => {
        batchUpdates.push(request);
        return { data: {} };
      },
      values: {
        get: async (request) => {
          valueReads.push(request);
          return {
            data: {
              values: [
                ["Name", "Preis", "Brutto"],
                ["Tastatur", 49.9, 59.381],
                ["Maus", 19.9, 23.681],
              ],
            },
          };
        },
        update: async (request) => {
          updates.push(request);
          return { data: {} };
        },
        clear: async (request) => {
          clears.push(request);
          return { data: {} };
        },
      },
    },
  };

  const connection = new GoogleSheetConnection(SHEET_URL, TEST_CREDENTIALS);
  connection.api = client;
  return { connection, reads, valueReads, updates, clears, batchUpdates };
}

test("validates a Google Sheets URL and extracts its spreadsheet id", () => {
  const connection = new GoogleSheetConnection(SHEET_URL, TEST_CREDENTIALS);
  assert.equal(connection.getSpreadsheetId(), SHEET_ID);
  assert.throws(() => new GoogleSheetConnection("", TEST_CREDENTIALS), /sheetUrl is required/);
  assert.throws(
    () => new GoogleSheetConnection("https://example.com/spreadsheets/d/abc/edit", TEST_CREDENTIALS),
    /Invalid Google Sheets URL/,
  );
  assert.throws(
    () => new GoogleSheetConnection("https://docs.google.com/spreadsheets/d/e/published/pubhtml", TEST_CREDENTIALS),
    /Invalid Google Sheets URL/,
  );
});

test("keeps getSheet lazy and delegates the first read to the connection", async () => {
  const { connection, reads } = createConnection();
  const workbook = new GoogleSheetsDataType(connection);

  const sheet = workbook.getSheet("Produkte");
  assert.equal(reads.length, 0);
  assert.strictEqual(workbook.getSheet("Produkte"), sheet);

  await sheet.dataRead();
  assert.equal(reads.length, 1);
  assert.deepEqual(reads[0].ranges, ["'Produkte'"]);
  assert.equal(reads[0].includeGridData, true);
});

test("lets workbook and table choose their own read options", async () => {
  const { connection, reads, valueReads } = createConnection();
  const workbook = new GoogleSheetsDataType(connection);

  await workbook.dataRead();
  assert.equal(reads[0].spreadsheetId, SHEET_ID);
  assert.equal(reads[0].includeGridData, true);
  assert.equal(workbook.getName(), "Shop");
  assert.equal(workbook.getSheets().length, 1);

  const sheet = workbook.getSheet("Produkte");
  assert.equal(sheet.getCell("B2").getValue(), 49.9);
  assert.equal(sheet.getCell("C2").getFormula(), "=B2*1.19");
  assert.equal(sheet.getCell(1, 2).getAddress(), "C2");

  const table = sheet.asTable();
  await table.dataRead();
  assert.equal(valueReads.length, 1);
  assert.equal(valueReads[0].range, "'Produkte'");
  assert.equal(valueReads[0].majorDimension, "ROWS");
  assert.equal(valueReads[0].valueRenderOption, "UNFORMATTED_VALUE");
  assert.deepEqual(table.getHeaders(), ["Name", "Preis", "Brutto"]);
  assert.equal(table.getRow(0).getValue("Name"), "Tastatur");
  assert.equal(table.getColumn("Preis").getAttribute().getType(), "number");
});

test("filters locally and delegates value writes to the connection", async () => {
  const { connection, updates, clears } = createConnection();
  const workbook = new GoogleSheetsDataType(connection);
  const table = workbook.getTable("Produkte");

  table.addFilter(new GoogleSheetsFilter("Preis", "greaterThan", 20));
  await table.dataRead();
  assert.equal(table.getRows().length, 1);

  table.addRow(["Monitor", 199, null]);
  await table.dataSave();
  assert.equal(clears.length, 1);
  assert.equal(updates.length, 1);
  assert.equal(updates[0].spreadsheetId, SHEET_ID);
  assert.equal(updates[0].range, "'Produkte'");
  assert.deepEqual(updates[0].requestBody.values[2], ["Monitor", 199, null]);
});

test("validates lazy addresses and formulas without a request", () => {
  const { connection, reads } = createConnection();
  const workbook = new GoogleSheetsDataType(connection);
  const sheet = workbook.getSheet("Produkte");

  assert.equal(sheet.getCell("AA20").getAddress(), "AA20");
  assert.equal(sheet.getCell("AA20").isResolved(), false);
  assert.throws(() => sheet.getCell("20AA"), /Invalid A1/);
  assert.throws(() => sheet.setFormula("A1", "SUM(B1:B2)"), /start with '='/);
  assert.equal(reads.length, 0);
});

test("keeps the standalone link utility for common shared links", () => {
  assert.equal(extractSpreadsheetId(SHEET_URL), SHEET_ID);
  assert.equal(extractSpreadsheetId(`https://drive.google.com/open?id=${SHEET_ID}`), SHEET_ID);
});
