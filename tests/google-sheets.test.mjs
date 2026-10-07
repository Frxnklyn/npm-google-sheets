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

function createConnection({ clearCache = true, cacheTtlMs } = {}) {
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

  const connection = new GoogleSheetConnection(
    SHEET_URL,
    TEST_CREDENTIALS,
    cacheTtlMs === undefined ? undefined : { cacheTtlMs },
  );
  if (clearCache) connection.clearCache();
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

test("supports finite and manual-only cache policies", async () => {
  const defaults = new GoogleSheetConnection(SHEET_URL, TEST_CREDENTIALS);
  assert.equal(defaults.getCacheTtlMs(), 60_000);

  const manual = createConnection({ cacheTtlMs: null });
  assert.equal(manual.connection.getCacheTtlMs(), null);

  const firstTable = new GoogleSheetsDataType(manual.connection).getTable("Produkte");
  await firstTable.dataRead();
  await new Promise((resolve) => setTimeout(resolve, 15));

  const second = createConnection({ clearCache: false, cacheTtlMs: null });
  const secondTable = new GoogleSheetsDataType(second.connection).getTable("Produkte");
  await secondTable.dataRead();

  assert.equal(
    manual.valueReads.length + second.valueReads.length,
    1,
    "manual cache mode does not expire automatically",
  );

  assert.throws(
    () => new GoogleSheetConnection(SHEET_URL, TEST_CREDENTIALS, { cacheTtlMs: -1 }),
    /cacheTtlMs/,
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
  const { connection, valueReads, updates, clears } = createConnection();
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

  await table.dataRead();
  assert.equal(valueReads.length, 2, "a write invalidates the cached table read");
});

test("shares reads across table, workbook and connection instances", async () => {
  const first = createConnection();
  const second = createConnection({ clearCache: false });

  const firstTable = new GoogleSheetsDataType(first.connection).getTable("Produkte", {
    where: { attribute: "Preis", operator: "greaterThan", value: 20 },
  });
  const secondTable = new GoogleSheetsDataType(second.connection).getTable("Produkte", {
    where: { attribute: "Name", operator: "contains", value: "Maus" },
  });

  await Promise.all([firstTable.dataRead(), secondTable.dataRead()]);

  assert.equal(first.valueReads.length + second.valueReads.length, 1);
  assert.equal(firstTable.getRows().length, 1);
  assert.equal(secondTable.getRows().length, 1);

  second.connection.clearCache();
  await secondTable.dataRead();
  assert.equal(second.valueReads.length, 1, "manual clearCache forces the next Google read");
});

test("refreshes a cached read on the first request after the TTL", async () => {
  const first = createConnection({ cacheTtlMs: 5 });
  const firstTable = new GoogleSheetsDataType(first.connection).getTable("Produkte");
  await firstTable.dataRead();
  assert.equal(first.valueReads.length, 1);

  await new Promise((resolve) => setTimeout(resolve, 15));

  const second = createConnection({ clearCache: false, cacheTtlMs: 5 });
  const secondTable = new GoogleSheetsDataType(second.connection).getTable("Produkte");
  await secondTable.dataRead();

  assert.equal(second.valueReads.length, 1);
});

test("executes portable SQL-like queries lazily", async () => {
  const { connection, valueReads } = createConnection();
  const workbook = new GoogleSheetsDataType(connection);

  const plain = workbook.getTable("Produkte");
  const queried = workbook.getTable("Produkte", {
    where: {
      comparator: "and",
      conditions: [
        { attribute: "Preis", operator: "greaterThan", value: 10 },
        {
          comparator: "or",
          conditions: [
            { attribute: "Name", operator: "contains", value: "ast" },
            { attribute: "Name", operator: "contains", value: "Mau" },
          ],
        },
      ],
    },
    orderBy: [{ attribute: "Preis", direction: "desc" }],
    offset: 1,
    limit: 1,
    select: ["Name", "Preis"],
  });

  assert.equal(valueReads.length, 0);
  assert.notStrictEqual(queried, plain);
  assert.strictEqual(workbook.getTable("Produkte"), plain);

  await queried.dataRead();

  assert.equal(valueReads.length, 1);
  assert.deepEqual(queried.getHeaders(), ["Name", "Preis"]);
  assert.equal(queried.getRows().length, 1);
  assert.equal(queried.getRow(0).getValue("Name"), "Maus");
  assert.equal(queried.getRow(0).getValue("Preis"), 19.9);
  await assert.rejects(() => queried.dataSave(), /read-only/);
});

test("supports grouping and aggregations through the portable query contract", async () => {
  const { connection } = createConnection();
  const workbook = new GoogleSheetsDataType(connection);

  const aggregate = workbook.getTable("Produkte", {
    aggregations: [
      { function: "count", as: "Anzahl" },
      { function: "sum", attribute: "Preis", as: "Summe" },
      { function: "avg", attribute: "Preis", as: "Durchschnitt" },
      { function: "min", attribute: "Preis", as: "Minimum" },
      { function: "max", attribute: "Preis", as: "Maximum" },
    ],
  });

  await aggregate.dataRead();

  assert.deepEqual(
    aggregate.getHeaders(),
    ["Anzahl", "Summe", "Durchschnitt", "Minimum", "Maximum"],
  );
  assert.deepEqual(aggregate.getRow(0).getValues(), [2, 69.8, 34.9, 19.9, 49.9]);

  const grouped = workbook.getTable("Produkte", {
    groupBy: ["Name"],
    aggregations: [{ function: "count", as: "Anzahl" }],
    orderBy: [{ attribute: "Name", direction: "asc" }],
  });

  await grouped.dataRead();

  assert.deepEqual(grouped.getHeaders(), ["Name", "Anzahl"]);
  assert.deepEqual(grouped.getRows().map((row) => row.getValues()), [
    ["Maus", 1],
    ["Tastatur", 1],
  ]);
});

test("implements TableSource existence checks with metadata-only reads", async () => {
  const { connection, reads } = createConnection();
  const workbook = new GoogleSheetsDataType(connection);

  assert.equal(await workbook.hasTable("Produkte"), true);
  assert.equal(await workbook.hasTable("Fehlt"), false);
  assert.equal(reads.length, 1, "identical metadata reads share the cache");
  assert.equal(reads.every((request) => request.includeGridData === false), true);
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
