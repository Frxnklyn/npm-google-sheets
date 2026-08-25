import type {
  ExcelSheetDataTypeInterface,
  TableCellValue,
} from "@frxnklyn/datatypes";
import type { sheets_v4 } from "googleapis";
import { GoogleSheetsCell } from "./GoogleSheetsCell.js";
import { GoogleSheetsTable } from "./table/GoogleSheetsTable.js";
import { assertIndex, fromA1Address, quoteSheetName } from "./internal/a1.js";
import { readCellValue, readFormula, toApiValue } from "./internal/value.js";
import type { GoogleSheetsDataType } from "./GoogleSheetsDataType.js";
import type { GoogleSheetConnection } from "./GoogleSheetConnection.js";

export class GoogleSheet implements ExcelSheetDataTypeInterface {
  private name: string | undefined;
  private index: number | undefined;
  private sheetId: number | undefined;
  private readonly cells = new Map<string, GoogleSheetsCell>();
  private readonly table: GoogleSheetsTable;
  private readonly connection: GoogleSheetConnection;

  public constructor(
    private readonly workbook: GoogleSheetsDataType,
    reference: string | number,
  ) {
    this.connection = workbook.getConnection();
    if (typeof reference === "string") this.name = reference;
    else {
      assertIndex(reference, "sheet index");
      this.index = reference;
    }
    this.table = new GoogleSheetsTable(this);
  }

  public getExcel(): GoogleSheetsDataType { return this.workbook; }
  public getConnection(): GoogleSheetConnection { return this.connection; }
  public getName(): string | undefined { return this.name; }
  public getIndex(): number | undefined { return this.index; }
  public getSheetId(): number | undefined { return this.sheetId; }

  public getCells(): readonly GoogleSheetsCell[] {
    return [...this.cells.values()].sort((a, b) =>
      a.getRowIndex() - b.getRowIndex() || a.getColumnIndex() - b.getColumnIndex(),
    );
  }

  public getCell(rowIndex: number, columnIndex: number): GoogleSheetsCell;
  public getCell(address: string): GoogleSheetsCell;
  public getCell(first: number | string, second?: number): GoogleSheetsCell {
    const [rowIndex, columnIndex] = typeof first === "string"
      ? fromA1Address(first)
      : [first, requiredColumn(second)];
    assertIndex(rowIndex, "rowIndex");
    assertIndex(columnIndex, "columnIndex");
    const key = `${rowIndex}:${columnIndex}`;
    let cell = this.cells.get(key);
    if (!cell) {
      cell = new GoogleSheetsCell(this, rowIndex, columnIndex);
      this.cells.set(key, cell);
    }
    return cell;
  }

  public setCell(rowIndex: number, columnIndex: number, value: TableCellValue): this;
  public setCell(address: string, value: TableCellValue): this;
  public setCell(first: number | string, second: number | TableCellValue, third?: TableCellValue): this {
    if (typeof first === "string") this.getCell(first).setValue(second as TableCellValue);
    else this.getCell(first, second as number).setValue(third ?? null);
    return this;
  }

  public setFormula(rowIndex: number, columnIndex: number, formula: string): this;
  public setFormula(address: string, formula: string): this;
  public setFormula(first: number | string, second: number | string, third?: string): this {
    if (typeof first === "string") this.getCell(first).setFormula(second as string);
    else this.getCell(first, second as number).setFormula(requiredFormula(third));
    return this;
  }

  public asTable(): GoogleSheetsTable { return this.table; }

  public replaceTable(headers: readonly string[], rows: readonly (readonly TableCellValue[])[]): void {
    const previous = this.getCells();
    const width = Math.max(headers.length, ...rows.map((row) => row.length), 0);
    const height = Math.max(rows.length + 1, previous.length === 0 ? 0 : Math.max(...previous.map((cell) => cell.getRowIndex())) + 1);
    const oldWidth = previous.length === 0 ? 0 : Math.max(...previous.map((cell) => cell.getColumnIndex())) + 1;
    for (let row = 0; row < height; row += 1) {
      for (let column = 0; column < Math.max(width, oldWidth); column += 1) {
        const value = row === 0 ? (headers[column] ?? null) : (rows[row - 1]?.[column] ?? null);
        this.getCell(row, column).setValue(value);
      }
    }
  }

  public async dataRead(): Promise<this> {
    const name = await this.workbook.resolveSheetName(this);
    const spreadsheet = await this.connection.read({
      ranges: [quoteSheetName(name)],
      includeGridData: true,
    });
    const sheet = spreadsheet.sheets?.find((entry) => entry.properties?.title === name);
    if (!sheet) throw new Error(`Sheet '${name}' does not exist or is not readable.`);
    this.workbook.registerSheet(sheet, this);
    this.hydrate(sheet);
    return this;
  }

  public async dataSave(): Promise<this> {
    await this.workbook.ensureSheetExists(this);
    const name = this.requireName();
    const known = this.getCells();
    if (known.length === 0) {
      await this.connection.clearValues({ range: quoteSheetName(name) });
      return this;
    }
    const maxRow = Math.max(...known.map((cell) => cell.getRowIndex()));
    const maxColumn = Math.max(...known.map((cell) => cell.getColumnIndex()));
    const values = Array.from({ length: maxRow + 1 }, (_, row) =>
      Array.from({ length: maxColumn + 1 }, (_, column) => {
        const cell = this.getCell(row, column);
        return cell.hasFormula() ? cell.getFormula()! : toApiValue(cell.getValue() ?? null);
      }),
    );
    const range = quoteSheetName(name);
    await this.connection.clearValues({ range });
    await this.connection.writeValues({
      range,
      values,
      valueInputOption: "USER_ENTERED",
    });
    return this;
  }

  public hydrate(sheet: sheets_v4.Schema$Sheet): void {
    const properties = sheet.properties;
    this.name = properties?.title ?? this.name;
    this.index = properties?.index ?? this.index;
    this.sheetId = properties?.sheetId ?? this.sheetId;
    const merges = sheet.merges ?? [];
    for (const cell of this.cells.values()) cell.hydrate(null, undefined, false);
    for (const grid of sheet.data ?? []) {
      const startRow = grid.startRow ?? 0;
      const startColumn = grid.startColumn ?? 0;
      grid.rowData?.forEach((row, rowOffset) => {
        row.values?.forEach((cellData, columnOffset) => {
          const rowIndex = startRow + rowOffset;
          const columnIndex = startColumn + columnOffset;
          this.getCell(rowIndex, columnIndex).hydrate(
            readCellValue(cellData),
            readFormula(cellData),
            merges.some((merge) => inRange(rowIndex, columnIndex, merge)),
          );
        });
      });
    }
  }

  public updateIdentity(name: string, index: number, sheetId: number): void {
    this.name = name;
    this.index = index;
    this.sheetId = sheetId;
  }

  public requireName(): string {
    if (this.name === undefined) throw new Error("The sheet name is unresolved. Call dataRead() first.");
    return this.name;
  }
}

function requiredColumn(value: number | undefined): number {
  if (value === undefined) throw new TypeError("columnIndex is required.");
  return value;
}

function requiredFormula(value: string | undefined): string {
  if (value === undefined) throw new TypeError("formula is required.");
  return value;
}

function inRange(row: number, column: number, range: sheets_v4.Schema$GridRange): boolean {
  return row >= (range.startRowIndex ?? 0)
    && row < (range.endRowIndex ?? Number.POSITIVE_INFINITY)
    && column >= (range.startColumnIndex ?? 0)
    && column < (range.endColumnIndex ?? Number.POSITIVE_INFINITY);
}
