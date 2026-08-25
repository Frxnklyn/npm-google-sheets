import type {
  AttributeType,
  CellDataTypeInterface,
  TableCellValue,
} from "@frxnklyn/datatypes";
import type { GoogleSheet } from "./GoogleSheet.js";
import { toA1Address } from "./internal/a1.js";

export class GoogleSheetsCell implements CellDataTypeInterface {
  private resolved = false;
  private merged = false;
  private value: TableCellValue = null;
  private formula: string | undefined;
  private type: AttributeType | undefined;

  public constructor(
    private readonly sheet: GoogleSheet,
    private readonly rowIndex: number,
    private readonly columnIndex: number,
  ) {}

  public getSheet(): GoogleSheet { return this.sheet; }
  public getRowIndex(): number { return this.rowIndex; }
  public getColumnIndex(): number { return this.columnIndex; }
  public getAddress(): string { return toA1Address(this.rowIndex, this.columnIndex); }
  public isResolved(): boolean { return this.resolved; }
  public isMerged(): boolean { return this.merged; }
  public getValue(): TableCellValue | undefined { return this.resolved ? this.value : undefined; }
  public hasFormula(): boolean { return this.formula !== undefined; }
  public getFormula(): string | undefined { return this.formula; }
  public getType(): AttributeType | undefined { return this.type; }

  public hydrate(value: TableCellValue, formula: string | undefined, merged: boolean): void {
    this.resolved = true;
    this.value = value;
    this.formula = formula;
    this.merged = merged;
    this.type = typeOf(value);
  }

  public setValue(value: TableCellValue): void {
    this.hydrate(value, undefined, this.merged);
  }

  public setFormula(formula: string): void {
    if (!formula.startsWith("=")) {
      throw new TypeError("A Google Sheets formula must start with '='.");
    }
    this.resolved = true;
    this.formula = formula;
  }
}

function typeOf(value: TableCellValue): AttributeType {
  if (value === null) return "unknown";
  if (value instanceof Date) return "datetime";
  if (typeof value === "boolean") return "boolean";
  if (typeof value === "bigint") return "integer";
  if (typeof value === "number") return Number.isInteger(value) ? "integer" : "number";
  return "string";
}
