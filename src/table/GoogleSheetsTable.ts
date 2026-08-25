import type {
  AttributeInterface,
  AttributeType,
  ColumnDataTypeInterface,
  FilterComparator,
  FilterInterface,
  RelationInterface,
  RowDataTypeInterface,
  TableCellValue,
  TableDataTypeInterface,
} from "@frxnklyn/datatypes";
import type { GoogleSheet } from "../GoogleSheet.js";
import type { GoogleSheetConnection } from "../GoogleSheetConnection.js";
import { columnIndexToLetters, quoteSheetName } from "../internal/a1.js";
import { inferType } from "../internal/value.js";

export class GoogleSheetsTable implements TableDataTypeInterface {
  private headers: string[] = [];
  private rows: GoogleSheetsRow[] = [];
  private relations: RelationInterface[] = [];
  private filters: FilterInterface[] = [];
  private comparator: FilterComparator = "and";
  private readonly connection: GoogleSheetConnection;

  public constructor(private readonly sheet: GoogleSheet) {
    this.connection = sheet.getConnection();
  }

  public getName(): string | undefined { return this.sheet.getName(); }
  public getConnection(): GoogleSheetConnection { return this.connection; }
  public getHeaders(): readonly string[] { return this.headers; }
  public getAttributes(): readonly AttributeInterface[] {
    return this.headers.map((header, index) => new GoogleSheetsAttribute(this, index, header));
  }
  public getRows(): readonly GoogleSheetsRow[] { return this.rows; }
  public getRow(index: number): GoogleSheetsRow | undefined { return this.rows[index]; }

  public addRow(row: readonly TableCellValue[] | RowDataTypeInterface): this {
    const values = "getValues" in row ? row.getValues() : row;
    const width = Math.max(this.headers.length, values.length);
    this.ensureHeaders(width);
    this.rows.push(new GoogleSheetsRow(this, this.rows.length, normalized(values, width)));
    return this;
  }

  public getColumns(): readonly GoogleSheetsColumn[] {
    return this.headers.map((_, index) => new GoogleSheetsColumn(this, index));
  }
  public getColumn(index: number): GoogleSheetsColumn | undefined;
  public getColumn(header: string): GoogleSheetsColumn | undefined;
  public getColumn(reference: number | string): GoogleSheetsColumn | undefined {
    const index = typeof reference === "number" ? reference : this.headers.indexOf(reference);
    return index >= 0 && index < this.headers.length ? new GoogleSheetsColumn(this, index) : undefined;
  }

  public getRelations(): readonly RelationInterface[] { return this.relations; }
  public addRelation(relation: RelationInterface): this { this.relations.push(relation); return this; }
  public setRelations(relations: readonly RelationInterface[]): this { this.relations = [...relations]; return this; }
  public getFilters(): readonly FilterInterface[] { return this.filters; }
  public addFilter(filter: FilterInterface): this { this.filters.push(filter); return this; }
  public setFilters(filters: readonly FilterInterface[]): this { this.filters = [...filters]; return this; }
  public clearFilters(): this { this.filters = []; return this; }
  public getComparator(): FilterComparator { return this.comparator; }
  public setComparator(comparator: FilterComparator): this { this.comparator = comparator; return this; }

  public async dataRead(): Promise<this> {
    const name = this.sheet.requireName();
    const values = await this.connection.readValues({
      range: quoteSheetName(name),
      majorDimension: "ROWS",
      valueRenderOption: "UNFORMATTED_VALUE",
      dateTimeRenderOption: "FORMATTED_STRING",
    });
    this.rebuildFromValues(values);
    if (this.filters.length > 0) {
      this.rows = this.rows.filter((row) => {
        const results = this.filters.map((filter) => matches(row.getValue(filter.getAttribute()), filter));
        return this.comparator === "and" ? results.every(Boolean) : results.some(Boolean);
      }).map((row, index) => new GoogleSheetsRow(this, index, row.getValues()));
    }
    return this;
  }

  public async dataSave(): Promise<this> {
    this.sheet.replaceTable(this.headers, this.rows.map((row) => row.getValues()));
    await this.sheet.dataSave();
    return this;
  }

  private rebuildFromValues(values: readonly (readonly unknown[])[]): void {
    const width = values.length === 0 ? 0 : Math.max(...values.map((row) => row.length));
    const headerRow = values[0] ?? [];
    this.headers = Array.from({ length: width }, (_, column) => {
      const value = headerRow[column];
      return value === null || value === undefined || value === "" ? columnIndexToLetters(column) : String(value);
    });
    this.rows = Array.from({ length: Math.max(0, values.length - 1) }, (_, row) =>
      new GoogleSheetsRow(
        this,
        row,
        Array.from({ length: width }, (_, column) => toTableCellValue(values[row + 1]?.[column])),
      ),
    );
  }

  private ensureHeaders(width: number): void {
    while (this.headers.length < width) this.headers.push(columnIndexToLetters(this.headers.length));
  }
}

export class GoogleSheetsRow implements RowDataTypeInterface {
  private readonly values: readonly TableCellValue[];
  public constructor(
    private readonly table: GoogleSheetsTable,
    private readonly index: number,
    values: readonly TableCellValue[],
  ) { this.values = [...values]; }
  public getTable(): GoogleSheetsTable { return this.table; }
  public getIndex(): number { return this.index; }
  public getHeaders(): readonly string[] { return this.table.getHeaders(); }
  public getValues(): readonly TableCellValue[] { return this.values; }
  public getValue(columnIndex: number): TableCellValue | undefined;
  public getValue(header: string): TableCellValue | undefined;
  public getValue(reference: number | string): TableCellValue | undefined {
    const index = typeof reference === "number" ? reference : this.getHeaders().indexOf(reference);
    return index < 0 ? undefined : this.values[index];
  }
}

export class GoogleSheetsColumn implements ColumnDataTypeInterface {
  public constructor(private readonly table: GoogleSheetsTable, private readonly index: number) {}
  public getTable(): GoogleSheetsTable { return this.table; }
  public getIndex(): number { return this.index; }
  public getHeader(): string { return this.table.getHeaders()[this.index] ?? columnIndexToLetters(this.index); }
  public getAttribute(): GoogleSheetsAttribute { return new GoogleSheetsAttribute(this.table, this.index, this.getHeader()); }
  public getValues(): readonly TableCellValue[] { return this.table.getRows().map((row) => row.getValue(this.index) ?? null); }
  public getValue(rowIndex: number): TableCellValue | undefined { return this.table.getRow(rowIndex)?.getValue(this.index); }
}

export class GoogleSheetsAttribute implements AttributeInterface {
  public constructor(
    private readonly table: GoogleSheetsTable,
    private readonly index: number,
    private readonly name: string,
  ) {}
  public getTable(): GoogleSheetsTable { return this.table; }
  public getIndex(): number { return this.index; }
  public getName(): string { return this.name; }
  public getType(): AttributeType { return inferType(this.table.getColumn(this.index)?.getValues() ?? []); }
  public isNullable(): boolean { return (this.table.getColumn(this.index)?.getValues() ?? []).some((value) => value === null); }
}

function normalized(values: readonly TableCellValue[], width: number): TableCellValue[] {
  return Array.from({ length: width }, (_, index) => values[index] ?? null);
}

function toTableCellValue(value: unknown): TableCellValue {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value;
  if (["string", "number", "boolean", "bigint"].includes(typeof value)) {
    return value as TableCellValue;
  }
  return String(value);
}

function matches(actual: TableCellValue | undefined, filter: FilterInterface): boolean {
  const expected = filter.getValue();
  switch (filter.getOperator()) {
    case "equals": return equal(actual, expected);
    case "notEquals": return !equal(actual, expected);
    case "contains": return String(actual ?? "").includes(String(expected ?? ""));
    case "in": return Array.isArray(expected) && expected.some((value) => equal(actual, value));
    case "greaterThan": return compare(actual, expected) > 0;
    case "greaterThanOrEquals": return compare(actual, expected) >= 0;
    case "lessThan": return compare(actual, expected) < 0;
    case "lessThanOrEquals": return compare(actual, expected) <= 0;
  }
}

function equal(left: unknown, right: unknown): boolean {
  if (left instanceof Date && right instanceof Date) return left.getTime() === right.getTime();
  return left === right;
}

function compare(left: unknown, right: unknown): number {
  const a = left instanceof Date ? left.getTime() : left;
  const b = right instanceof Date ? right.getTime() : right;
  if (typeof a === "string" && typeof b === "string") return a.localeCompare(b);
  if ((typeof a === "number" || typeof a === "bigint") && (typeof b === "number" || typeof b === "bigint")) {
    const leftNumber = Number(a);
    const rightNumber = Number(b);
    return leftNumber === rightNumber ? 0 : leftNumber > rightNumber ? 1 : -1;
  }
  return Number.NaN;
}
