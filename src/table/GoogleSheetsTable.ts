import type {
  AttributeInterface,
  AttributeType,
  ColumnDataTypeInterface,
  FilterComparator,
  FilterInterface,
  FilterOperator,
  RelationInterface,
  RowDataTypeInterface,
  TableCellValue,
  TableDataTypeInterface,
  TableQueryAggregation,
  TableQueryCondition,
  TableQueryInterface,
  TableQueryOrder,
} from "@frxnklyn/datatypes";
import type { GoogleSheet } from "../GoogleSheet.js";
import type { GoogleSheetConnection } from "../GoogleSheetConnection.js";
import { columnIndexToLetters, quoteSheetName } from "../internal/a1.js";
import { inferType } from "../internal/value.js";

type QueryState = {
  headers: string[];
  rows: TableCellValue[][];
};

export class GoogleSheetsTable implements TableDataTypeInterface {
  private headers: string[] = [];
  private rows: GoogleSheetsRow[] = [];
  private relations: RelationInterface[] = [];
  private filters: FilterInterface[] = [];
  private comparator: FilterComparator = "and";
  private readonly connection: GoogleSheetConnection;

  public constructor(
    private readonly sheet: GoogleSheet,
    private readonly query?: TableQueryInterface,
  ) {
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
    let state = this.currentState();

    if (this.filters.length > 0) {
      state.rows = state.rows.filter((row) => {
        const results = this.filters.map((filter) =>
          matchesValue(
            valueByHeader(state.headers, row, filter.getAttribute()),
            filter.getOperator(),
            filter.getValue(),
          ));
        return this.comparator === "and" ? results.every(Boolean) : results.some(Boolean);
      });
    }

    if (this.query?.where) {
      state.rows = state.rows.filter((row) => matchesCondition(state.headers, row, this.query!.where!));
    }

    if ((this.query?.groupBy?.length ?? 0) > 0 || (this.query?.aggregations?.length ?? 0) > 0) {
      state = groupAndAggregate(state, this.query?.groupBy ?? [], this.query?.aggregations ?? []);
    }

    if ((this.query?.orderBy?.length ?? 0) > 0) {
      state.rows = orderRows(state.headers, state.rows, this.query!.orderBy!);
    }

    if (this.query?.offset !== undefined || this.query?.limit !== undefined) {
      state.rows = paginate(state.rows, this.query?.offset, this.query?.limit);
    }

    if (this.query?.select !== undefined) {
      state = project(state, this.query.select);
    }

    this.replaceState(state);
    return this;
  }

  public async dataSave(): Promise<this> {
    if (this.query !== undefined) {
      throw new Error("Queried table views are read-only. Use getTable(name) without a query for writes.");
    }

    this.sheet.replaceTable(this.headers, this.rows.map((row) => row.getValues()));
    await this.sheet.dataSave();
    return this;
  }

  private currentState(): QueryState {
    return {
      headers: [...this.headers],
      rows: this.rows.map((row) => [...row.getValues()]),
    };
  }

  private replaceState(state: QueryState): void {
    this.headers = [...state.headers];
    this.rows = state.rows.map((values, index) => new GoogleSheetsRow(this, index, values));
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

function matchesCondition(
  headers: readonly string[],
  row: readonly TableCellValue[],
  condition: TableQueryCondition,
): boolean {
  if ("conditions" in condition) {
    const results = condition.conditions.map((entry) => matchesCondition(headers, row, entry));
    return condition.comparator === "and" ? results.every(Boolean) : results.some(Boolean);
  }

  return matchesValue(
    row[requiredColumnIndex(headers, condition.attribute)],
    condition.operator,
    condition.value,
  );
}

function matchesValue(
  actual: TableCellValue | undefined,
  operator: FilterOperator,
  expected: TableCellValue | readonly TableCellValue[] | undefined,
): boolean {
  switch (operator) {
    case "equals": return equal(actual, expected);
    case "notEquals": return !equal(actual, expected);
    case "contains": return String(actual ?? "").includes(String(expected ?? ""));
    case "in": return Array.isArray(expected) && expected.some((value) => equal(actual, value));
    case "greaterThan": return compareValues(actual, scalar(expected)) > 0;
    case "greaterThanOrEquals": return compareValues(actual, scalar(expected)) >= 0;
    case "lessThan": return compareValues(actual, scalar(expected)) < 0;
    case "lessThanOrEquals": return compareValues(actual, scalar(expected)) <= 0;
  }
}

function groupAndAggregate(
  state: QueryState,
  groupBy: readonly string[],
  aggregations: readonly TableQueryAggregation[],
): QueryState {
  const groupIndexes = groupBy.map((attribute) => requiredColumnIndex(state.headers, attribute));
  const groups = new Map<string, TableCellValue[][]>();

  if (groupIndexes.length === 0) {
    groups.set("__all__", state.rows);
  } else {
    for (const row of state.rows) {
      const key = groupIndexes.map((index) => keyPart(row[index])).join("\u001f");
      const rows = groups.get(key);
      if (rows) rows.push(row);
      else groups.set(key, [row]);
    }
  }

  const headers = [...groupBy, ...aggregations.map((aggregation) => aggregation.as)];
  const rows = [...groups.values()].map((groupRows) => {
    const first = groupRows[0] ?? [];
    const groupValues = groupIndexes.map((index) => first[index] ?? null);
    const aggregateValues = aggregations.map((aggregation) =>
      aggregate(state.headers, groupRows, aggregation));
    return [...groupValues, ...aggregateValues];
  });

  return { headers, rows };
}

function aggregate(
  headers: readonly string[],
  rows: readonly (readonly TableCellValue[])[],
  aggregation: TableQueryAggregation,
): TableCellValue {
  if (aggregation.function === "count") {
    if (aggregation.attribute === undefined) return rows.length;
    const index = requiredColumnIndex(headers, aggregation.attribute);
    return rows.filter((row) => row[index] !== null && row[index] !== undefined).length;
  }

  const index = requiredColumnIndex(headers, aggregation.attribute);
  const values = rows
    .map((row) => row[index])
    .filter((value): value is Exclude<TableCellValue, null> => value !== null && value !== undefined);

  if (aggregation.function === "sum" || aggregation.function === "avg") {
    const numeric = values
      .filter((value) => typeof value === "number" || typeof value === "bigint")
      .map((value) => Number(value));

    if (numeric.length === 0) return null;
    const sum = numeric.reduce((total, value) => total + value, 0);
    return aggregation.function === "sum" ? sum : sum / numeric.length;
  }

  if (values.length === 0) return null;
  return values.reduce((current, value) => {
    const comparison = compareValues(value, current);
    return aggregation.function === "min"
      ? (comparison < 0 ? value : current)
      : (comparison > 0 ? value : current);
  });
}

function orderRows(
  headers: readonly string[],
  rows: readonly TableCellValue[][],
  orderBy: readonly TableQueryOrder[],
): TableCellValue[][] {
  const orders = orderBy.map((order) => ({
    index: requiredColumnIndex(headers, order.attribute),
    direction: order.direction,
  }));

  return [...rows].sort((left, right) => {
    for (const order of orders) {
      const comparison = compareValues(left[order.index], right[order.index]);
      if (comparison !== 0) return order.direction === "asc" ? comparison : -comparison;
    }
    return 0;
  });
}

function paginate(
  rows: readonly TableCellValue[][],
  offset = 0,
  limit?: number,
): TableCellValue[][] {
  if (!Number.isInteger(offset) || offset < 0) {
    throw new RangeError("Table query offset must be a non-negative integer.");
  }
  if (limit !== undefined && (!Number.isInteger(limit) || limit < 0)) {
    throw new RangeError("Table query limit must be a non-negative integer.");
  }

  return rows.slice(offset, limit === undefined ? undefined : offset + limit);
}

function project(state: QueryState, select: readonly string[]): QueryState {
  const indexes = select.map((attribute) => requiredColumnIndex(state.headers, attribute));
  return {
    headers: [...select],
    rows: state.rows.map((row) => indexes.map((index) => row[index] ?? null)),
  };
}

function valueByHeader(
  headers: readonly string[],
  row: readonly TableCellValue[],
  header: string,
): TableCellValue | undefined {
  const index = headers.indexOf(header);
  return index < 0 ? undefined : row[index];
}

function requiredColumnIndex(headers: readonly string[], attribute: string): number {
  const index = headers.indexOf(attribute);
  if (index < 0) throw new Error(`Unknown table attribute '${attribute}'.`);
  return index;
}

function keyPart(value: TableCellValue | undefined): string {
  if (value === null || value === undefined) return "null";
  if (value instanceof Date) return `date:${value.getTime()}`;
  return `${typeof value}:${String(value)}`;
}

function scalar(
  value: TableCellValue | readonly TableCellValue[] | undefined,
): TableCellValue | undefined {
  return Array.isArray(value) ? undefined : value as TableCellValue | undefined;
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

function equal(left: unknown, right: unknown): boolean {
  if (left instanceof Date && right instanceof Date) return left.getTime() === right.getTime();
  return left === right;
}

function compareValues(left: unknown, right: unknown): number {
  if (left === null || left === undefined) return right === null || right === undefined ? 0 : -1;
  if (right === null || right === undefined) return 1;

  const a = left instanceof Date ? left.getTime() : left;
  const b = right instanceof Date ? right.getTime() : right;

  if (typeof a === "string" && typeof b === "string") return a.localeCompare(b);
  if ((typeof a === "number" || typeof a === "bigint") && (typeof b === "number" || typeof b === "bigint")) {
    const leftNumber = Number(a);
    const rightNumber = Number(b);
    return leftNumber === rightNumber ? 0 : leftNumber > rightNumber ? 1 : -1;
  }
  if (typeof a === "boolean" && typeof b === "boolean") return Number(a) - Number(b);

  return String(a).localeCompare(String(b));
}
