import type { AttributeType, TableCellValue } from "@frxnklyn/datatypes";
import type { sheets_v4 } from "googleapis";

const DAY_IN_MS = 86_400_000;
const GOOGLE_EPOCH = Date.UTC(1899, 11, 30);

export function readCellValue(cell: sheets_v4.Schema$CellData): TableCellValue {
  const value = cell.effectiveValue;
  if (!value) return null;
  if (value.boolValue !== null && value.boolValue !== undefined) return value.boolValue;
  if (value.stringValue !== null && value.stringValue !== undefined) return value.stringValue;
  if (value.numberValue !== null && value.numberValue !== undefined) {
    const format = cell.effectiveFormat?.numberFormat?.type;
    if (format === "DATE" || format === "DATE_TIME" || format === "TIME") {
      return new Date(GOOGLE_EPOCH + value.numberValue * DAY_IN_MS);
    }
    return value.numberValue;
  }
  if (value.errorValue) return cell.formattedValue ?? null;
  return null;
}

export function readFormula(cell: sheets_v4.Schema$CellData): string | undefined {
  return cell.userEnteredValue?.formulaValue ?? undefined;
}

export function toApiValue(value: TableCellValue): string | number | boolean | null {
  if (typeof value === "bigint") return value.toString();
  if (value instanceof Date) return value.toISOString();
  return value;
}

export function inferType(values: readonly TableCellValue[]): AttributeType {
  const present = values.filter((value) => value !== null);
  if (present.length === 0) return "unknown";
  if (present.every((value) => typeof value === "boolean")) return "boolean";
  if (present.every((value) => value instanceof Date)) {
    return present.every((value) => value instanceof Date && value.getUTCHours() === 0 && value.getUTCMinutes() === 0 && value.getUTCSeconds() === 0 && value.getUTCMilliseconds() === 0)
      ? "date"
      : "datetime";
  }
  if (present.every((value) => typeof value === "bigint" || (typeof value === "number" && Number.isInteger(value)))) return "integer";
  if (present.every((value) => typeof value === "number" || typeof value === "bigint")) return "number";
  if (present.every((value) => typeof value === "string")) return "string";
  return "unknown";
}
