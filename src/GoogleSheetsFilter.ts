import type {
  FilterInterface,
  FilterOperator,
  TableCellValue,
} from "@frxnklyn/datatypes";

export class GoogleSheetsFilter implements FilterInterface {
  public constructor(
    private readonly attribute: string,
    private readonly operator: FilterOperator,
    private readonly value?: TableCellValue | readonly TableCellValue[],
  ) {
    if (!attribute.trim()) throw new TypeError("attribute is required.");
  }

  public getAttribute(): string { return this.attribute; }
  public getOperator(): FilterOperator { return this.operator; }
  public getValue(): TableCellValue | readonly TableCellValue[] | undefined { return this.value; }
}
