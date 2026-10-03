import type {
  ExcelDataTypeInterface,
  ExcelSheetDataTypeInterface,
  TableQueryInterface,
  TableSchemaInterface,
  TableSourceInterface,
} from "@frxnklyn/datatypes";
import type { sheets_v4 } from "googleapis";
import { GoogleSheet } from "./GoogleSheet.js";
import type { GoogleSheetConnection } from "./GoogleSheetConnection.js";

/**
 * Fachliche Repräsentation eines vollständigen Google Spreadsheets.
 * Authentifizierung und Requests werden vollständig an die Connection delegiert.
 *
 * Gleichzeitig implementiert das Workbook den allgemeinen TableSource-Vertrag:
 * Ein Sheet entspricht dabei einer fachlichen Table. Dadurch kann ein Consumer
 * dieselben TableQueryInterface-Objekte später gegen Google Sheets, SQL oder
 * eine andere TableSource ausführen.
 */
export class GoogleSheetsDataType implements ExcelDataTypeInterface, TableSourceInterface {
  private name: string | undefined;
  private sheets: GoogleSheet[] = [];
  private readonly referencesByName = new Map<string, GoogleSheet>();
  private readonly referencesByIndex = new Map<number, GoogleSheet>();
  private readonly removedNames = new Set<string>();

  /** Erzeugt ein lazy Workbook; der Konstruktor führt keinen Google-Request aus. */
  public constructor(private readonly connection: GoogleSheetConnection) {}

  public getName(): string | undefined { return this.name; }
  public getSheets(): readonly GoogleSheet[] { return this.sheets; }

  /** Gibt die gemeinsame technische Connection für die enthaltenen Sheets zurück. */
  public getConnection(): GoogleSheetConnection { return this.connection; }

  public getSheet(name: string): GoogleSheet;
  public getSheet(index: number): GoogleSheet;
  /** Erstellt oder liefert eine lazy Referenz, ohne Google anzufragen. */
  public getSheet(reference: string | number): GoogleSheet {
    if (typeof reference === "string") {
      const existing = this.referencesByName.get(reference);
      if (existing) return existing;
      const sheet = new GoogleSheet(this, reference);
      this.referencesByName.set(reference, sheet);
      return sheet;
    }

    const existing = this.referencesByIndex.get(reference);
    if (existing) return existing;
    const sheet = new GoogleSheet(this, reference);
    this.referencesByIndex.set(reference, sheet);
    return sheet;
  }

  /**
   * Liefert eine lazy Table-Referenz. Eine optionale portable Query wird erst
   * beim späteren dataRead() ausgewertet.
   */
  public getTable(
    sheetName: string,
    query?: TableQueryInterface,
  ): ReturnType<GoogleSheet["asTable"]> {
    return this.getSheet(sheetName).asTable(query);
  }

  /** Prüft über einen reinen Metadata-Read, ob das Sheet existiert. */
  public async hasTable(name: string): Promise<boolean> {
    const metadata = await this.fetchMetadata();
    return metadata.sheets?.some((sheet) => sheet.properties?.title === name) ?? false;
  }

  /**
   * Legt ein neues Sheet aus einem allgemeinen TableSchema an. Die Attribute
   * werden in Tabellenreihenfolge als Header-Zeile geschrieben.
   */
  public async addTable(schema: TableSchemaInterface): Promise<ReturnType<GoogleSheet["asTable"]>> {
    const name = schema.getName();
    if (await this.hasTable(name)) {
      throw new Error(`Table '${name}' already exists.`);
    }

    const sheet = this.getSheet(name);
    await this.ensureSheetExists(sheet);
    schema.getAttributes().forEach((attribute, index) => {
      sheet.setCell(0, index, attribute.getName());
    });
    await sheet.dataSave();
    return sheet.asTable();
  }

  /** Entfernt ein Sheet unmittelbar aus der externen TableSource. */
  public async removeTable(name: string): Promise<void> {
    const metadata = await this.fetchMetadata();
    const remote = metadata.sheets?.find((sheet) => sheet.properties?.title === name);
    const sheetId = remote?.properties?.sheetId;

    if (sheetId === null || sheetId === undefined) return;

    await this.connection.update({ requests: [{ deleteSheet: { sheetId } }] });
    this.sheets = this.sheets.filter((sheet) => sheet.getName() !== name);
    this.referencesByName.delete(name);

    const index = remote?.properties?.index;
    if (index !== null && index !== undefined) this.referencesByIndex.delete(index);
    this.removedNames.delete(name);
  }

  public addSheet(sheet: string | ExcelSheetDataTypeInterface): this {
    let candidate: GoogleSheet;
    if (typeof sheet === "string") candidate = this.getSheet(sheet);
    else if (sheet instanceof GoogleSheet && sheet.getExcel() === this) candidate = sheet;
    else {
      const name = sheet.getName();
      if (!name) throw new Error("Cannot add a custom sheet whose name is unresolved.");
      candidate = this.getSheet(name);
      for (const cell of sheet.getCells()) {
        const formula = cell.getFormula();
        if (formula !== undefined) candidate.setFormula(cell.getRowIndex(), cell.getColumnIndex(), formula);
        else if (cell.isResolved()) candidate.setCell(cell.getRowIndex(), cell.getColumnIndex(), cell.getValue() ?? null);
      }
    }

    const name = candidate.getName();
    if (name) this.removedNames.delete(name);
    if (!this.sheets.includes(candidate)) this.sheets.push(candidate);
    return this;
  }

  public removeSheet(sheet: string | ExcelSheetDataTypeInterface): this {
    const name = typeof sheet === "string" ? sheet : sheet.getName();
    if (!name) throw new Error("Cannot remove a sheet whose name is unresolved.");
    this.removedNames.add(name);
    this.sheets = this.sheets.filter((entry) => entry.getName() !== name && entry !== sheet);
    return this;
  }

  /** Lädt das komplette Workbook inklusive Grid-Daten über die Connection. */
  public async dataRead(): Promise<this> {
    const spreadsheet = await this.connection.read({ includeGridData: true });
    this.name = spreadsheet.properties?.title ?? this.name;
    this.applySheets(spreadsheet.sheets ?? []);
    this.removedNames.clear();
    return this;
  }

  /** Ermittelt fachliche Strukturänderungen und lässt sie von der Connection senden. */
  public async dataSave(): Promise<this> {
    const metadata = await this.fetchMetadata();
    const remoteByName = new Map(
      (metadata.sheets ?? []).map((sheet) => [sheet.properties?.title, sheet]),
    );
    const requests: sheets_v4.Schema$Request[] = [];

    for (const name of this.removedNames) {
      const sheetId = remoteByName.get(name)?.properties?.sheetId;
      if (sheetId !== null && sheetId !== undefined) requests.push({ deleteSheet: { sheetId } });
    }
    for (const sheet of this.sheets) {
      const name = sheet.getName();
      if (name && !remoteByName.has(name)) requests.push({ addSheet: { properties: { title: name } } });
    }

    await this.connection.update({ requests });
    this.removedNames.clear();
    await this.refreshIdentities();
    for (const sheet of this.sheets) await sheet.dataSave();
    return this;
  }

  /** Löst eine indexbasierte lazy Referenz über einen fachlich bestimmten Metadata-Read auf. */
  public async resolveSheetName(sheet: GoogleSheet): Promise<string> {
    const knownName = sheet.getName();
    if (knownName !== undefined) return knownName;
    const metadata = await this.fetchMetadata();
    const name = metadata.sheets?.find(
      (entry) => entry.properties?.index === sheet.getIndex(),
    )?.properties?.title ?? undefined;
    if (name === undefined) throw new Error(`No sheet exists at index ${sheet.getIndex()}.`);
    return name;
  }

  /** Legt ein fachlich neues Sheet bei Google an, falls es noch nicht existiert. */
  public async ensureSheetExists(sheet: GoogleSheet): Promise<void> {
    const name = sheet.requireName();
    const metadata = await this.fetchMetadata();
    const existing = metadata.sheets?.find((entry) => entry.properties?.title === name);
    if (!existing) {
      await this.connection.update({ requests: [{ addSheet: { properties: { title: name } } }] });
      await this.refreshIdentities();
      if (!this.sheets.includes(sheet)) this.sheets.push(sheet);
    } else {
      this.registerSheet(existing, sheet);
    }
  }

  private async fetchMetadata(): Promise<sheets_v4.Schema$Spreadsheet> {
    const metadata = await this.connection.read({
      includeGridData: false,
      fields: "spreadsheetId,properties.title,sheets.properties",
    });
    this.name = metadata.properties?.title ?? this.name;
    return metadata;
  }

  private async refreshIdentities(): Promise<void> {
    const metadata = await this.fetchMetadata();
    for (const remote of metadata.sheets ?? []) this.registerSheet(remote);
  }

  private applySheets(remoteSheets: sheets_v4.Schema$Sheet[]): void {
    const loaded: GoogleSheet[] = [];
    for (const remote of remoteSheets) {
      const sheet = this.registerSheet(remote);
      sheet.hydrate(remote);
      loaded.push(sheet);
    }
    this.sheets = loaded;
  }

  /** Verknüpft eine Google-Antwort mit einer bestehenden oder neuen lazy Sheet-Referenz. */
  public registerSheet(remote: sheets_v4.Schema$Sheet, preferred?: GoogleSheet): GoogleSheet {
    const name = remote.properties?.title;
    const index = remote.properties?.index;
    const sheetId = remote.properties?.sheetId;
    if (name === null || name === undefined || index === null || index === undefined
      || sheetId === null || sheetId === undefined) {
      throw new Error("Google Sheets returned incomplete sheet metadata.");
    }

    const sheet = preferred
      ?? this.referencesByName.get(name)
      ?? this.referencesByIndex.get(index)
      ?? new GoogleSheet(this, name);
    sheet.updateIdentity(name, index, sheetId);
    this.referencesByName.set(name, sheet);
    this.referencesByIndex.set(index, sheetId === undefined ? sheet : sheet);
    this.referencesByIndex.set(index, sheet);
    if (!this.sheets.includes(sheet)) this.sheets.push(sheet);
    return sheet;
  }
}
