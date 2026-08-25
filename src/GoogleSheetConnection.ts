import { GoogleAuth, type JWTInput } from "google-auth-library";
import { google, type sheets_v4 } from "googleapis";

const WRITE_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const GOOGLE_SHEETS_PATH = /^\/spreadsheets\/(?:u\/\d+\/)?d\/([A-Za-z0-9_-]+)(?:\/|$)/;

export interface SpreadsheetReadRequest {
  includeGridData?: boolean;
  ranges?: readonly string[];
  fields?: string;
}

export interface SpreadsheetValueReadRequest {
  range: string;
  majorDimension?: "ROWS" | "COLUMNS";
  valueRenderOption?: "FORMATTED_VALUE" | "UNFORMATTED_VALUE" | "FORMULA";
  dateTimeRenderOption?: "SERIAL_NUMBER" | "FORMATTED_STRING";
}

export interface SpreadsheetUpdateRequest {
  requests: readonly sheets_v4.Schema$Request[];
}

export interface SpreadsheetValueClearRequest {
  range: string;
}

export interface SpreadsheetValueWriteRequest {
  range: string;
  values: readonly (readonly unknown[])[];
  valueInputOption: "RAW" | "USER_ENTERED";
  includeValuesInResponse?: boolean;
  responseValueRenderOption?: "FORMATTED_VALUE" | "UNFORMATTED_VALUE" | "FORMULA";
  responseDateTimeRenderOption?: "SERIAL_NUMBER" | "FORMATTED_STRING";
}

/**
 * Technischer Transport für genau ein Google Spreadsheet.
 * Die Connection kennt URL, Spreadsheet-ID, Authentifizierung und API-Client.
 * Inhalt und Optionen eines Requests werden vollständig vom Aufrufer bestimmt.
 */
export class GoogleSheetConnection {
  private readonly spreadsheetId: string;
  private readonly api: sheets_v4.Sheets;

  /** Konfiguriert den Transport, ohne bereits einen Google-Request auszuführen. */
  public constructor(sheetUrl: string, credentials: JWTInput) {
    this.spreadsheetId = spreadsheetIdFromUrl(sheetUrl);
    const auth = new GoogleAuth({ credentials, scopes: [WRITE_SCOPE] });
    this.api = google.sheets({ version: "v4", auth });
  }

  public getSpreadsheetId(): string { return this.spreadsheetId; }

  /** Sendet einen vom Aufrufer beschriebenen Spreadsheet-Read. */
  public async read(request: SpreadsheetReadRequest = {}): Promise<sheets_v4.Schema$Spreadsheet> {
    const params: sheets_v4.Params$Resource$Spreadsheets$Get = {
      spreadsheetId: this.spreadsheetId,
    };
    if (request.includeGridData !== undefined) params.includeGridData = request.includeGridData;
    if (request.ranges !== undefined) params.ranges = [...request.ranges];
    if (request.fields !== undefined) params.fields = request.fields;
    const response = await this.api.spreadsheets.get(params);
    return response.data;
  }

  /** Sendet einen vom Aufrufer beschriebenen Values-Read. */
  public async readValues(request: SpreadsheetValueReadRequest): Promise<unknown[][]> {
    const params: sheets_v4.Params$Resource$Spreadsheets$Values$Get = {
      spreadsheetId: this.spreadsheetId,
      range: request.range,
    };
    if (request.majorDimension !== undefined) params.majorDimension = request.majorDimension;
    if (request.valueRenderOption !== undefined) params.valueRenderOption = request.valueRenderOption;
    if (request.dateTimeRenderOption !== undefined) params.dateTimeRenderOption = request.dateTimeRenderOption;
    const response = await this.api.spreadsheets.values.get(params);
    return response.data.values ?? [];
  }

  /** Sendet die vom Aufrufer zusammengestellten Strukturänderungen. */
  public async update(request: SpreadsheetUpdateRequest): Promise<void> {
    if (request.requests.length === 0) return;
    await this.api.spreadsheets.batchUpdate({
      spreadsheetId: this.spreadsheetId,
      requestBody: { requests: [...request.requests] },
    });
  }

  /** Leert den vom Aufrufer festgelegten Range. */
  public async clearValues(request: SpreadsheetValueClearRequest): Promise<void> {
    await this.api.spreadsheets.values.clear({
      spreadsheetId: this.spreadsheetId,
      range: request.range,
    });
  }

  /** Sendet einen vollständig vom Aufrufer beschriebenen Values-Write. */
  public async writeValues(request: SpreadsheetValueWriteRequest): Promise<void> {
    const params: sheets_v4.Params$Resource$Spreadsheets$Values$Update = {
      spreadsheetId: this.spreadsheetId,
      range: request.range,
      valueInputOption: request.valueInputOption,
      requestBody: { values: request.values.map((row) => [...row]) },
    };
    if (request.includeValuesInResponse !== undefined) {
      params.includeValuesInResponse = request.includeValuesInResponse;
    }
    if (request.responseValueRenderOption !== undefined) {
      params.responseValueRenderOption = request.responseValueRenderOption;
    }
    if (request.responseDateTimeRenderOption !== undefined) {
      params.responseDateTimeRenderOption = request.responseDateTimeRenderOption;
    }
    await this.api.spreadsheets.values.update(params);
  }
}

function spreadsheetIdFromUrl(sheetUrl: string): string {
  const value = sheetUrl.trim();
  if (!value) throw new TypeError("sheetUrl is required.");

  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new TypeError(`Invalid Google Sheets URL: ${value}`);
  }

  if (url.protocol !== "https:" || url.hostname !== "docs.google.com") {
    throw new TypeError(`Invalid Google Sheets URL: ${value}`);
  }

  const spreadsheetId = GOOGLE_SHEETS_PATH.exec(url.pathname)?.[1];
  if (!spreadsheetId || spreadsheetId === "e") {
    throw new TypeError(`Invalid Google Sheets URL: ${value}`);
  }
  return spreadsheetId;
}
