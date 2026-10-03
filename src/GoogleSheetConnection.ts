import { GoogleAuth, type JWTInput } from "google-auth-library";
import { google, type sheets_v4 } from "googleapis";

const WRITE_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const GOOGLE_SHEETS_PATH = /^\/spreadsheets\/(?:u\/\d+\/)?d\/([A-Za-z0-9_-]+)(?:\/|$)/;
const DEFAULT_READ_CACHE_TTL_MS = 60_000;

type SharedReadCacheEntry<T> = {
  promise: Promise<T>;
  settledAt?: number;
};

/**
 * Prozessweiter Cache: verschiedene Connection-/Workbook-/Table-Instanzen
 * teilen Reads, solange Spreadsheet und Service-Account identisch sind.
 */
const sharedReadCache = new Map<string, SharedReadCacheEntry<unknown>>();

export interface GoogleSheetConnectionOptions {
  /**
   * Cache-Lebensdauer in Millisekunden.
   *
   * - undefined: Standardwert 60 Sekunden
   * - number: Eintrag läuft nach dieser Zeit beim nächsten Read ab
   * - null: kein automatischer Ablauf; nur clearCache() oder ein Write invalidiert
   */
  cacheTtlMs?: number | null;
}

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
 *
 * Reads werden standardmäßig 60 Sekunden pro Prozess gecacht. Mit
 * cacheTtlMs: null bleibt ein Read unbegrenzt im Cache, bis clearCache() oder
 * ein Write ihn invalidiert. Es läuft kein Hintergrund-Timer.
 *
 * Parallele identische Reads teilen denselben laufenden Request.
 */
export class GoogleSheetConnection {
  private readonly spreadsheetId: string;
  private readonly api: sheets_v4.Sheets;
  private readonly cacheNamespace: string;
  private readonly cachePolicyKey: string;
  private readonly cacheTtlMs: number | null;

  /** Konfiguriert den Transport, ohne bereits einen Google-Request auszuführen. */
  public constructor(
    sheetUrl: string,
    credentials: JWTInput,
    options: GoogleSheetConnectionOptions = {},
  ) {
    this.spreadsheetId = spreadsheetIdFromUrl(sheetUrl);
    this.cacheTtlMs = options.cacheTtlMs === undefined
      ? DEFAULT_READ_CACHE_TTL_MS
      : options.cacheTtlMs;
    if (
      this.cacheTtlMs !== null
      && (!Number.isFinite(this.cacheTtlMs) || this.cacheTtlMs < 0)
    ) {
      throw new RangeError("cacheTtlMs must be null or a non-negative finite number.");
    }

    const principal = credentials.client_email ?? credentials.client_id ?? "default";
    this.cacheNamespace = `${this.spreadsheetId}\u0000${principal}\u0000`;
    this.cachePolicyKey = this.cacheTtlMs === null
      ? "manual"
      : `ttl:${this.cacheTtlMs}`;

    const auth = new GoogleAuth({ credentials, scopes: [WRITE_SCOPE] });
    this.api = google.sheets({ version: "v4", auth });
  }

  public getSpreadsheetId(): string { return this.spreadsheetId; }

  /** null bedeutet: kein automatischer Ablauf des Read-Caches. */
  public getCacheTtlMs(): number | null { return this.cacheTtlMs; }

  /**
   * Verwirft alle gecachten Reads dieses Spreadsheets für den verwendeten
   * Service Account. Der nächste Read geht wieder an Google.
   */
  public clearCache(): void {
    for (const key of sharedReadCache.keys()) {
      if (key.startsWith(this.cacheNamespace)) sharedReadCache.delete(key);
    }
  }

  /** Sendet einen vom Aufrufer beschriebenen Spreadsheet-Read. */
  public async read(request: SpreadsheetReadRequest = {}): Promise<sheets_v4.Schema$Spreadsheet> {
    const cacheKey = this.cacheKey("spreadsheet", {
      includeGridData: request.includeGridData ?? null,
      ranges: request.ranges === undefined ? null : [...request.ranges],
      fields: request.fields ?? null,
    });

    return this.cachedRead(cacheKey, async () => {
      const params: sheets_v4.Params$Resource$Spreadsheets$Get = {
        spreadsheetId: this.spreadsheetId,
      };
      if (request.includeGridData !== undefined) params.includeGridData = request.includeGridData;
      if (request.ranges !== undefined) params.ranges = [...request.ranges];
      if (request.fields !== undefined) params.fields = request.fields;
      const response = await this.api.spreadsheets.get(params);
      return response.data;
    });
  }

  /** Sendet einen vom Aufrufer beschriebenen Values-Read. */
  public async readValues(request: SpreadsheetValueReadRequest): Promise<unknown[][]> {
    const cacheKey = this.cacheKey("values", {
      range: request.range,
      majorDimension: request.majorDimension ?? null,
      valueRenderOption: request.valueRenderOption ?? null,
      dateTimeRenderOption: request.dateTimeRenderOption ?? null,
    });

    return this.cachedRead(cacheKey, async () => {
      const params: sheets_v4.Params$Resource$Spreadsheets$Values$Get = {
        spreadsheetId: this.spreadsheetId,
        range: request.range,
      };
      if (request.majorDimension !== undefined) params.majorDimension = request.majorDimension;
      if (request.valueRenderOption !== undefined) params.valueRenderOption = request.valueRenderOption;
      if (request.dateTimeRenderOption !== undefined) params.dateTimeRenderOption = request.dateTimeRenderOption;
      const response = await this.api.spreadsheets.values.get(params);
      return response.data.values ?? [];
    });
  }

  /** Sendet die vom Aufrufer zusammengestellten Strukturänderungen. */
  public async update(request: SpreadsheetUpdateRequest): Promise<void> {
    if (request.requests.length === 0) return;
    await this.withCacheInvalidation(async () => {
      await this.api.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: { requests: [...request.requests] },
      });
    });
  }

  /** Leert den vom Aufrufer festgelegten Range. */
  public async clearValues(request: SpreadsheetValueClearRequest): Promise<void> {
    await this.withCacheInvalidation(async () => {
      await this.api.spreadsheets.values.clear({
        spreadsheetId: this.spreadsheetId,
        range: request.range,
      });
    });
  }

  /** Sendet einen vollständig vom Aufrufer beschriebenen Values-Write. */
  public async writeValues(request: SpreadsheetValueWriteRequest): Promise<void> {
    await this.withCacheInvalidation(async () => {
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
    });
  }

  private cacheKey(kind: string, request: object): string {
    return `${this.cacheNamespace}${this.cachePolicyKey}\u0000${kind}\u0000${JSON.stringify(request)}`;
  }

  private async cachedRead<T>(key: string, loader: () => Promise<T>): Promise<T> {
    const existing = sharedReadCache.get(key) as SharedReadCacheEntry<T> | undefined;
    const now = Date.now();

    if (
      existing
      && (
        existing.settledAt === undefined
        || this.cacheTtlMs === null
        || now - existing.settledAt < this.cacheTtlMs
      )
    ) {
      return cloneCacheValue(await existing.promise);
    }

    if (existing) sharedReadCache.delete(key);

    const entry: SharedReadCacheEntry<T> = {
      promise: Promise.resolve().then(loader),
    };

    entry.promise = entry.promise.then(
      (value) => {
        entry.settledAt = Date.now();
        return value;
      },
      (error) => {
        if (sharedReadCache.get(key) === entry) sharedReadCache.delete(key);
        throw error;
      },
    );

    sharedReadCache.set(key, entry as SharedReadCacheEntry<unknown>);
    return cloneCacheValue(await entry.promise);
  }

  private async withCacheInvalidation(action: () => Promise<void>): Promise<void> {
    // Vorher invalidieren, damit niemand während eines Writes bewusst alte
    // Daten aus dem Cache bekommt. Danach erneut, um einen Read zu entfernen,
    // der parallel zum Write gestartet wurde.
    this.clearCache();
    try {
      await action();
    } finally {
      this.clearCache();
    }
  }
}

function cloneCacheValue<T>(value: T): T {
  return structuredClone(value);
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
