const GOOGLE_HOST_PATTERN = /(^|\.)google\.com$/i;
const SPREADSHEET_ID_PATTERN = /^[A-Za-z0-9_-]+$/;

/**
 * Extracts the spreadsheet id from a Google Sheets/Drive URL.
 * A plain spreadsheet id is accepted as well.
 *
 * Published `/d/e/...` URLs intentionally throw: their publication id is not
 * the spreadsheet id required by the Google Sheets API.
 */
export function extractSpreadsheetId(input: string | URL): string {
  const value = input instanceof URL ? input.href : input.trim();
  if (!value) throw new TypeError("A Google Sheets URL or spreadsheet id is required.");

  if (SPREADSHEET_ID_PATTERN.test(value)) return value;

  let url: URL;
  try {
    url = input instanceof URL ? input : new URL(value);
  } catch {
    throw new TypeError(`Invalid Google Sheets URL: ${value}`);
  }

  if (!GOOGLE_HOST_PATTERN.test(url.hostname)) {
    throw new TypeError(`The URL is not a Google URL: ${value}`);
  }

  const segments = url.pathname.split("/").filter(Boolean);
  const dIndex = segments.lastIndexOf("d");
  if (dIndex >= 0 && segments[dIndex + 1] === "e") {
    throw new TypeError("Published Google Sheets URLs do not contain the spreadsheet id required by the API.");
  }

  const pathId = dIndex >= 0 ? segments[dIndex + 1] : undefined;
  const queryId = url.searchParams.get("id") ?? undefined;
  const spreadsheetId = pathId ?? queryId;

  if (!spreadsheetId || !SPREADSHEET_ID_PATTERN.test(spreadsheetId)) {
    throw new TypeError(`No spreadsheet id was found in the Google URL: ${value}`);
  }

  return spreadsheetId;
}

/** @deprecated Use the more precise `extractSpreadsheetId` name. */
export const extractGoogleSheetsId = extractSpreadsheetId;
