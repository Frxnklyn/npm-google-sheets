# @frxnklyn/google-sheets

Eine TypeScript-Implementierung der Excel- und Table-Interfaces aus
`@frxnklyn/datatypes` für die Google Sheets API v4.

Das Paket bietet lazy Workbook-, Sheet- und Cell-Referenzen, Lesen und Speichern
von Werten und Formeln, eine Table-Sicht mit lokalen Filtern sowie das Anlegen
und Entfernen von Sheets.

## Voraussetzungen

- Node.js 20 oder neuer
- ein Google-Cloud-Projekt mit aktivierter Google Sheets API
- ein Service Account oder ein anderer von `googleapis` unterstützter OAuth-Client
- Zugriff des verwendeten Kontos auf das Spreadsheet

## Installation

`@frxnklyn/datatypes` ist als Peer Dependency ab Version `0.1.0` eingetragen.
Für die lokale Entwicklung verweist die Dev Dependency auf den benachbarten
Ordner `../npm-datatypes`:

```text
NPM Packages/
├── npm-datatypes/
└── npm-google-sheets/
```

```bash
npm install
npm run build
```

Der lokale `file:`-Eintrag wird nicht an Nutzer installiert; veröffentlichte
Pakete verwenden die semantische Peer Dependency.

## Service-Account-Key erstellen

Für `GoogleSheetsDataType` werden die JSON-Zugangsdaten eines Service Accounts
verwendet. Der Service Account ist ein technischer Google-Nutzer, dem das Sheet
explizit freigegeben wird.

1. In der [Google Cloud Console](https://console.cloud.google.com/) ein Projekt
   erstellen oder auswählen.
2. Unter **APIs & Dienste > Bibliothek** die **Google Sheets API** suchen und
   aktivieren.
3. Unter **IAM & Verwaltung > Dienstkonten** einen Service Account erstellen.
4. Den Service Account öffnen und **Schlüssel > Schlüssel hinzufügen > Neuen
   Schlüssel erstellen > JSON** wählen.
5. Die heruntergeladene JSON-Datei sicher speichern. Ein privater Schlüssel
   kann später nicht erneut heruntergeladen werden; bei Verlust muss ein neuer
   erstellt werden.
6. Im gewünschten Google Sheet auf **Freigeben** klicken und der
   `client_email` aus der JSON-Datei mindestens Lesezugriff geben.

Die Datei enthält einen privaten Schlüssel. Sie darf nicht committet,
veröffentlicht oder an den Browser ausgeliefert werden. Lege zum Beispiel
`credentials.json` in `.gitignore` ab oder übergib die Werte über eine sichere
Environment Variable.

## Spreadsheet öffnen

Die öffentliche Architektur trennt Google-Kommunikation und DataType-Logik:

```text
Google-Sheets-URL + Credentials
              ↓
GoogleSheetConnection
              ↓
GoogleSheetsDataType
              ↓
GoogleSheet / Cells / Tables
```

Die Connection erhält den vollständigen normalen Google-Sheets-Link und
extrahiert die Spreadsheet-ID selbst:

```ts
import credentials from "./credentials.json" with { type: "json" };
import {
  GoogleSheetConnection,
  GoogleSheetsDataType,
} from "@frxnklyn/google-sheets";

const connection = new GoogleSheetConnection(
  "https://docs.google.com/spreadsheets/d/1ABCDEF123456/edit#gid=0",
  credentials,
);

const workbook = new GoogleSheetsDataType(connection);

await workbook.dataRead();
console.log(workbook.getName());
console.log(workbook.getSheets().map((sheet) => sheet.getName()));
```

`GoogleSheetConnection` akzeptiert einen vollständigen HTTPS-Link unter
`docs.google.com/spreadsheets/...`. Leere Links, fremde Hosts und
Veröffentlichungslinks mit `/d/e/...` werden abgelehnt. Die separat exportierte
Utility `extractSpreadsheetId()` bleibt für Fälle verfügbar, in denen eine ID
außerhalb der Connection benötigt wird.

Die Spreadsheet-ID ist der Teil zwischen `/d/` und dem nächsten `/`:

```text
https://docs.google.com/spreadsheets/d/1ABCDEF123456/edit
                                        ^^^^^^^^^^^^^^^
```

Die Connection verwendet den Scope
`https://www.googleapis.com/auth/spreadsheets` für Lesen und Schreiben. Der
Service Account benötigt zusätzlich die passende Freigabe im Spreadsheet.
`GoogleSheetsDataType` und `GoogleSheet` erstellen keine eigene
Authentifizierung und greifen nicht direkt auf den Google-Client zu.

Die Connection stellt nur einheitliche technische Request-Primitiven bereit:

```ts
connection.read({ includeGridData: true });
connection.readValues({ range: "'Produkte'" });
connection.update({ requests });
connection.clearValues({ range: "'Produkte'" });
connection.writeValues({
  range: "'Produkte'",
  values,
  valueInputOption: "USER_ENTERED",
});
```

Die Connection ergänzt dabei nur die bekannte `spreadsheetId` und sendet die
Anfrage. Ob Grid-Daten, Metadaten, ein bestimmter Range oder ein bestimmter
Schreibmodus gebraucht wird, entscheidet die jeweils aufrufende Fachklasse.

## Cells lesen und schreiben

```ts
const sheet = workbook.getSheet("Produkte");
await sheet.dataRead();

console.log(sheet.getCell("B2").getValue());

sheet
  .setCell("A2", "Tastatur")
  .setCell(1, 1, 49.9)
  .setFormula("C2", "=B2*1.19");

await sheet.dataSave();
```

Indizes sind nullbasiert: `(0, 0)` entspricht `A1`. Eine lazy Cell liefert vor
dem Laden `isResolved() === false`; eine geladene, tatsächlich leere Cell den
Wert `null`.

## Als Table verwenden

Die erste Zeile eines Sheets wird als Header-Zeile interpretiert. Alle weiteren
Zeilen sind Datenzeilen. Die Table lädt dafür direkt Werte über
`connection.readValues()`; sie fordert nicht das vollständige Cell-/Grid-Modell
eines Sheets an.

```ts
import { GoogleSheetsFilter } from "@frxnklyn/google-sheets";

const table = workbook.getTable("Produkte");
table
  .addFilter(new GoogleSheetsFilter("Preis", "greaterThan", 20))
  .setComparator("and");

await table.dataRead();

for (const row of table.getRows()) {
  console.log(row.getValue("Name"), row.getValue("Preis"));
}

table.addRow(["Maus", 19.9, null]);
await table.dataSave();
```

Filter werden nach dem Google-Sheets-Read lokal ausgewertet. Ein anschließendes
`dataSave()` speichert genau die aktuell sichtbaren Table-Rows. Eine gefilterte
Table sollte daher nur gespeichert werden, wenn dieses Ersetzen beabsichtigt ist.

## Sheets verwalten

```ts
workbook.addSheet("Archiv");
workbook.getSheet("Archiv").setCell("A1", "Status");
await workbook.dataSave();

workbook.removeSheet("Archiv");
await workbook.dataSave();
```

Google Sheets verlangt mindestens ein sichtbares Sheet. Der Versuch, das letzte
Sheet zu entfernen, wird von der Google API abgelehnt.

## Scripts

```bash
npm run build
npm test
```

Ein manueller Integrationstest kann mit einer lokalen Service-Account-Datei
und einem echten Spreadsheet ausgeführt werden:

```powershell
$env:GOOGLE_SERVICE_ACCOUNT_FILE = "C:\path\service-account.json"
$env:GOOGLE_SHEETS_URL = "https://docs.google.com/spreadsheets/d/.../edit"
npm run test:integration
```

## Verhalten beim Speichern

Werte und Formeln werden mit `USER_ENTERED` geschrieben. `Date` wird als ISO-
String und `bigint` als String übertragen. Zellformatierungen und Merge-Bereiche
werden gelesen, von diesem Paket aber nicht verändert.

Hinweise zu den unveränderten DataTypes-Verträgen stehen in
[`INTERFACE_NOTES.md`](./INTERFACE_NOTES.md).

## Lizenz

MIT
