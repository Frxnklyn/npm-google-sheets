# Notizen zu `@frxnklyn/datatypes`

Stand: 3. Oktober 2026

`npm-google-sheets` verwendet jetzt zusätzlich die portablen Table-Contracts aus
`@frxnklyn/datatypes`:

- `TableQueryInterface` für datenquellenunabhängige Queries
- `TableSourceInterface` als austauschbare Source-Abstraktion
- `TableSchemaInterface` für das Anlegen neuer Tables/Sheets

Dabei bleibt die Google-spezifische Technik außerhalb der gemeinsamen Contracts.
URL, `spreadsheetId`, Credentials und Google-Requests liegen weiterhin in
`GoogleSheetConnection`; die interne numerische `sheetId` bleibt in
`GoogleSheet`.

Ein Google Sheet wird auf Table-Ebene als fachliche Table behandelt. Der Ablauf
entspricht damit dem allgemeinen Contract:

```text
TableSourceInterface
  -> getTable(name, query)
     -> lazy TableDataTypeInterface
        -> dataRead()
```

Die portable Query wird von Google Sheets lokal auf den über die Values API
gelesenen Rows ausgeführt. Eine spätere SQL-Implementierung kann denselben
`TableQueryInterface`-Vertrag in parameterisiertes SQL übersetzen, ohne dass
Consumer ihre Query-Struktur ändern müssen.

Die bestehenden Excel-, Sheet-, Cell- und Legacy-Filter-APIs bleiben erhalten.
