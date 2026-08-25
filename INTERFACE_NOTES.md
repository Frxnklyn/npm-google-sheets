# Notizen zu `@frxnklyn/datatypes`

Stand: 23. August 2026

Für die aktuelle Implementierung waren **keine Änderungen oder Ergänzungen** an
den Interfaces in `npm-datatypes` erforderlich. Das Projekt `npm-datatypes`
wurde nicht bearbeitet.

Implementierungsspezifische Angaben wie URL, `spreadsheetId`, Credentials und
Google-Requests bleiben in `GoogleSheetConnection`; die interne numerische
`sheetId` bleibt in `GoogleSheet`. Sie gehören nicht in die allgemeinen
Excel-/Table-Verträge.

Mögliche spätere Erweiterungen müssen zuerst hier dokumentiert werden, bevor
eine Änderung an den gemeinsamen Interfaces vorgeschlagen wird.
