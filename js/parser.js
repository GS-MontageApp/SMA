/**
 * @file parser.js
 * @description Zentrales Parser-Modul für die Schlauchmanagement PWA.
 * Implementiert die strikten Vorgaben aus dem Master-Pflichtenheft:
 * - Strikte Tabellenblatt-Einschränkung auf "Tabelle1" (ohne Leerzeichen)[cite: 1].
 * - Metadaten-Extraktion (Kunde & Anlage)[cite: 1].
 * - Header-Normalisierung (Tolerierung von Zeilenumbrüchen).
 * - Vollständiger Scan mit robustem Kennz.-Zeilenfilter[cite: 1].
 */

class HoseExcelParser {
    /**
     * Normalisiert einen Tabellen-Header (entfernt Zeilenumbrüche und überschüssige Leerzeichen).
     */
    static normalizeHeader(headerStr) {
        if (!headerStr) return "";
        return String(headerStr)
            .replace(/[\r\n]+/g, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    /**
     * Strenge Validierung: Prüft, ob exakt das Arbeitsblatt "Tabelle1" existiert.
     */
    static validateAndGetSheet(workbook) {
        if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) {
            throw new Error("Kritischer Fehler: Das Excel-Workbook ist leer oder ungültig.");
        }

        const targetSheetName = "Tabelle1";

        if (!workbook.SheetNames.includes(targetSheetName)) {
            throw new Error(
                `Sicherheits-Abbruch: Das obligatorische Arbeitsblatt '${targetSheetName}' ` +
                `(exakter Name ohne Leerzeichen) wurde in der hochgeladenen Datei nicht gefunden. ` +
                `Vorhandene Blätter: [${workbook.SheetNames.join(", ")}].`
            );
        }

        return workbook.Sheets[targetSheetName];
    }

    /**
     * Hauptparsing-Funktion für Tabelle1.
     */
    static parseTabelle1(workbook) {
        const sheet = this.validateAndGetSheet(workbook);
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

        if (!rows || rows.length === 0) {
            throw new Error("Kritischer Fehler: Das Arbeitsblatt 'Tabelle1' enthält keine Daten.");
        }

        let kunde = "";
        let anlage = "";
        let headerRowIndex = -1;
        let columnMapping = {};

        // 1. Metadaten-Extraktion & Header-Erkennung
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            for (let j = 0; j < row.length; j++) {
                const cellVal = String(row[j]).trim();
                if (cellVal === "Kunde" && j + 1 < row.length) {
                    kunde = String(row[j + 1]).trim();
                }
                if (cellVal === "Anlage" && j + 1 < row.length) {
                    anlage = String(row[j + 1]).trim();
                }
            }

            const rowStringJoined = row.map(cell => this.normalizeHeader(cell)).join(" ");
            if (rowStringJoined.includes("Kennz.") && rowStringJoined.includes("Schlauch")) {
                headerRowIndex = i;

                row.forEach((colHeader, colIndex) => {
                    const norm = this.normalizeHeader(colHeader);
                    if (norm.includes("Kennz.")) columnMapping.kennz = colIndex;
                    if (norm.includes("Schlauch")) columnMapping.schlauch = colIndex;
                    if (norm === "NW") columnMapping.nw = colIndex;
                    if (norm === "Anschluß A") columnMapping.anschlussA = colIndex;
                    if (norm === "Anschluß B") columnMapping.anschlussB = colIndex;
                    if (norm === "Länge") columnMapping.laenge = colIndex;
                    if (norm === "Lage A") columnMapping.lageA = colIndex;
                    if (norm === "Lage B") columnMapping.lageB = colIndex;
                    if (norm.includes("max. Druck")) columnMapping.druck = colIndex;
                    if (norm.includes("Herstell") && norm.includes("datum")) columnMapping.herstelldatum = colIndex;
                    if (norm.includes("Sicherheits") && norm.includes("Bewertung")) columnMapping.sicherheit = colIndex;
                    if (norm.includes("Theor.") && norm.includes("Lebens")) columnMapping.lebensdauer = colIndex;
                    if (norm.includes("Prüfung*")) columnMapping.pruefungStatus = colIndex;
                    if (norm === "Nächste Prüfung") columnMapping.naechstePruefung = colIndex;
                    if (norm === "Prüfer") columnMapping.pruefer = colIndex;
                    if (norm === "Einbauort") columnMapping.einbauort = colIndex;
                    if (norm === "Bemerkung") columnMapping.bemerkung = colIndex;
                });
                break;
            }
        }

        if (headerRowIndex === -1) {
            throw new Error("Kritischer Fehler: Die Spaltenköpfe (Header) konnten in 'Tabelle1' nicht identifiziert werden.");
        }

        // Exakte Spalte "Prüfung am" ermitteln
        rows[headerRowIndex].forEach((colHeader, colIndex) => {
            if (this.normalizeHeader(colHeader) === "Prüfung am") columnMapping.pruefungAm = colIndex;
        });

        const records = [];

        // 2. Vollständiger Scan mit robustem Zeilen-Filter (Kennz.-Prüfung)
        for (let i = headerRowIndex + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const kennzVal = columnMapping.kennz !== undefined ? String(row[columnMapping.kennz]).trim() : "";

            // Zeilen-Validierungsregel: Überspringe Zeilen ohne gültige Kennz.-Nummer[cite: 1]
            if (!kennzVal || !/^\d/.test(kennzVal)) {
                continue;
            }

            const record = {
                kennz: kennzVal,
                schlauch: columnMapping.schlauch !== undefined ? row[columnMapping.schlauch] : "",
                nw: columnMapping.nw !== undefined ? row[columnMapping.nw] : "",
                anschlussA: columnMapping.anschlussA !== undefined ? row[columnMapping.anschlussA] : "",
                anschlussB: columnMapping.anschlussB !== undefined ? row[columnMapping.anschlussB] : "",
                laenge: columnMapping.laenge !== undefined ? row[columnMapping.laenge] : "",
                lageA: columnMapping.lageA !== undefined ? row[columnMapping.lageA] : "",
                lageB: columnMapping.lageB !== undefined ? row[columnMapping.lageB] : "",
                druck: columnMapping.druck !== undefined ? row[columnMapping.druck] : "",
                herstelldatum: columnMapping.herstelldatum !== undefined ? row[columnMapping.herstelldatum] : "",
                sicherheit: columnMapping.sicherheit !== undefined ? row[columnMapping.sicherheit] : "",
                lebensdauer: columnMapping.lebensdauer !== undefined ? row[columnMapping.lebensdauer] : "",
                pruefungAm: columnMapping.pruefungAm !== undefined ? row[columnMapping.pruefungAm] : "",
                pruefungStatus: columnMapping.pruefungStatus !== undefined ? row[columnMapping.pruefungStatus] : "",
                naechstePruefung: columnMapping.naechstePruefung !== undefined ? row[columnMapping.naechstePruefung] : "",
                pruefer: columnMapping.pruefer !== undefined ? row[columnMapping.pruefer] : "",
                einbauort: columnMapping.einbauort !== undefined ? row[columnMapping.einbauort] : "",
                bemerkung: columnMapping.bemerkung !== undefined ? row[columnMapping.bemerkung] : ""
            };

            records.push(record);
        }

        return {
            kunde,
            anlage,
            recordsCount: records.length,
            records
        };
    }
}

// Global verfügbar machen
window.HoseExcelParser = HoseExcelParser;
