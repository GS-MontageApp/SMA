/**
 * Modul: Excel-Parser für Tabelle1 (Schlauchmanagement PWA)
 * Beschreibung: Extrahiert Metadaten (Kunde, Anlage) und filtert/parsiert 
 * die Schlauch-Datensätze aus "Tabelle1" gemäß Master-Pflichtenheft.
 */

class HoseExcelParser {
    /**
     * Normalisiert einen Tabellen-Header (entfernt Zeilenumbrüche und überschüssige Leerzeichen).
     * @param {string} headerStr 
     * @returns {string}
     */
    static normalizeHeader(headerStr) {
        if (!headerStr) return "";
        return String(headerStr)
            .replace(/[\r\n]+/g, " ") // Ersetzt Zeilenumbrüche durch Leerzeichen
            .replace(/\s+/g, " ")     // Mehrfache Leerzeichen reduzieren
            .trim();
    }

    /**
     * Durchsucht das geladene Workbook nach "Tabelle1" und extrahiert die Daten.
     * @param {Object} workbook - Das eingelesene SheetJS / XLSX Workbook-Objekt
     * @returns {Object} { kunde, anlage, recordsCount, records }
     */
    static parseTabelle1(workbook) {
        // 1. Strenge Prüfung auf Existenz von "Tabelle1" (ohne Leerzeichen)
        if (!workbook.SheetNames.includes("Tabelle1")) {
            throw new Error("Kritischer Fehler: Das Arbeitsblatt 'Tabelle1' wurde in der Datei nicht gefunden.");
        }

        const sheet = workbook.Sheets["Tabelle1"];
        // Konvertiere das Blatt in ein 2D-Array (Rohdaten)
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

        let kunde = "";
        let anlage = "";
        let headerRowIndex = -1;
        let columnMapping = {};

        // 2. Metadaten-Extraktion (Suche nach "Kunde" und "Anlage") & Header-Erkennung
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            for (let j = 0; j < row.length; j++) {
                const cellVal = String(row[j]).trim();

                // Suche nach "Kunde" -> Wert steht in der Zelle rechts daneben
                if (cellVal === "Kunde" && j + 1 < row.length) {
                    kunde = String(row[j + 1]).trim();
                }

                // Suche nach "Anlage" -> Wert steht in der Zelle rechts daneben
                if (cellVal === "Anlage" && j + 1 < row.length) {
                    anlage = String(row[j + 1]).trim();
                }
            }

            // Erkennung der Kopfzeile anhand der Pflichtbegriffe "Kennz." und "Schlauch"
            const rowStringJoined = row.map(cell => this.normalizeHeader(cell)).join(" ");
            if (rowStringJoined.includes("Kennz.") && rowStringJoined.includes("Schlauch")) {
                headerRowIndex = i;
                
                // Mache die Spalten-Indizes anhand der normalisierten Header ausfindig
                row.forEach((colHeader, colIndex) => {
                    const normalized = this.normalizeHeader(colHeader);
                    if (normalized.includes("Kennz.")) columnMapping.kennz = colIndex;
                    if (normalized.includes("Schlauch")) columnMapping.schlauch = colIndex;
                    if (normalized === "NW") columnMapping.nw = colIndex;
                    if (normalized === "Anschluß A") columnMapping.anschlussA = colIndex;
                    if (normalized === "Anschluß B") columnMapping.anschlussB = colIndex;
                    if (normalized === "Länge") columnMapping.laenge = colIndex;
                    if (normalized === "Lage A") columnMapping.lageA = colIndex;
                    if (normalized === "Lage B") columnMapping.lageB = colIndex;
                    if (normalized.includes("max. Druck")) columnMapping.druck = colIndex;
                    if (normalized.includes("Herstell") && normalized.includes("datum")) columnMapping.herstelldatum = colIndex;
                    if (normalized.includes("Sicherheits") && normalized.includes("Bewertung")) columnMapping.sicherheit = colIndex;
                    if (normalized.includes("Theor.") && normalized.includes("Lebens")) columnMapping.lebensdauer = colIndex;
                    if (normalized.includes("Prüfung*")) columnMapping.pruefungStatus = colIndex;
                    if (normalized === "Nächste Prüfung") columnMapping.naechstePruefung = colIndex;
                    if (normalized === "Prüfer") columnMapping.pruefer = colIndex;
                    if (normalized === "Einbauort") columnMapping.einbauort = colIndex;
                    if (normalized === "Bemerkung") columnMapping.bemerkung = colIndex;
                });
                break;
            }
        }

        if (headerRowIndex === -1) {
            throw new Error("Kritischer Fehler: Die Spaltenköpfe (Header) konnten in 'Tabelle1' nicht identifiziert werden.");
        }

        // Separat prüfen für exakte Spalte "Prüfung am"
        rows[headerRowIndex].forEach((colHeader, colIndex) => {
            if (this.normalizeHeader(colHeader) === "Prüfung am") columnMapping.pruefungAm = colIndex;
        });

        const records = [];

        // 3. Zeilenweises Einlesen mit robustem Filter (nur Zeilen mit gültiger Nummer in Spalte "Kennz.")
        for (let i = headerRowIndex + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const kennzVal = columnMapping.kennz !== undefined ? String(row[columnMapping.kennz]).trim() : "";

            // Filter-Regel: Hat die Zeile keine Nummer (leer oder beginnt nicht mit Ziffer), wird sie übersprungen.
            // Das erlaubt vollständigen Durchsuch des Dokuments trotz Lücken (z.B. fehlende Nummern).
            if (!kennzVal || !/^\d/.test(kennzVal)) {
                continue;
            }

            // Datensatz auslesen
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
