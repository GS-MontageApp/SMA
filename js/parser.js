/**
 * @file parser.js
 * @description Zentrales Parser-Modul für die Schlauchmanagement PWA.
 * Implementiert die strikten Vorgaben aus dem Master-Pflichtenheft:
 * - Strikte Tabellenblatt-Einschränkung auf "Tabelle1" (ohne Leerzeichen)[cite: 1].
 */

class HoseExcelParser {
    /**
     * Überprüft das Workbook strikt auf das Vorhandensein von "Tabelle1".
     * Niemals wird blind das erste Tabellenblatt eingelesen.
     * 
     * @param {Object} workbook - Das via SheetJS eingelesene Workbook-Objekt
     * @returns {Object} Das validierte Worksheet von "Tabelle1"
     * @throws {Error} Harter Abbruch bei fehlendem oder falschem Tabellenblatt
     */
    static validateAndGetSheet(workbook) {
        if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) {
            throw new Error("Kritischer Fehler: Das Excel-Workbook ist leer oder ungültig.");
        }

        const targetSheetName = "Tabelle1";

        // Strenge Prüfung gemäß Pflichtenheft: Exakter Name ohne Leerzeichen
        if (!workbook.SheetNames.includes(targetSheetName)) {
            throw new Error(
                `Sicherheits-Abbruch: Das obligatorische Arbeitsblatt '${targetSheetName}' ` +
                `(exakter Name ohne Leerzeichen) wurde in der hochgeladenen Datei nicht gefunden. ` +
                `Vorhandene Blätter in der Datei: [${workbook.SheetNames.join(", ")}].`
            );
        }

        return workbook.Sheets[targetSheetName];
    }

    /**
     * Liest die Rohdaten aus dem validierten Tabellenblatt "Tabelle1" als 2D-Array ein.
     * 
     * @param {Object} workbook - Das Workbook-Objekt
     * @returns {Array<Array>} Zeilenbasiertes Array der Rohdaten
     */
    static parseRawData(workbook) {
        const sheet = this.validateAndGetSheet(workbook);
        
        // Konvertierung in ein 2D-Array (header: 1 erzeugt ein reines Array von Zeilen-Arrays)
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        
        if (!rows || rows.length === 0) {
            throw new Error("Kritischer Fehler: Das Arbeitsblatt 'Tabelle1' enthält keine Daten.");
        }

        return rows;
    }
}

// Export für die PWA-Modulstruktur
window.HoseExcelParser = HoseExcelParser;
