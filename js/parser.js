/**
 * ============================================================================
 * MODUL: parser.js (Schlauchmanagement-App v0.1.7)
 * ============================================================================
 * Verantwortlich für:
 * - Einlesen und Verarbeiten von .xls / .xlsx-Dateien via SheetJS
 * - Strenge Vorab-Validierung (Crash-Schutz & Dateigrößen-Check)
 * - Automatisches Extrahieren des Kundennamens (Schlüsselwort "Kunde")
 * - Bereitstellung der echten Tabellendaten für die Aktive Bearbeitungs-Bühne
 * ============================================================================
 */

window.ExcelParser = {
  /**
   * Validiert die Rohdatei vor dem Einlesen (Maximalgröße ca. 600 KB zum Schutz mobil. Browser).
   */
  validateFile: function(file) {
    if (!file) {
      throw new Error("Keine Datei ausgewählt.");
    }
    const validExtensions = ['.xls', '.xlsx'];
    const fileNameLower = file.name.toLowerCase();
    const isValidExt = validExtensions.some(ext => fileNameLower.endsWith(ext));
    
    if (!isValidExt) {
      throw new Error("Ungültiges Dateiformat. Bitte nur .xls oder .xlsx Dateien verwenden.");
    }

    // Limitprüfung (Max. ~650 KB)
    const maxSize = 650 * 1024;
    if (file.size > maxSize) {
      throw new Error("Die Datei ist zu groß (> 650 KB). Zum Schutz des Arbeitsspeichers mobiler Browser ist die Dateigröße limitiert.");
    }

    return true;
  },

  /**
   * Parst die Excel-Datei, sucht nach dem Kunden und extrahiert die Tabellenzeilen.
   */
  parseFileBuffer: function(arrayBuffer, fileName) {
    try {
      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      
      let foundCustomer = null;
      let extractedRows = [];

      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const jsonSheet = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        if (jsonSheet.length > 0 && extractedRows.length === 0) {
          extractedRows = jsonSheet;
        }

        jsonSheet.forEach(row => {
          row.forEach((cellVal, colIdx) => {
            if (cellVal && typeof cellVal === 'string' && cellVal.trim().toLowerCase() === 'kunde') {
              if (row[colIdx + 1] !== undefined && row[colIdx + 1] !== null) {
                foundCustomer = String(row[colIdx + 1]).trim();
              }
            }
          });
        });
      });

      if (!foundCustomer) {
        foundCustomer = fileName.replace(/\.[^/.]+$/, "");
      }

      if (extractedRows.length === 0) {
        extractedRows = [["Info", "Die Excel-Tabelle enthält keine lesbaren Daten."]];
      }

      return {
        client: foundCustomer,
        filename: fileName,
        rawData: extractedRows
      };

    } catch (err) {
      console.error("Parser-Fehler:", err);
      throw new Error("Fehler beim Einlesen der Excel-Struktur. Bitte prüfen Sie das Dateiformat.");
    }
  }
};
