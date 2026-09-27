/**
 * ============================================================================
 * MODUL: parser.js (Schlauchmanagement-App v0.1.31)
 * ============================================================================
 * Liest die komplette Excel-Datei mit ALLEN Arbeitsblättern zur restlosen 
 * Datensicherung ein und extrahiert den Kundennamen.
 */

window.ExcelParser = {
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

    const maxSize = 650 * 1024;
    if (file.size > maxSize) {
      throw new Error("Die Datei ist zu groß (> 650 KB). Zum Schutz mobiler Browser limitiert.");
    }

    return true;
  },

  parseFileBuffer: function(arrayBuffer, fileName) {
    try {
      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      
      let foundCustomer = null;
      let allSheetsData = {};

      // Vollständiges Einlesen ALLER Arbeitsblätter (inklusive Auswahlseite) für den Hintergrund
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const jsonSheet = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        allSheetsData[sheetName] = jsonSheet;

        // Exakte Suche nach dem Wort "Kunde" (Großes K)
        jsonSheet.forEach(row => {
          row.forEach((cellVal, colIdx) => {
            if (cellVal !== undefined && cellVal !== null) {
              const cellStr = String(cellVal).trim();
              if (cellStr === "Kunde") {
                if (row[colIdx + 1] !== undefined && row[colIdx + 1] !== null) {
                  foundCustomer = String(row[colIdx + 1]).trim();
                }
              }
            }
          });
        });
      });

      // Fallback für den Kundennamen
      if (!foundCustomer) {
        foundCustomer = fileName.replace(/\.[^/.]+$/, "");
      }

      return {
        client: foundCustomer,
        filename: fileName,
        sheets: allSheetsData // Gesamte Datei mit allen Blättern im Hintergrund gesichert
      };

    } catch (err) {
      console.error("Parser-Fehler:", err);
      throw new Error("Fehler beim Einlesen der Excel-Struktur. Bitte prüfen Sie das Dateiformat.");
    }
  }
};
