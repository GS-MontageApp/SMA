/**
 * ============================================================================
 * MODUL: parser.js (Schlauchmanagement-App v0.1.25)
 * ============================================================================
 * Sucht strikt nach dem exakten Wort "Kunde" (mit großem K) und extrahiert
 * den Namen aus der direkt rechts daneben liegenden Zelle.
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
      let foundPlant = null;
      let rawRows = [];

      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const jsonSheet = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        if (jsonSheet.length > 0 && rawRows.length === 0) {
          rawRows = jsonSheet;
        }

        // Exakte Suche nach dem Wort "Kunde" (Großes K, exakter String-Vergleich ohne unscharfe Suffixe)
        jsonSheet.forEach(row => {
          row.forEach((cellVal, colIdx) => {
            if (cellVal !== undefined && cellVal !== null) {
              const cellStr = String(cellVal).trim();
              
              if (cellStr === "Kunde") {
                if (row[colIdx + 1] !== undefined && row[colIdx + 1] !== null) {
                  foundCustomer = String(row[colIdx + 1]).trim();
                }
              }

              if (cellStr === "Anlage" || cellStr === "Anlagenbezeichnung") {
                if (row[colIdx + 1] !== undefined && row[colIdx + 1] !== null) {
                  foundPlant = String(row[colIdx + 1]).trim();
                }
              }
            }
          });
        });
      });

      // Fallback für den Kundennamen, falls kein Label "Kunde" im Dokument gefunden wurde
      if (!foundCustomer) {
        foundCustomer = fileName.replace(/\.[^/.]+$/, "");
      }

      // Filterung & Strukturierung der Tabellenzeilen
      let filteredRows = [];
      let headerFound = false;

      rawRows.forEach(row => {
        const hasContent = row.some(cell => cell !== undefined && cell !== null && String(cell).trim() !== '');
        if (!hasContent) return;

        const rowString = row.join(' ').toLowerCase();
        if (rowString.includes('kunde') || rowString.includes('anlage')) {
          filteredRows.push(row);
          return;
        }

        if (!headerFound && (rowString.includes('id') || rowString.includes('typ') || rowString.includes('länge') || rowString.includes('druck'))) {
          headerFound = true;
          filteredRows.push(row);
          return;
        }

        if (headerFound) {
          filteredRows.push(row);
        }
      });

      if (filteredRows.length === 0) {
        filteredRows = rawRows.length > 0 ? rawRows : [["Info", "Die Excel-Tabelle enthält keine lesbaren Daten."]];
      }

      const finalClientName = foundPlant ? `${foundCustomer} (${foundPlant})` : foundCustomer;

      return {
        client: finalClientName,
        filename: fileName,
        rawData: filteredRows
      };

    } catch (err) {
      console.error("Parser-Fehler:", err);
      throw new Error("Fehler beim Einlesen der Excel-Struktur. Bitte prüfen Sie das Dateiformat.");
    }
  }
};
