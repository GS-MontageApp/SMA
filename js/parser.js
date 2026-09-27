/**
 * ============================================================================
 * MODUL: parser.js (Schlauchmanagement-App v0.1.28)
 * ============================================================================
 * Liest die komplette Excel-Datei mit allen Arbeitsblättern zur Datensicherung ein
 * und bereitet das Arbeitsblatt "Tabelle1" für die grafische Ausgabe auf[cite: 1].
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
      let tabelle1Rows = [];

      // 1. Vollständiges Einlesen aller Arbeitsblätter zur Datensicherung[cite: 1]
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const jsonSheet = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        allSheetsData[sheetName] = jsonSheet;

        // Exakte Suche nach dem Wort "Kunde" (Großes K) -> Wert in der Zelle rechts daneben
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

        // Gezieltes Erfassen von "Tabelle1" für die grafische Ansicht[cite: 1]
        if (sheetName.toLowerCase() === "tabelle1") {
          tabelle1Rows = jsonSheet;
        }
      });

      // Fallback, falls "Tabelle1" abweichend benannt ist
      if (tabelle1Rows.length === 0 && workbook.SheetNames.length > 0) {
        tabelle1Rows = allSheetsData[workbook.SheetNames[0]];
      }

      // Fallback für den Kundennamen
      if (!foundCustomer) {
        foundCustomer = fileName.replace(/\.[^/.]+$/, "");
      }

      // 2. Filterung und Strukturierung exklusiv für "Tabelle1" (Grafische Ausgabe)[cite: 1]
      let filteredRows = [];
      let headerFound = false;

      tabelle1Rows.forEach(row => {
        const hasContent = row.some(cell => cell !== undefined && cell !== null && String(cell).trim() !== '');
        if (!hasContent) return;

        const rowString = row.join(' ').toLowerCase();
        if (rowString.includes('kunde') || rowString.includes('anlage')) {
          filteredRows.push(row);
          return;
        }

        if (!headerFound && (rowString.includes('id') || rowString.includes('typ') || rowString.includes('länge') || rowString.includes('druck') || rowString.includes('kennz'))) {
          headerFound = true;
          filteredRows.push(row);
          return;
        }

        if (headerFound) {
          filteredRows.push(row);
        }
      });

      if (filteredRows.length === 0) {
        filteredRows = tabelle1Rows.length > 0 ? tabelle1Rows : [["Info", "Die Tabelle1 enthält keine lesbaren Daten."]];
      }

      return {
        client: foundCustomer,
        filename: fileName,
        rawData: filteredRows,     // Aufbereitetes "Tabelle1" für die UI-Ansicht auf der Bühne[cite: 1]
        allSheets: allSheetsData   // Komplette Datei mit allen Blättern zur restlosen Datensicherung[cite: 1]
      };

    } catch (err) {
      console.error("Parser-Fehler:", err);
      throw new Error("Fehler beim Einlesen der Excel-Struktur. Bitte prüfen Sie das Dateiformat.");
    }
  }
};
