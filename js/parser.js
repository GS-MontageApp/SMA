/**
 * ============================================================================
 * MODUL: parser.js (Schlauchmanagement-App v0.1.21)
 * ============================================================================
 * Intelligent erfasst der Parser Metadaten (Kunde & Anlage in der rechten Zelle)
 * und filtert relevante Tabellenstrukturen von unwichtigen Randdaten.
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

      // Alle Sheets durchgehen
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const jsonSheet = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        if (jsonSheet.length > 0 && rawRows.length === 0) {
          rawRows = jsonSheet;
        }

        // Intelligente Suche nach Labels "Kunde" und "Anlage" (Wert steht in der Zelle rechts daneben)
        jsonSheet.forEach(row => {
          row.forEach((cellVal, colIdx) => {
            if (cellVal && typeof cellVal === 'string') {
              const text = cellVal.trim().toLowerCase();
              
              // Suche nach Kunde -> Wert rechts daneben
              if (text === 'kunde' || text === 'kundenname') {
                if (row[colIdx + 1] !== undefined && row[colIdx + 1] !== null) {
                  foundCustomer = String(row[colIdx + 1]).trim();
                }
              }

              // Suche nach Anlage -> Wert rechts daneben
              if (text === 'anlage' || text === 'anlagenbezeichnung') {
                if (row[colIdx + 1] !== undefined && row[colIdx + 1] !== null) {
                  foundPlant = String(row[colIdx + 1]).trim();
                }
              }
            }
          });
        });
      });

      // Fallback für Kundenname, falls nicht im Dokument gefunden
      if (!foundCustomer) {
        foundCustomer = fileName.replace(/\.[^/.]+$/, "");
      }

      // Filterung & Strukturierung: Wichtige Tabellenzeilen extrahieren
      let filteredRows = [];
      let headerFound = false;

      rawRows.forEach(row => {
        // Prüfen, ob die Zeile echte Daten oder Header enthält (nicht nur leer)
        const hasContent = row.some(cell => cell !== undefined && cell !== null && String(cell).trim() !== '');
        if (!hasContent) return; // Leerzeilen (unwichtig) verwerfen

        // Prüfen ob es sich um eine Metadaten-Zeile "Kunde / Anlage" handelt (optional in Tabelle)
        const rowString = row.join(' ').toLowerCase();
        if (rowString.includes('kunde') || rowString.includes('anlage')) {
          // Als Metadaten erkannt, aber für die reine Datentabelle optional oder als Info
          filteredRows.push(row);
          return;
        }

        // Erkennung des Tabellenkopfes (Schlauch-Attribute)
        if (!headerFound && (rowString.includes('id') || rowString.includes('typ') || rowString.includes('länge') || rowString.includes('druck'))) {
          headerFound = true;
          filteredRows.push(row);
          return;
        }

        if (headerFound) {
          filteredRows.push(row);
        }
      });

      // Fallback falls keine strukturierte Tabelle erkannt wurde
      if (filteredRows.length === 0) {
        filteredRows = rawRows.length > 0 ? rawRows : [["Info", "Die Excel-Tabelle enthält keine lesbaren Daten."]];
      }

      // Wenn eine Anlage gefunden wurde, hängen wir sie dem Kundennamen an oder sichern sie im Objekt
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
