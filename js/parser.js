/**
 * ============================================================================
 * MODUL: parser.js (Schlauchmanagement-App v0.1.62)
 * ============================================================================
 * Sucht strikt nach dem exakten Wort "Kunde" (mit großem K) und extrahiert
 * den Kundennamen pur aus der rechten Nachbarzelle für die automatische Kundenerstellung.
 * ÄNDERUNG in v0.1.62: 
 * - Der nachgelagerte Zeilenfilter (der Zeilen vor Erkennung von Typ/Länge/Druck weggeworfen hat) wurde entfernt.
 * - Sämtliche Zeilen ab Zeile 0 (inklusive der ersten 20 Zeilen und des Vorspanns) bleiben vollständig erhalten.
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
      let rawRows = [];

      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const jsonSheet = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        
        // Wir greifen primär auf "Tabelle1" zu (oder das erste gefundene Blatt mit Daten)
        if (jsonSheet.length > 0 && (rawRows.length === 0 || sheetName === "Tabelle1")) {
          rawRows = jsonSheet;
        }

        // Exakte Suche nach dem Wort "Kunde" (Großes K) -> Wert in der Zelle rechts daneben ist der Kundenname
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

      // Fallback für den Kundennamen, falls kein Label "Kunde" im Dokument gefunden wurde
      if (!foundCustomer) {
        foundCustomer = fileName.replace(/\.[^/.]+$/, "");
      }

      // ROHDATEN-ÜBERNAHME OHNE ZEILENFILTER:
      // Jede Zeile aus dem Sheet (ab Index 0, inklusive der ersten 20 Zeilen und des Vorspanns) 
      // wird ungefiltert und vollständig für die Rohansicht übernommen.
      let filteredRows = [];

      if (rawRows && Array.isArray(rawRows)) {
        rawRows.forEach(row => {
          if (Array.isArray(row)) {
            filteredRows.push(row);
          }
        });
      }

      if (filteredRows.length === 0) {
        filteredRows = [["Info", "Die Excel-Tabelle enthält keine lesbaren Daten."]];
      }

      return {
        client: foundCustomer,
        filename: fileName,
        rawData: filteredRows,
        sheets: workbook.Sheets // Alle Sheets für direkten Zugriff auf "Tabelle1"
      };

    } catch (err) {
      console.error("Parser-Fehler:", err);
      throw new Error("Fehler beim Einlesen der Excel-Struktur. Bitte prüfen Sie das Dateiformat.");
    }
  }
};
