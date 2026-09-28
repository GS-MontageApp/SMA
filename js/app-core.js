/**
 * ============================================================================
 * MODUL: parser.js (Schlauchmanagement-App v0.1.70)
 * ============================================================================
 * Kapselt das sichere Einlesen und Validieren von Excel-Dateien (.xls/.xlsx).
 */

window.ExcelParser = {
  validateFile: function(file) {
    if (!file) throw new Error("Keine Datei ausgewählt.");
    const validExtensions = ['xls', 'xlsx'];
    const extension = file.name.split('.').pop().toLowerCase();
    if (!validExtensions.includes(extension)) {
      throw new Error("Ungültiges Dateiformat. Bitte verwenden Sie .xls oder .xlsx.");
    }
    if (file.size > 600 * 1024) {
      throw new Error(`Die Datei ist zu groß (${(file.size/1024).toFixed(1)} KB). Maximal erlaubt sind 600 KB.`);
    }
  },

  parseFileBuffer: function(buffer, filename) {
    try {
      const workbook = XLSX.read(buffer, { type: 'array' });
      let parsedSheets = {};
      
      workbook.SheetNames.forEach(sheetName => {
        parsedSheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
      });

      let clientName = filename.split(/[_.]/)[0] || "Unbekannter Kunde";
      
      return { client: clientName, filename: filename, sheets: parsedSheets };
    } catch (err) {
      console.error("Fehler im Parser:", err);
      throw new Error("Fehler beim Lesen der Excel-Struktur. Ist die Datei beschädigt?");
    }
  }
};
