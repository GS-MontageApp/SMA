/**
 * ============================================================================
 * MODUL: parser.js (Schlauchmanagement-App v0.1.69)
 * ============================================================================
 * Kapselt das Einlesen der Excel-Dateien über SheetJS.
 * FIX in v0.1.69: defval: "" hinzugefügt, um leere Spalten zwingend aufzufüllen.
 */

window.ExcelParser = {
  validateFile: function(file) {
    if (!file) throw new Error("Keine Datei ausgewählt.");
    const validExtensions = ['xls', 'xlsx'];
    const extension = file.name.split('.').pop().toLowerCase();
    if (!validExtensions.includes(extension)) {
      throw new Error("Ungültiges Dateiformat. Bitte verwenden Sie .xls oder .xlsx.");
    }
    // Limitierung auf 600 KB zum Schutz des Arbeitsspeichers mobiler Browser
    if (file.size > 600 * 1024) {
      throw new Error(`Die Datei ist zu groß (${(file.size/1024).toFixed(1)} KB). Maximal erlaubt sind 600 KB.`);
    }
  },

  parseFileBuffer: function(buffer, filename) {
    try {
      const workbook = XLSX.read(buffer, { type: 'array' });
      let parsedSheets = {};
      
      workbook.SheetNames.forEach(sheetName => {
        // DER ENTSCHEIDENDE FIX FÜR DIE SPALTEN-TREUE:
        // header: 1 gibt ein 2D-Array zurück.
        // defval: "" zwingt den Parser, auch komplett leere oder verbundene Zellen 
        // mit einem leeren String aufzufüllen. So ist Index 0 immer Spalte A und Index 17 immer Spalte R.
        parsedSheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { 
          header: 1, 
          defval: "" 
        });
      });

      // Simple Zuweisung des Kunden-Namens basierend auf Dateinamen oder Vorgabe
      let clientName = filename.split(/[_.]/)[0] || "Unbekannter Kunde";
      
      return { client: clientName, filename: filename, sheets: parsedSheets };
    } catch (err) {
      console.error("Fehler im Parser:", err);
      throw new Error("Fehler beim Lesen der Excel-Struktur. Ist die Datei beschädigt?");
    }
  }
};
