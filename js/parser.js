/**
 * ============================================================================
 * MODUL: parser.js (Schlauchmanagement-App v0.1.49)
 * ============================================================================
 * Verantwortlich für das Einlesen und Validieren von Excel-Dateien (.xls, .xlsx).
 * Absolut unangetasteter Original-Parser.
 */

window.ExcelParser = {
  validateFile: function(file) {
    if (!file) throw new Error("Keine Datei übergeben.");
    const validExts = ['.xls', '.xlsx'];
    const fileName = file.name.toLowerCase();
    const isValid = validExts.some(ext => fileName.endsWith(ext));
    if (!isValid) {
      throw new Error("Ungültiges Dateiformat. Bitte nur .xls oder .xlsx Dateien importieren.");
    }
    if (file.size > 2 * 1024 * 1024) {
      throw new Error("Die Datei ist zu groß (max. 2 MB erlaubt).");
    }
    return true;
  },

  parseFileBuffer: function(arrayBuffer, fileName) {
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array' });

    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error("Die Excel-Datei enthält keine Arbeitsblätter.");
    }

    let clientName = fileName.replace(/\.[^/.]+$/, "").replace(/^[^-]+-_?/, "").trim();
    if (!clientName) clientName = "Unbekannter Kunde";

    let sheetsData = {};
    workbook.SheetNames.forEach(sheetName => {
      const ws = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      sheetsData[sheetName] = jsonRows;
    });

    return {
      client: clientName,
      filename: fileName,
      sheets: sheetsData
    };
  }
};
