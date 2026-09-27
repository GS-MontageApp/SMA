/**
 * ============================================================================
 * MODUL: parser.js (Schlauchmanagement-App v0.1.45)
 * ============================================================================
 * Verantwortlich für das Einlesen und Validieren von Excel-Dateien (.xls, .xlsx).
 * Extrahiert beim Parsen ausschliesslich "Tabelle1" und filtert jegliche
 * Auswahlseiten oder Zusatzblätter rigoros heraus.
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
    // Maximale Dateigröße 2 MB als Obergrenze gegen Speicherüberlauf
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

    // Kundenname aus Dateiname extrahieren (z.B. "Kunde_Hubtisch4WZB.xlsx" -> "Hubtisch4WZB" oder Dateiname ohne Endung)
    let clientName = fileName.replace(/\.[^/.]+$/, "").replace(/^[^-]+-_?/, "").trim();
    if (!clientName) clientName = "Unbekannter Kunde";

    // EXKLUSIVE FILTERUNG BEIM PARSEN: Nur "Tabelle1" (bzw. primäres Tabellenblatt) wird extrahiert
    let sheetsData = {};
    const sheetNames = workbook.SheetNames;
    
    // Finde das exakte Blatt "Tabelle1" oder das erste Blatt, das keine Auswahlseite ist
    let targetSheetName = sheetNames.find(name => name === "Tabelle1" || name.toLowerCase().replace(/\s+/g, '') === "tabelle1");
    
    if (!targetSheetName) {
      targetSheetName = sheetNames.find(name => {
        const lower = name.toLowerCase();
        return !lower.includes('auswahl') && !lower.includes('choice') && !lower.includes('menu');
      });
    }

    if (!targetSheetName) {
      targetSheetName = sheetNames[0];
    }

    // Hole ausschließlich die Daten dieses EINEN Blatts und speichere es unter "Tabelle1" ab
    const ws = workbook.Sheets[targetSheetName];
    const jsonRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    
    sheetsData["Tabelle1"] = jsonRows;

    return {
      client: clientName,
      filename: fileName,
      sheets: sheetsData
    };
  }
};
