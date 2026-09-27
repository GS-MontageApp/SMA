/**
 * ============================================================================
 * UI-RENDER-LOGIK: Exklusive Ansicht von "Tabelle1" auf der Bühne
 * ============================================================================
 * Diese Funktion greift auf die gespeicherten Dateidaten zu, sucht gezielt 
 * nach dem Arbeitsblatt "Tabelle1" und rendert ausschließlich dieses auf der Bühne,
 * während alle anderen Blätter im Hintergrund (im Dateipool) erhalten bleiben.
 */

window.UIPool = window.UIPool || {};

// Beispiel für die Funktion, die eine Datei auf der Bühne öffnet und darstellt:
window.UIPool.renderFileOnStage = function(fileData) {
  const container = document.getElementById('buehne_table_container');
  if (!container) return;

  try {
    // 1. Prüfen, ob die gespeicherten Daten mehrere Blätter beinhalten (allSheets) 
    // oder ob wir auf die Rohdaten von Tabelle1 zugreifen.
    let targetRows = [];

    if (fileData.allSheets && typeof fileData.allSheets === 'object') {
      // Suche im mehrblättrigen Objekt exakt nach "Tabelle1" (Groß-/Kleinschreibung tolerant oder exakt)
      const exactSheetKey = Object.keys(fileData.allSheets).find(
        name => name.toLowerCase() === 'tabelle1'
      );

      if (exactSheetKey) {
        targetRows = fileData.allSheets[exactSheetKey];
      } else {
        // Fallback, falls "Tabelle1" nicht exakt existiert: Das erste verfügbare Blatt nehmen
        const firstKey = Object.keys(fileData.allSheets)[0];
        targetRows = firstKey ? fileData.allSheets[firstKey] : (fileData.rawData || []);
      }
    } else {
      // Fallback auf Standard-Rohdaten
      targetRows = fileData.rawData || [];
    }

    // 2. Grafische Aufbereitung und Generierung der Tabellenansicht für die Bühne
    let html = '<div class="overflow-x-auto h-full"><table class="w-full border-collapse text-xs text-slate-700 bg-white">';
    
    targetRows.forEach((row, rowIndex) => {
      html += '<tr class="border-b border-slate-100 hover:bg-slate-50 transition-colors">';
      
      // Spaltenzellen generieren
      row.forEach(cell => {
        const cellValue = (cell !== undefined && cell !== null) ? String(cell) : '';
        
        if (rowIndex === 0) {
          // Tabellenkopf (Header)
          html += `<th class="bg-slate-100 border border-slate-200 px-3 py-2 font-bold text-slate-600 text-left sticky top-0 z-10">${escapeHtml(cellValue)}</th>`;
        } else {
          // Reguläre Tabellenzellen
          html += `<td class="border border-slate-200 px-3 py-2 text-slate-700">${escapeHtml(cellValue)}</td>`;
        }
      });
      
      html += '</tr>';
    });

    html += '</table></div>';
    container.innerHTML = html;

  } catch (err) {
    console.error("Fehler beim Rendern von Tabelle1 auf der Bühne:", err);
    container.innerHTML = '<div class="p-4 text-red-600 text-xs font-semibold">Fehler beim Laden von Tabelle1.</div>';
  }
};

// Hilfsfunktion zur HTML-Sicherheit
function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace('/&#039;/g', "&#039;");
}
