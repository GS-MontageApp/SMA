/**
 * ============================================================================
 * MODUL: excel-io.js (Schlauchmanagement-App v0.1.110)
 * ============================================================================
 * Zentrales Einlese- (Parser) und Export-Modul.
 */

(function(window) {
  'use strict';

  const BACKEND_BASE_URL = "https://sma-63h4.onrender.com";
  const BACKEND_EXPORT_URL = `${BACKEND_BASE_URL}/api/export`;

  const ExcelIO = {
    pingBackend: function() {
      fetch(BACKEND_BASE_URL + "/")
        .then(res => res.json())
        .then(data => console.log("Backend Status:", data))
        .catch(err => console.log("Backend pinging in background...", err));
    },

    validateFile: function(file) {
      if (!file) throw new Error("Keine Datei ausgewählt.");
      const validExtensions = ['.xls', '.xlsx'];
      const fileNameLower = file.name.toLowerCase();
      const isValidExt = validExtensions.some(ext => fileNameLower.endsWith(ext));
      
      if (!isValidExt) {
        throw new Error("Ungültiges Dateiformat. Bitte wählen Sie eine .xls oder .xlsx Datei aus.");
      }
      
      const maxSize = 600 * 1024;
      const minSize = 10 * 1024;
      
      if (file.size > maxSize) {
        throw new Error(`Die Datei ist zu groß (${Math.round(file.size / 1024)} KB). Maximal zulässig sind 600 KB.`);
      }
      if (file.size < minSize) {
        throw new Error("Die Datei ist zu klein oder beschädigt.");
      }
      return true;
    },

    parseFileBuffer: function(arrayBuffer, fileName) {
      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      
      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        throw new Error("Die Excel-Datei enthält keine Tabellenblätter.");
      }

      let sheetsData = {};
      workbook.SheetNames.forEach(sheetName => {
        const ws = workbook.Sheets[sheetName];
        const jsonRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        sheetsData[sheetName] = jsonRows;
      });

      let detectedClient = "Unbekannter Kunde";
      const primarySheet = sheetsData[workbook.SheetNames[0]];
      
      if (primarySheet && primarySheet.length > 1) {
        const row1 = primarySheet[1];
        if (Array.isArray(row1)) {
          for (let i = 0; i < row1.length; i++) {
            const val = String(row1[i]).trim();
            if (val.toLowerCase() === "kunde" && row1[i + 1]) {
              detectedClient = String(row1[i + 1]).trim();
              break;
            }
          }
          if (detectedClient === "Unbekannter Kunde") {
            for (let cell of row1) {
              if (cell && String(cell).trim() !== "") {
                detectedClient = String(cell).trim();
                break;
              }
            }
          }
        }
      }

      return {
        client: detectedClient,
        filename: fileName,
        sheets: sheetsData
      };
    },

    handleExcelImport: function(event) {
      const file = event.target.files[0];
      if (!file) return;

      try {
        ExcelIO.validateFile(file);
      } catch (err) {
        window.showSystemModal('Validierungsfehler', err.message, null, false);
        event.target.value = '';
        return;
      }

      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const parsed = ExcelIO.parseFileBuffer(e.target.result, file.name);

          if (window.AppData.hasFile(parsed.client, parsed.filename)) {
            window.showSystemModal(
              'Datei bereits vorhanden',
              `Die Datei "${parsed.filename}" existiert bereits für ${parsed.client}. Möchten Sie die vorhandene Version überschreiben?`,
              function() {
                window.AppData.addFileToClient(parsed.client, parsed.filename, parsed);
                if (window.UIPool && typeof window.UIPool.renderDateipool === 'function') {
                  window.UIPool.renderDateipool();
                }
                window.showSystemModal('Erfolgreich', `Die Datei "${parsed.filename}" wurde für ${parsed.client} aktualisiert.`, null, false);
                window.switchApp('dateipool', 'Dateipool');
              }
            );
          } else {
            window.AppData.addFileToClient(parsed.client, parsed.filename, parsed);
            if (window.UIPool && typeof window.UIPool.renderDateipool === 'function') {
              window.UIPool.renderDateipool();
            }
            window.showSystemModal('Erfolgreich', `Erfolgreich importiert!\nKunde: ${parsed.client}\nDatei: ${parsed.filename}`, null, false);
            window.switchApp('dateipool', 'Dateipool');
          }

        } catch (err) {
          console.error(err);
          window.showSystemModal('Fehler', err.message, null, false);
        } finally {
          event.target.value = '';
        }
      };
      reader.readAsArrayBuffer(file);
    },

    saveCurrentStageFile: async function() {
      if (!window.currentActiveCustomer || !window.currentActiveFileName) {
        window.showSystemModal('Hinweis', 'Es ist keine aktive Datei zum Speichern geöffnet.', null, false);
        return;
      }

      const clients = window.AppData.getClients();
      const fileList = clients[window.currentActiveCustomer] || [];
      const fileObj = fileList.find(f => f && f.name === window.currentActiveFileName);

      if (!fileObj || !fileObj.sheets) {
        window.showSystemModal('Fehler', 'Die aktuelle Datei konnte im Speicher nicht gefunden werden.', null, false);
        return;
      }

      // Export-Modal im höhenoptimierten Kompaktmodus (isCompact = true) mit kompaktem Logo (z.B. 120px)
      window.showSystemModal(
        'Export läuft', 
        '<img src="templates/logo.gif" alt="Export läuft..." style="width: 120px; height: 120px; object-fit: contain; display: block; margin: 0 auto; padding: 0;" />', 
        null, 
        false, 
        true
      );

      try {
        let updates = [];
        for (let sheetName in fileObj.sheets) {
          const rows = fileObj.sheets[sheetName];
          rows.forEach((row, rIdx) => {
            if (Array.isArray(row)) {
              row.forEach((cellVal, cIdx) => {
                if (cellVal !== undefined && cellVal !== null && cellVal !== "") {
                  updates.push({
                    sheet_name: sheetName,
                    row: rIdx + 1,
                    col: cIdx + 1,
                    value: cellVal
                  });
                }
              });
            }
          });
        }

        const payload = {
          filename: window.currentActiveFileName,
          updates: updates
        };

        const response = await fetch(BACKEND_EXPORT_URL, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(payload)
        });

        if (!response.ok) {
          const errData = await response.json().catch(() => ({}));
          throw new Error(errData.detail || `Server-Fehler (HTTP ${response.status})`);
        }

        const blob = await response.blob();
        const exportFileName = "Aktualisiert_" + window.currentActiveFileName;
        
        const blobUrl = URL.createObjectURL(blob);
        const downloadLink = document.createElement('a');
        downloadLink.href = blobUrl;
        downloadLink.download = exportFileName;
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);

        // Das Modal nach erfolgreichem Download automatisch wieder schließen
        const modal = document.getElementById('system_generic_modal');
        if (modal) modal.classList.add('hidden');

      } catch (err) {
        console.error("Backend Export Error:", err);
        window.showSystemModal('Export-Fehler', `Verbindung zum Export-Server fehlgeschlagen: ${err.message}`, null, false);
      }
    }
  };

  window.ExcelIO = ExcelIO;

})(window);
