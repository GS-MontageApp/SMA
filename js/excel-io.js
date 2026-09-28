/**
 * ============================================================================
 * MODUL: excel-io.js (Schlauchmanagement-App v0.1.81)
 * ============================================================================
 * Zentrales Einlese- (Parser) und Export-Modul (Template-Filling basierend auf templates/XLSX-Muster.xlsx).
 */

window.ExcelIO = {
  // Validierung der hochgeladenen Datei (Größe & Format)
  validateFile: function(file) {
    if (!file) throw new Error("Keine Datei ausgewählt.");
    const validExtensions = ['.xls', '.xlsx'];
    const fileNameLower = file.name.toLowerCase();
    const isValidExt = validExtensions.some(ext => fileNameLower.endsWith(ext));
    
    if (!isValidExt) {
      throw new Error("Ungültiges Dateiformat. Bitte wählen Sie eine .xls oder .xlsx Datei aus.");
    }
    
    const maxSize = 600 * 1024; // 600 KB OOM-Schutz für mobile Browser
    const minSize = 10 * 1024;  // 10 KB Mindestgröße
    
    if (file.size > maxSize) {
      throw new Error(`Die Datei ist zu groß (${Math.round(file.size / 1024)} KB). Maximal zulässig sind 600 KB.`);
    }
    if (file.size < minSize) {
      throw new Error("Die Datei ist zu klein oder beschädigt.");
    }
    return true;
  },

  // Einlesen und Parsen des Datei-Buffers
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

  // Zentraler Import-Handler für den Dateiupload
  handleExcelImport: function(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
      window.ExcelIO.validateFile(file);
    } catch (err) {
      window.showSystemModal('Validierungsfehler', err.message, null, false);
      event.target.value = '';
      return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const parsed = window.ExcelIO.parseFileBuffer(e.target.result, file.name);

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

  // Export-Manager: Lädt templates/XLSX-Muster.xlsx, füllt Daten ein und bietet Download an
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

    const templatePath = 'templates/XLSX-Muster.xlsx';
    window.showSystemModal('Export läuft', 'Lade originale Vorlage (templates/XLSX-Muster.xlsx) und bereite Export vor...', null, false);

    try {
      const response = await fetch(templatePath);
      if (!response.ok) {
        throw new Error(`Vorlage konnte unter ${templatePath} nicht geladen werden (HTTP ${response.status}).`);
      }
      const arrayBuffer = await response.arrayBuffer();
      
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true, cellFormulas: true });

      for (let sheetName in fileObj.sheets) {
        if (workbook.Sheets[sheetName]) {
          const targetSheet = workbook.Sheets[sheetName];
          const sourceRows = fileObj.sheets[sheetName];

          sourceRows.forEach((row, rIdx) => {
            if (Array.isArray(row)) {
              row.forEach((cellVal, cIdx) => {
                const cellAddress = XLSX.utils.encode_cell({ r: rIdx, c: cIdx });
                if (cellVal !== undefined && cellVal !== null && cellVal !== "") {
                  if (!targetSheet[cellAddress]) {
                    targetSheet[cellAddress] = { t: 's', v: cellVal };
                  } else {
                    targetSheet[cellAddress].v = cellVal;
                    targetSheet[cellAddress].t = typeof cellVal === 'number' ? 'n' : 's';
                  }
                }
              });
            }
          });
        }
      }

      const exportFileName = "Aktualisiert_" + window.currentActiveFileName;
      XLSX.writeFile(workbook, exportFileName);

      window.showSystemModal('Erfolgreich gespeichert', `Die Datei "${exportFileName}" wurde basierend auf der Originalvorlage erfolgreich exportiert und heruntergeladen. Alle Layouts, Rahmen und Dropdowns wurden beibehalten.`, null, false);

    } catch (err) {
      console.error("Template Export Error:", err);
      try {
        const wb = XLSX.utils.book_new();
        for (let sName in fileObj.sheets) {
          const ws = XLSX.utils.aoa_to_sheet(fileObj.sheets[sName]);
          XLSX.utils.book_append_sheet(wb, ws, sName);
        }
        XLSX.writeFile(wb, "Export_" + window.currentActiveFileName);
        window.showSystemModal('Export-Hinweis', 'Vorlage konnte nicht per fetch geladen werden (CORS/Lokal-Modus). Es wurde ein Standard-Export der aktuellen Daten durchgeführt.', null, false);
      } catch (fallbackErr) {
        window.showSystemModal('Export-Fehler', err.message, null, false);
      }
    }
  }
};
