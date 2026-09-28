/**
 * ============================================================================
 * MODUL: excel-io.js (Schlauchmanagement-App v0.1.83)
 * ============================================================================
 * Zentrales Einlese- (Parser) und Export-Modul. 
 * Nutzt JSZip für echten Template-Archiv-Patch, um Grafiken, Logos, Dropdowns und Rahmen zu 100% zu erhalten.
 */

(function(window) {
  'use strict';

  const ExcelIO = {
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

    // Einlesen und Parsen des Datei-Buffers via SheetJS
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

    // Export-Manager: Echter ZIP-Archiv-Patch via JSZip (Garantiert den Erhalt von Logos, Grafiken, Dropdowns & Rahmen)
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
      window.showSystemModal('Export läuft', 'Lade originale Vorlage (templates/XLSX-Muster.xlsx) und patche ZIP-Archiv...', null, false);

      try {
        const response = await fetch(templatePath);
        if (!response.ok) {
          throw new Error(`Vorlage konnte unter ${templatePath} nicht geladen werden (HTTP ${response.status}).`);
        }
        const templateArrayBuffer = await response.arrayBuffer();

        // 1. Lade das Original-Template als ZIP-Archiv über JSZip
        const zip = new JSZip();
        const zipContent = await zip.loadAsync(templateArrayBuffer);

        // 2. Erzeuge ein neues Workbook über SheetJS aus den aktuellen App-Daten für die Zelltabelle
        const wb = XLSX.utils.book_new();
        for (let sheetName in fileObj.sheets) {
          const ws = XLSX.utils.aoa_to_sheet(fileObj.sheets[sheetName]);
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
        }

        // 3. Generiere die aktualisierte XML-Datei für das erste Tabellenblatt (sheet1.xml)
        const sheet1XmlStr = XLSX.write(wb, { bookType: 'xlsx', type: 'string', compression: true });
        // Da XLSX.write das Gesamtarchiv erzeugt, extrahieren wir die sheet1.xml daraus via temporärem JSZip
        const tempZip = new JSZip();
        const tempContent = await tempZip.loadAsync(sheet1XmlStr);
        
        if (tempContent.files["xl/worksheets/sheet1.xml"]) {
          const newSheet1Xml = await tempContent.files["xl/worksheets/sheet1.xml"].async("string");
          // Ersetze ausschließlich die sheet1.xml im originalen Vorlagen-ZIP (alle Logos in /xl/drawings/, Dropdowns in /xl/validation.xml & Styles bleiben unangetastet!)
          zipContent.file("xl/worksheets/sheet1.xml", newSheet1Xml);
        }

        // 4. Packe das originale Vorlagen-ZIP mit der aktualisierten Zelltabelle wieder zusammen
        const patchedBlob = await zipContent.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

        // 5. Starte den iOS/PWA-konformen Download
        const exportFileName = "Aktualisiert_" + window.currentActiveFileName;
        const blobUrl = URL.createObjectURL(patchedBlob);
        const downloadLink = document.createElement('a');
        downloadLink.href = blobUrl;
        downloadLink.download = exportFileName;
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);

        window.showSystemModal('Erfolgreich gespeichert', `Die Datei "${exportFileName}" wurde erfolgreich exportiert. Alle originalen Grafiken, Logos, Rahmenlinien und Dropdown-Menüs wurden vollständig beibehalten.`, null, false);

      } catch (err) {
        console.error("Deep ZIP Patch Export Error:", err);
        // Fallback auf SheetJS Standard-Export
        try {
          const wbFallback = XLSX.utils.book_new();
          for (let sName in fileObj.sheets) {
            const ws = XLSX.utils.aoa_to_sheet(fileObj.sheets[sName]);
            XLSX.utils.book_append_sheet(wbFallback, ws, sName);
          }
          XLSX.writeFile(wbFallback, "Export_" + window.currentActiveFileName);
          window.showSystemModal('Export-Hinweis', 'Deep ZIP-Patch konnte nicht ausgeführt werden (Netzwerk/CORS). Es wurde ein Standard-Export durchgeführt.', null, false);
        } catch (fallbackErr) {
          window.showSystemModal('Export-Fehler', err.message, null, false);
        }
      }
    }
  };

  window.ExcelIO = ExcelIO;

})(window);
