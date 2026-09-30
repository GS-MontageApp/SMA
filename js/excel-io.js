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

      // Export-Modal mit animiertem GIF (templates/logo.gif) anzeigen
      window.showSystemModal(
        'Export läuft', 
        '<div style="display: flex; flex-direction: column; align-items: center; gap: 10px;"><img src="templates/logo.gif" alt="Export läuft..." style="width: 60px; height: 60px;" /><span>Sende Daten an den Export-Server (100% Vorlagenerhalt)...</span></div>', 
        null, 
        false
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

        window.showSystemModal('Erfolgreich gespeichert', `Die Datei "${exportFileName}" wurde erfolgreich exportiert. Autofilter und Dropdown-Bereiche wurden fehlerfrei ausgerichtet.`, null, false);

      } catch (err) {
        console.error("Backend Export Error:", err);
        window.showSystemModal('Export-Fehler', `Verbindung zum Export-Server fehlgeschlagen: ${err.message}`, null, false);
      }
}
