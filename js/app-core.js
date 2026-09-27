/**
 * ============================================================================
 * MODUL: app-core.js (Schlauchmanagement-App v0.1.36)
 * ============================================================================
 * Kapselt die zentrale App-Logik, Datenverwaltung, Session-Persistenz und Routing.
 * Die grafische Ansicht auf der Bühne filtert exklusiv auf "Tabelle1" (Auswahlseite wird ignoriert).
 */

window.currentActiveCustomer = null;
window.currentActiveFileName = null;
window.openedFilesStack = [];

window.AppData = {
  getClients: function() {
    const stored = localStorage.getItem('sma_clients_data');
    if (stored) {
      try {
        const parsed = JSON.parse(stored);
        let normalized = {};
        for (let client in parsed) {
          if (Array.isArray(parsed[client])) {
            normalized[client] = parsed[client];
          } else if (parsed[client]) {
            normalized[client] = [parsed[client]];
          } else {
            normalized[client] = [];
          }
        }
        return normalized;
      } catch(e) { console.error(e); }
    }
    const nowStr = new Date().toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    return {
      "Beispielkunde 1": [
        { 
          name: "Beispielschlauchliste.xlsx", 
          timestamp: nowStr, 
          sheets: { 
            "Tabelle1": [["Kunde", "Beispielkunde 1"], ["Schlauch-ID", "Typ", "Länge", "Druck"], ["SH-001", "2SN", "1500", "210"], ["SH-002", "4SH", "2000", "420"]],
            "Auswahlseite": [["Auswahl", "Wert"], ["Option", "1"]]
          } 
        }
      ]
    };
  },
  saveClients: function(data) {
    localStorage.setItem('sma_clients_data', JSON.stringify(data));
  },
  hasFile: function(clientName, fileName) {
    let clients = this.getClients();
    if (!clients[clientName] || !Array.isArray(clients[clientName])) return false;
    return clients[clientName].some(f => f && f.name === fileName);
  },
  addFileToClient: function(clientName, fileName, sheetsData) {
    let clients = this.getClients();
    if (!clients[clientName] || !Array.isArray(clients[clientName])) {
      clients[clientName] = [];
    }
    const nowStr = new Date().toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    
    const existingIdx = clients[clientName].findIndex(f => f && f.name === fileName);
    if (existingIdx >= 0) {
      clients[clientName][existingIdx].timestamp = nowStr;
      if (sheetsData) clients[clientName][existingIdx].sheets = sheetsData;
    } else {
      clients[clientName].push({ name: fileName, timestamp: nowStr, sheets: sheetsData });
    }
    this.saveClients(clients);
  },
  removeFileFromClient: function(clientName, fileName) {
    let clients = this.getClients();
    if (clients[clientName] && Array.isArray(clients[clientName])) {
      clients[clientName] = clients[clientName].filter(f => f && f.name !== fileName);
      this.saveClients(clients);
    }
  },
  removeClient: function(clientName) {
    let clients = this.getClients();
    delete clients[clientName];
    this.saveClients(clients);
  }
};

window.saveSessionState = function() {
  const sessionState = {
    activeCustomer: window.currentActiveCustomer,
    activeFile: window.currentActiveFileName,
    stack: window.openedFilesStack
  };
  localStorage.setItem('sma_session_state', JSON.stringify(sessionState));
};

window.loadSessionState = function() {
  const stored = localStorage.getItem('sma_session_state');
  if (!stored) return false;
  try {
    const parsed = JSON.parse(stored);
    if (parsed && Array.isArray(parsed.stack) && parsed.stack.length > 0) {
      window.openedFilesStack = parsed.stack;
      if (parsed.activeFile && parsed.activeCustomer) {
        window.openFileOnStage(parsed.activeCustomer, parsed.activeFile, false);
        return true;
      }
    }
  } catch (e) {
    console.error("Fehler beim Laden des Session-States:", e);
  }
  return false;
};

window.setTheme = function(themeName) {
  const body = document.body;
  if (!body) return;
  const btnStandard = document.getElementById('theme_btn_standard');
  const btnGs = document.getElementById('theme_btn_gs');
  if (themeName === 'gs') {
    body.classList.add('bg-slate-900', 'text-slate-100');
    body.classList.remove('bg-slate-100', 'text-slate-800');
    if (btnGs) btnGs.className = 'p-2.5 rounded-xl text-xs font-bold border-2 border-yellow-400 bg-slate-900 text-yellow-400 shadow-md transition-all text-center';
    if (btnStandard) btnStandard.className = 'p-2.5 rounded-xl text-xs font-bold border border-slate-600 bg-slate-800 text-slate-300 transition-all text-center';
  } else {
    body.classList.remove('bg-slate-900', 'text-slate-100');
    body.classList.add('bg-slate-100', 'text-slate-800');
    if (btnStandard) btnStandard.className = 'p-2.5 rounded-xl text-xs font-bold border-2 border-indigo-600 bg-white text-indigo-700 shadow-md transition-all text-center';
    if (btnGs) btnGs.className = 'p-2.5 rounded-xl text-xs font-bold border border-slate-300 bg-slate-100 text-slate-700 transition-all text-center';
  }
  localStorage.setItem('schlauchmanagement_theme', themeName);
};

window.switchApp = function(viewId, titleText) {
  document.querySelectorAll('.app-view').forEach(el => el.classList.add('hidden'));
  const target = document.getElementById('view-' + viewId);
  if (target) target.classList.remove('hidden');

  const titleEl = document.getElementById('header-title');
  if (titleEl) titleEl.textContent = titleText;

  if (viewId === 'dateipool') {
    if (window.UIPool && typeof window.UIPool.renderDateipool === 'function') {
      window.UIPool.renderDateipool();
    }
  }
};

window.updateFooterOpenFiles = function() {
  const container = document.getElementById('file_tabs_container');
  if (!container) return;
  container.innerHTML = '';

  if (window.openedFilesStack.length === 0) {
    container.innerHTML = '<span class="text-xs text-slate-400 px-2">Keine Datei geöffnet</span>';
    return;
  }

  window.openedFilesStack.forEach(item => {
    const isCurrent = (item.clientName === window.currentActiveCustomer && item.fileName === window.currentActiveFileName);
    const badge = document.createElement('div');
    badge.className = (isCurrent ? 'bg-emerald-600 text-white shadow-md ring-2 ring-emerald-300 ' : 'bg-slate-200 hover:bg-slate-300 text-slate-700 ') + 
                      'px-3.5 py-2 rounded-xl text-xs font-bold flex items-center gap-2 whitespace-nowrap cursor-pointer transition-colors shrink-0';
    badge.title = `${item.clientName} / ${item.fileName}`;
    badge.innerHTML = `<span>${item.fileName}</span>`;
    badge.onclick = () => {
      window.openFileOnStage(item.clientName, item.fileName, false);
    };
    container.appendChild(badge);
  });
};

window.openCustomerFiles = function(customerName) {
  window.currentActiveCustomer = customerName;
  document.querySelectorAll('.app-view').forEach(el => el.classList.add('hidden'));
  const target = document.getElementById('view-customer-files');
  if (target) target.classList.remove('hidden');

  const titleEl = document.getElementById('header-title');
  if (titleEl) titleEl.textContent = customerName;

  if (window.UIPool && typeof window.UIPool.renderCustomerFilesList === 'function') {
    window.UIPool.renderCustomerFilesList(window.currentActiveCustomer);
  }
};

window.openFileOnStage = function(clientName, fileName, pushToStack = true) {
  const clients = window.AppData.getClients();
  const fileList = clients[clientName] || [];
  const fileObj = fileList.find(f => f && f.name === fileName);

  if (!fileObj) {
    window.showSystemModal('Fehler', `Die angeforderte Datei "${fileName}" für Kunde "${clientName}" konnte nicht gefunden werden.`, null, false);
    return;
  }

  window.currentActiveCustomer = clientName;
  window.currentActiveFileName = fileName;

  if (pushToStack) {
    window.openedFilesStack = window.openedFilesStack.filter(item => !(item.clientName === clientName && item.fileName === fileName));
    window.openedFilesStack.unshift({ clientName, fileName });
  }

  window.updateFooterOpenFiles();
  window.saveSessionState();

  // EXKLUSIVE UI-FILTERUNG: Suche gezielt nach Tabelle1, ignoriere Auswahlseite komplett
  let targetRows = null;
  if (fileObj.sheets) {
    const sheetKeys = Object.keys(fileObj.sheets);
    // 1. Priorität: Exakter Treffer auf "Tabelle1" oder "Tabelle 1"
    let tab1Key = sheetKeys.find(k => k.toLowerCase() === 'tabelle1' || k.toLowerCase() === 'tabelle 1');
    
    if (tab1Key) {
      targetRows = fileObj.sheets[tab1Key];
    } else {
      // 2. Fallback: Nimm das erste Blatt, das NICHT "auswahlseite" oder "auswahl" heißt
      const validKeys = sheetKeys.filter(k => {
        const lower = k.toLowerCase();
        return !lower.includes('auswahl') && !lower.includes('choice');
      });
      if (validKeys.length > 0) {
        targetRows = fileObj.sheets[validKeys[0]];
      } else {
        targetRows = fileObj.sheets[sheetKeys[0]];
      }
    }
  } else if (fileObj.rawData) {
    targetRows = fileObj.rawData;
  }

  let rawData = [];
  let headerFound = false;
  if (targetRows && Array.isArray(targetRows)) {
    targetRows.forEach(row => {
      const hasContent = row.some(cell => cell !== undefined && cell !== null && String(cell).trim() !== '');
      if (!hasContent) return;

      const rowString = row.join(' ').toLowerCase();
      if (rowString.includes('kunde') || rowString.includes('anlage')) {
        rawData.push(row);
        return;
      }

      if (!headerFound && (rowString.includes('id') || rowString.includes('typ') || rowString.includes('länge') || rowString.includes('druck') || rowString.includes('kennz'))) {
        headerFound = true;
        rawData.push(row);
        return;
      }

      if (headerFound) {
        rawData.push(row);
      }
    });
    if (rawData.length === 0) {
      rawData = targetRows;
    }
  } else {
    rawData = [["Info", "Keine Tabellendaten in Tabelle1 verfügbar"]];
  }

  document.querySelectorAll('.app-view').forEach(el => el.classList.add('hidden'));
  const stageView = document.getElementById('view-buehne');
  if (stageView) stageView.classList.remove('hidden');

  const titleEl = document.getElementById('header-title');
  if (titleEl) titleEl.textContent = `${clientName} / ${fileName}`;

  const container = document.getElementById('buehne_table_container');
  if (!container) return;
  container.innerHTML = '';

  const wrapper = document.createElement('div');
  wrapper.className = 'overflow-x-auto h-[calc(100vh-170px)] bg-white shadow-none w-full';

  const table = document.createElement('table');
  table.className = 'w-full text-left border-collapse text-xs sm:text-sm text-slate-700';

  const thead = document.createElement('thead');
  thead.className = 'sticky top-0 bg-slate-100 text-slate-800 font-bold border-b border-slate-300 shadow-xs z-10';
  
  const tbody = document.createElement('tbody');
  tbody.className = 'divide-y divide-slate-100';

  rawData.forEach((row, rowIndex) => {
    const tr = document.createElement('tr');
    tr.className = rowIndex === 0 ? 'bg-slate-100' : 'hover:bg-slate-50/80 transition-colors';

    row.forEach((cellVal) => {
      const cell = rowIndex === 0 ? document.createElement('th') : document.createElement('td');
      cell.className = 'px-4 py-3 whitespace-nowrap ' + (rowIndex === 0 ? 'font-bold text-slate-700' : 'text-slate-600');
      cell.textContent = cellVal !== undefined && cellVal !== null ? cellVal : '';
      tr.appendChild(cell);
    });

    if (rowIndex === 0) {
      thead.appendChild(tr);
    } else {
      tbody.appendChild(tr);
    }
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  wrapper.appendChild(table);
  container.appendChild(wrapper);
};

window.closeFileOnStage = function() {
  if (window.currentActiveCustomer && window.currentActiveFileName) {
    const clients = window.AppData.getClients();
    const fileList = clients[window.currentActiveCustomer] || [];
    const fileObj = fileList.find(f => f && f.name === window.currentActiveFileName);
    if (fileObj) {
      window.AppData.addFileToClient(window.currentActiveCustomer, window.currentActiveFileName, fileObj.sheets);
    }
    window.openedFilesStack = window.openedFilesStack.filter(item => !(item.clientName === window.currentActiveCustomer && item.fileName === window.currentActiveFileName));
  }

  window.saveSessionState();

  if (window.openedFilesStack.length > 0) {
    const nextItem = window.openedFilesStack[0];
    window.openFileOnStage(nextItem.clientName, nextItem.fileName, false);
  } else {
    window.updateFooterOpenFiles();
    if (window.currentActiveCustomer) {
      window.openCustomerFiles(window.currentActiveCustomer);
    } else {
      window.switchApp('dateipool', 'Dateipool');
    }
  }
};

window.confirmDeleteFile = function(clientName, fileName) {
  window.showSystemModal(
    'Datei löschen',
    `Möchten Sie die Datei "${fileName}" wirklich unwiderruflich löschen?`,
    function() {
      window.AppData.removeFileFromClient(clientName, fileName);
      window.openedFilesStack = window.openedFilesStack.filter(item => !(item.clientName === clientName && item.fileName === fileName));
      window.updateFooterOpenFiles();
      window.saveSessionState();
      if (window.UIPool && typeof window.UIPool.renderCustomerFilesList === 'function') {
        window.UIPool.renderCustomerFilesList(window.currentActiveCustomer);
      }
      window.showSystemModal('Gelöscht', `Die Datei "${fileName}" wurde erfolgreich entfernt.`, null, false);
    }
  );
};

window.confirmDeleteClient = function(clientName, fileCount) {
  window.showSystemModal(
    'Kunden löschen',
    `Möchten Sie wirklich den Kunden "${clientName}" und die zugehörigen ${fileCount} ${fileCount === 1 ? 'Datei' : 'Dateien'} löschen?`,
    function() {
      window.AppData.removeClient(clientName);
      window.openedFilesStack = window.openedFilesStack.filter(item => item.clientName !== clientName);
      window.updateFooterOpenFiles();
      window.saveSessionState();
      if (window.UIPool && typeof window.UIPool.renderDateipool === 'function') {
        window.UIPool.renderDateipool();
      }
      window.showSystemModal('Gelöscht', `Der Kunde "${clientName}" wurde samt aller Dateien entfernt.`, null, false);
    }
  );
};

window.showSystemModal = function(title, message, onConfirm, showCancel = true) {
  const modal = document.getElementById('system_generic_modal');
  const titleEl = document.getElementById('system_modal_title');
  const msgEl = document.getElementById('system_modal_message');
  const confirmBtn = document.getElementById('system_modal_confirm_btn');
  const cancelBtn = document.getElementById('system_modal_cancel_btn');

  if (!modal) return;

  titleEl.textContent = title;
  msgEl.textContent = message;

  if (showCancel) {
    cancelBtn.classList.remove('hidden');
  } else {
    cancelBtn.classList.add('hidden');
  }

  const newConfirmBtn = confirmBtn.cloneNode(true);
  confirmBtn.parentNode.replaceChild(newConfirmBtn, confirmBtn);

  newConfirmBtn.onclick = function() {
    modal.classList.add('hidden');
    if (typeof onConfirm === 'function') onConfirm();
  };

  cancelBtn.onclick = function() {
    modal.classList.add('hidden');
  };

  modal.classList.remove('hidden');
};

let pendingImportData = null;

window.handleExcelImport = function(event) {
  const file = event.target.files[0];
  if (!file) return;

  try {
    window.ExcelParser.validateFile(file);
  } catch (err) {
    window.showSystemModal('Validierungsfehler', err.message, null, false);
    event.target.value = '';
    return;
  }

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const parsed = window.ExcelParser.parseFileBuffer(e.target.result, file.name);
      pendingImportData = parsed;

      if (window.AppData.hasFile(parsed.client, parsed.filename)) {
        window.showSystemModal(
          'Datei bereits vorhanden',
          `Die Datei "${parsed.filename}" existiert bereits für ${parsed.client}. Möchten Sie die vorhandene Version überschreiben?`,
          function() {
            window.AppData.addFileToClient(pendingImportData.client, pendingImportData.filename, pendingImportData.sheets);
            if (window.UIPool && typeof window.UIPool.renderDateipool === 'function') {
              window.UIPool.renderDateipool();
            }
            window.showSystemModal('Erfolgreich', `Die Datei "${pendingImportData.filename}" wurde für ${parsed.client} aktualisiert.`, null, false);
            window.switchApp('dateipool', 'Dateipool');
          }
        );
      } else {
        window.AppData.addFileToClient(parsed.client, parsed.filename, parsed.sheets);
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
};

window.openTopMenu = function() {
  const menu = document.getElementById('top_menu_modal');
  if (menu) menu.classList.remove('hidden');
};

window.closeTopMenu = function() {
  const menu = document.getElementById('top_menu_modal');
  if (menu) menu.classList.add('hidden');
};

window.openCacheClearModal = function() {
  window.closeTopMenu();
  const modal = document.getElementById('cache_clear_modal');
  if (modal) {
    document.querySelectorAll('.cache-checkbox').forEach(cb => cb.checked = false);
    modal.classList.remove('hidden');
  }
};

window.closeCacheClearModal = function() {
  const modal = document.getElementById('cache_clear_modal');
  if (modal) modal.classList.add('hidden');
};

window.executeGranularCacheClear = function() {
  const clearSession = document.getElementById('chk_session')?.checked;
  const clearClients = document.getElementById('chk_clients')?.checked;
  const clearAuth = document.getElementById('chk_auth')?.checked;
  const clearTheme = document.getElementById('chk_theme')?.checked;

  if (!clearSession && !clearClients && !clearAuth && !clearTheme) {
    window.showSystemModal('Hinweis', 'Es wurde keine Auswahl getroffen.', null, false);
    window.closeCacheClearModal();
    return;
  }

  if (clearSession) {
    localStorage.removeItem('sma_session_state');
  }
  if (clearClients) {
    localStorage.removeItem('sma_clients_data');
  }
  if (clearAuth) {
    localStorage.removeItem('sma_current_user');
    if (window.AuthManager && typeof window.AuthManager.forceReset === 'function') {
      window.AuthManager.forceReset();
    }
  }
  if (clearTheme) {
    localStorage.removeItem('schlauchmanagement_theme');
  }

  window.closeCacheClearModal();
  window.showSystemModal('Erfolgreich', 'Die ausgewählten Cache-Bereiche wurden bereinigt. Die App wird neu geladen.', function() {
    window.location.reload();
  }, false);
};
