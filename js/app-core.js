/**
 * ============================================================================
 * MODUL: app-core.js (Schlauchmanagement-App v0.1.106)
 * ============================================================================
 * Kernlogik mit fester Spaltenstruktur (A bis R), Spalte-B-Schlauchfilter und UI-Steuerung.
 */

window.currentActiveCustomer = null;
window.currentActiveFileName = null;
window.openedFilesStack = [];
window.currentActiveCoordinateMapping = [];

window.VALID_SCHLAUCH_TYPES = [
  "1SN", "2SN", "4SP", "4SH", "R13", "R15", "462", 
  "1TE", "2TE", "3TE", "Minimess", "Teflon", "R4", "2245N"
];

window.MASTER_COLUMNS = [
  { index: 0, key: "kennz", label: "Kennz." },
  { index: 1, key: "schlauch", label: "Schlauch" },
  { index: 2, key: "nw", label: "NW" },
  { index: 3, key: "anschluss_a", label: "Anschluss A" },
  { index: 4, key: "anschluss_b", label: "Anschluss B" },
  { index: 5, key: "laenge", label: "Länge" },
  { index: 6, key: "lage_a", label: "Lage A" },
  { index: 7, key: "lage_b", label: "Lage B" },
  { index: 8, key: "max_druck", label: "max.<br>Druck<br>(Bar)" },
  { index: 9, key: "herstelldatum", label: "Herstell-<br>datum" },
  { index: 10, key: "sicherheits_bewertung", label: "Sicherheits-<br>technische<br>Bewertung" },
  { index: 11, key: "theor_lebensdauer", label: "Theor.<br>Lebensdauer" },
  { index: 12, key: "pruefung_am", label: "Prüfung<br>am" },
  { index: 13, key: "pruefung_status", label: "Prüfung*" },
  { index: 14, key: "naechste_pruefung", label: "Nächste<br>Prüfung" },
  { index: 15, key: "pruefer", label: "Prüfer" },
  { index: 16, key: "einbauort", label: "Einbauort" },
  { index: 17, key: "bemerkung", label: "Bemerkung" }
];

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
            "Tabelle1": [
              ["Gustav Schmidt", "", "Schlauchmanagement in Anlehnung an DGUV 113-020", "", "", "", "", "", "", "", "", "Betreiber", "", "", "", "Datum", "", "Unterschrift"],
              ["Kunde", "Hettich", "", "", "", "", "", "", "", "Version: 1", "", "", "", "", "", "", "", ""],
              ["Anlage", "Hubtisch 4", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
              ["Kennz.", "Schlauch", "NW", "Anschluss A", "Anschluss B", "Länge", "Lage A", "Lage B", "max. Druck (Bar)", "Herstelldatum", "Sicherheitstechnische Bewertung", "Theor. Lebensdauer", "Prüfung am", "Prüfung*", "Nächste Prüfung", "Prüfer", "Einbauort", "Bemerkung"],
              ["1", "2SN", "8", "DKOL8-10L", "DKOL8-10L-90°", "300", "0", "0", "350", "Apr. 26", "2", "72", "Apr. 26", "OK", "Apr. 27", "MJ/ML", "Halle 1", "OK"], 
              ["2", "2SN", "8", "DKOL8-10L", "DKOL8-10L-90°", "950", "0", "0", "350", "Jun. 25", "2", "72", "Apr. 26", "OK", "Apr. 27", "MJ/ML", "Halle 2", "Wartung"]  
            ]
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
  addFileToClient: function(clientName, fileName, fileObj) {
    let clients = this.getClients();
    if (!clients[clientName] || !Array.isArray(clients[clientName])) {
      clients[clientName] = [];
    }
    const nowStr = new Date().toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    
    const existingIdx = clients[clientName].findIndex(f => f && f.name === fileName);
    if (existingIdx >= 0) {
      clients[clientName][existingIdx].timestamp = nowStr;
      if (fileObj.sheets) clients[clientName][existingIdx].sheets = fileObj.sheets;
      if (fileObj.rawData) clients[clientName][existingIdx].rawData = fileObj.rawData;
    } else {
      clients[clientName].push({ name: fileName, timestamp: nowStr, sheets: fileObj.sheets, rawData: fileObj.rawData });
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
    console.error(e);
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

  let targetRows = null;
  if (fileObj.sheets) {
    if (fileObj.sheets["Tabelle1"] && Array.isArray(fileObj.sheets["Tabelle1"])) {
      targetRows = fileObj.sheets["Tabelle1"];
    } else {
      const keys = Object.keys(fileObj.sheets);
      if (keys.length > 0 && Array.isArray(fileObj.sheets[keys[0]])) {
        targetRows = fileObj.sheets[keys[0]];
      }
    }
  } else if (fileObj.rawData && Array.isArray(fileObj.rawData)) {
    targetRows = fileObj.rawData;
  }

  let rawData = [];
  let coordinateMapping = [];
  const masterHeaderRowIndex = 3;

  function isValidSchlauchRow(rowArray) {
    if (!Array.isArray(rowArray) || rowArray.length < 2) return false;
    const cellB = rowArray[1];
    if (cellB !== undefined && cellB !== null && cellB !== "") {
      const valStr = String(cellB).trim();
      return window.VALID_SCHLAUCH_TYPES.includes(valStr);
    }
    return false;
  }

  if (targetRows && Array.isArray(targetRows) && targetRows.length > 0) {
    const headerLabels = window.MASTER_COLUMNS.map(col => col.label);
    rawData.push(headerLabels);
    coordinateMapping.push({ coords: headerLabels.map((_, idx) => ({ originalRow: masterHeaderRowIndex, originalCol: idx })), isHeader: true });

    targetRows.forEach((row, rIndex) => {
      if (!Array.isArray(row)) return;
      if (rIndex === masterHeaderRowIndex) return;

      if (isValidSchlauchRow(row)) {
        let extractedSlice = [];
        let rowCoords = [];
        
        for (let c = 0; c < 18; c++) {
          extractedSlice.push(row[c] !== undefined && row[c] !== null ? row[c] : "");
          rowCoords.push({ originalRow: rIndex, originalCol: c });
        }

        rawData.push(extractedSlice);
        coordinateMapping.push({ coords: rowCoords, isSchlauch: true });
      }
    });
  } else {
    rawData = [window.MASTER_COLUMNS.map(c => c.label), ["Info", "Keine Tabellendaten verfügbar"]];
    coordinateMapping = [];
  }

  window.currentActiveCoordinateMapping = coordinateMapping;

  document.querySelectorAll('.app-view').forEach(el => el.classList.add('hidden'));
  const stageView = document.getElementById('view-buehne');
  if (stageView) stageView.classList.remove('hidden');

  const titleEl = document.getElementById('header-title');
  if (titleEl) titleEl.textContent = `${clientName} / ${fileName}`;

  const container = document.getElementById('buehne_table_container');
  if (!container) return;
  container.innerHTML = '';

  const wrapper = document.createElement('div');
  wrapper.className = 'w-full overflow-x-auto overflow-y-auto h-[calc(100vh-140px)] bg-white shadow-none';

  const table = document.createElement('table');
  table.className = 'w-full text-left border-collapse text-xs sm:text-sm text-slate-700 min-w-max';

  const thead = document.createElement('thead');
  thead.className = 'sticky top-0 bg-slate-100 text-slate-800 font-bold border-b border-slate-300 shadow-xs z-10';
  
  const tbody = document.createElement('tbody');
  tbody.className = 'divide-y divide-slate-100';

  rawData.forEach((row, rowIndex) => {
    const isHeader = (rowIndex === 0);
    const tr = document.createElement('tr');

    if (isHeader) {
      tr.className = 'bg-slate-200 font-bold border-b-2 border-slate-400';
    } else {
      tr.className = 'hover:bg-slate-50 transition-colors';
    }

    row.forEach((cellVal, colIndex) => {
      const cell = isHeader ? document.createElement('th') : document.createElement('td');
      cell.className = 'px-4 py-3 whitespace-nowrap ' + (isHeader ? 'font-bold text-slate-900 bg-slate-200' : 'text-slate-600');
      
      if (isHeader) {
        cell.innerHTML = cellVal !== undefined && cellVal !== null ? cellVal : '';
      } else {
        const input = document.createElement('input');
        input.type = 'text';
        input.value = cellVal !== undefined && cellVal !== null ? cellVal : '';
        input.className = 'w-full bg-transparent border-0 focus:ring-1 focus:ring-indigo-500 rounded px-1 py-0.5 text-xs sm:text-sm text-slate-700';
        
        input.oninput = function(e) {
          const newVal = e.target.value;
          if (window.currentActiveCoordinateMapping && window.currentActiveCoordinateMapping[rowIndex]) {
            const coordEntry = window.currentActiveCoordinateMapping[rowIndex];
            if (coordEntry && coordEntry.coords && coordEntry.coords[colIndex]) {
              const targetCoord = coordEntry.coords[colIndex];
              const clients = window.AppData.getClients();
              const fileList = clients[window.currentActiveCustomer] || [];
              const fileObj = fileList.find(f => f && f.name === window.currentActiveFileName);
              if (fileObj && fileObj.sheets) {
                const sheetKeys = Object.keys(fileObj.sheets);
                if (sheetKeys.length > 0) {
                  const activeSheetName = sheetKeys[0];
                  if (fileObj.sheets[activeSheetName][targetCoord.originalRow]) {
                    fileObj.sheets[activeSheetName][targetCoord.originalRow][targetCoord.originalCol] = newVal;
                    window.AppData.addFileToClient(window.currentActiveCustomer, window.currentActiveFileName, fileObj);
                  }
                }
              }
            }
          }
        };

        cell.appendChild(input);
      }
      tr.appendChild(cell);
    });

    if (isHeader) {
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
      window.AppData.addFileToClient(window.currentActiveCustomer, window.currentActiveFileName, fileObj);
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
    }
  );
};

window.showSystemModal = function(title, message, onConfirm, showCancel = true, isCompact = false) {
  const modal = document.getElementById('system_generic_modal');
  const card = document.getElementById('system_modal_card');
  const headerEl = document.getElementById('system_modal_header');
  const titleEl = document.getElementById('system_modal_title');
  const msgEl = document.getElementById('system_modal_message');
  const confirmBtn = document.getElementById('system_modal_confirm_btn');
  const cancelBtn = document.getElementById('system_modal_cancel_btn');
  const footerEl = document.getElementById('system_modal_footer');

  if (!modal || !card) return;

  // Nur im Export-Kompaktmodus (isCompact = true) wird das Modal randlos und ohne Buttons/Header gerendert
  if (isCompact) {
    card.className = 'bg-transparent shadow-none p-0 m-0 max-w-max w-auto flex items-center justify-center animate-in fade-in zoom-in-95 duration-200';
    if (headerEl) headerEl.classList.add('hidden');
    if (cancelBtn) cancelBtn.classList.add('hidden');
    if (confirmBtn) confirmBtn.classList.add('hidden');
    if (footerEl) footerEl.classList.add('hidden');
  } else {
    card.className = 'bg-white rounded-2xl p-6 max-w-sm w-full space-y-4 text-slate-800 shadow-2xl animate-in fade-in zoom-in-95 duration-200';
    if (headerEl) headerEl.classList.remove('hidden');
    if (footerEl) footerEl.classList.remove('hidden');
    if (showCancel) {
      if (cancelBtn) cancelBtn.classList.remove('hidden');
    } else {
      if (cancelBtn) cancelBtn.classList.add('hidden');
    }
    // Sicherstellen, dass Bestätigen- und Abbrechen-Buttons für normale Modals aktiv sind
    if (confirmBtn) confirmBtn.classList.remove('hidden');
  }

  titleEl.textContent = title;
  
  if (message.includes('<img') || message.includes('<div')) {
    msgEl.innerHTML = message;
  } else {
    msgEl.textContent = message;
  }

  const newConfirmBtn = confirmBtn.cloneNode(true);
  confirmBtn.parentNode.replaceChild(newConfirmBtn, confirmBtn);

  newConfirmBtn.onclick = function() {
    modal.classList.add('hidden');
    if (typeof onConfirm === 'function') onConfirm();
  };

  if (cancelBtn) {
    cancelBtn.onclick = function() {
      modal.classList.add('hidden');
    };
  }

  modal.classList.remove('hidden');
};

window.openTopMenu = function() {
  const menu = document.getElementById('top_menu_modal');
  if (menu) menu.classList.remove('hidden');
};

window.closeTopMenu = function() {
  const menu = document.getElementById('top_menu_modal');
  if (menu) menu.classList.add('hidden');
};
