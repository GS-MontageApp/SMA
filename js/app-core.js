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

  // EXKLUSIVE UI-FILTERUNG FÜR DIE GRAFISCHE ANSICHT:
  // Wir greifen aus den gespeicherten Daten ausschließlich und gezielt auf "Tabelle1" zu.
  // Das "Auswahlblatt" wird grafisch komplett ignoriert.
  let targetRows = null;
  
  if (fileObj.sheets) {
    // Falls das multi-sheet Format vorliegt
    const sheetKeys = Object.keys(fileObj.sheets);
    const tab1Key = sheetKeys.find(k => k.toLowerCase() === 'tabelle1' || k.toLowerCase() === 'tabelle 1');
    if (tab1Key) {
      targetRows = fileObj.sheets[tab1Key];
    } else if (sheetKeys.length > 0) {
      targetRows = fileObj.sheets[sheetKeys[0]];
    }
  } 
  
  if (!targetRows && fileObj.rawData) {
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
