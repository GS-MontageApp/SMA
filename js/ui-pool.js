/**
 * ============================================================================
 * MODUL: ui-pool.js (Schlauchmanagement-App v0.1.28)
 * ============================================================================
 * Kapselt die UI-Rendering-Routinen für den Dateipool und die Kunden-Dateiliste.
 */

window.UIPool = {
  renderDateipool: function() {
    const container = document.getElementById('dateipool_grid');
    if (!container) return;
    container.innerHTML = '';
    const clients = window.AppData.getClients();

    const customerKeys = Object.keys(clients);
    if (customerKeys.length === 0) {
      container.innerHTML = '<p class="text-xs text-slate-400 py-3 col-span-full text-center">Keine Kunden im Dateipool vorhanden.</p>';
      return;
    }

    customerKeys.forEach(clientName => {
      const fileList = clients[clientName] || [];
      const fileCount = fileList.length;
      const card = document.createElement('div');
      card.className = 'p-4 bg-slate-50 hover:bg-emerald-50/50 rounded-xl border-2 border-slate-200 hover:border-emerald-500 transition-all shadow-xs flex items-center justify-between group';
      
      const infoDiv = document.createElement('div');
      infoDiv.className = 'flex-1 cursor-pointer pr-2';
      infoDiv.innerHTML = `<h3 class="font-bold text-slate-700 text-base truncate">${clientName}</h3><p class="text-xs text-slate-400 mt-0.5">${fileCount} ${fileCount === 1 ? 'Datei' : 'Dateien'}</p>`;
      infoDiv.onclick = () => window.openCustomerFiles(clientName);
      card.appendChild(infoDiv);

      const deleteBtn = document.createElement('button');
      deleteBtn.className = 'p-2.5 bg-red-50 hover:bg-red-100 text-red-700 rounded-xl transition-colors text-sm shrink-0 shadow-2xs';
      deleteBtn.title = 'Kunden löschen';
      deleteBtn.innerHTML = '🗑️';
      deleteBtn.onclick = (e) => {
        e.stopPropagation();
        window.confirmDeleteClient(clientName, fileCount);
      };
      card.appendChild(deleteBtn);

      container.appendChild(card);
    });
  },

  renderCustomerFilesList: function(currentActiveCustomer) {
    const listContainer = document.getElementById('customer_files_list');
    if (!listContainer || !currentActiveCustomer) return;
    listContainer.innerHTML = '';

    const clients = window.AppData.getClients();
    const files = clients[currentActiveCustomer] || [];

    if (files.length === 0) {
      listContainer.innerHTML = '<p class="text-xs text-slate-400 py-3 text-center">Keine Dateien vorhanden.</p>';
      return;
    }

    files.forEach(fileObj => {
      if (!fileObj) return;
      const fileName = fileObj.name;
      const timestamp = fileObj.timestamp || 'Unbekannt';

      const card = document.createElement('div');
      card.className = 'p-4 bg-slate-50 hover:bg-emerald-50/50 rounded-xl border-2 border-slate-200 hover:border-emerald-500 transition-all shadow-xs flex flex-col justify-between gap-3 cursor-pointer relative';
      
      card.onclick = () => window.openFileOnStage(currentActiveCustomer, fileName, true);

      const infoDiv = document.createElement('div');
      infoDiv.className = 'flex-1 min-w-0 pr-12';
      infoDiv.innerHTML = `<h3 class="font-bold text-slate-700 text-base break-all">${fileName}</h3><p class="text-xs text-slate-500 mt-1">Zuletzt lokal gespeichert: ${timestamp}</p>`;
      card.appendChild(infoDiv);

      const delFileBtn = document.createElement('button');
      delFileBtn.className = 'absolute bottom-3 right-3 p-2.5 bg-red-50 hover:bg-red-100 text-red-700 rounded-xl transition-colors text-sm shadow-2xs';
      delFileBtn.title = 'Datei löschen';
      delFileBtn.innerHTML = '🗑️';
      delFileBtn.onclick = (e) => {
        e.stopPropagation();
        window.confirmDeleteFile(currentActiveCustomer, fileName);
      };
      card.appendChild(delFileBtn);

      listContainer.appendChild(card);
    });
  }
};
