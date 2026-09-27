/**
 * === SCHLAUCHMANAGEMENT PWA - AUTHENTIFIZIERUNGS-MODUL ===
 * Lädt die Benutzer aus data/users.json und steuert die harte Eingangssperre.
 */

window.AuthManager = (() => {
  let users = [];

  // Lädt die Benutzer beim Start aus der JSON-Datei
  async function init() {
    try {
      const response = await fetch('data/users.json');
      if (!response.ok) {
        throw new Error('Fehler beim Laden der Benutzerdaten.');
      }
      users = await response.json();
      renderUserSelection();
      checkAuthentication();
    } catch (error) {
      console.error('AuthManager Fehler:', error);
      // Fallback, falls die JSON im Offline-Modus nicht geladen werden kann
      users = [
        { vorname: "Markus", nachname: "Jost", kurzform: "Markus" },
        { vorname: "Alexander", nachname: "Rippke", kurzform: "Alex" }
      ];
      renderUserSelection();
      checkAuthentication();
    }
  }

  // Befüllt das Auswahl-Modal / Dropdown mit den Benutzern aus der JSON
  function renderUserSelection() {
    const container = document.getElementById('user_selection_container');
    if (!container) return;

    container.innerHTML = '';
    users.forEach(user => {
      const btn = document.createElement('button');
      btn.className = 'w-full text-left p-3 bg-slate-50 hover:bg-emerald-50 rounded-xl font-semibold text-slate-700 transition-colors flex items-center justify-between border border-slate-200 shadow-xs';
      btn.innerHTML = `<span>👤 ${user.vorname} ${user.nachname}</span> <span class="text-xs font-mono bg-slate-200 px-2 py-1 rounded text-slate-600">${user.kurzform}</span>`;
      btn.onclick = () => loginUser(user);
      container.appendChild(btn);
    });
  }

  // Führt den Login für den ausgewählten Benutzer aus
  function loginUser(user) {
    localStorage.setItem('schlauchmanagement_current_user', JSON.stringify(user));
    closeAuthModal();
    updateUIState();
  }

  // Prüft, ob bereits ein Benutzer angemeldet ist (harte Eingangssperre)
  function checkAuthentication() {
    const savedUser = localStorage.getItem('schlauchmanagement_current_user');
    if (!savedUser) {
      openAuthModal();
    } else {
      updateUIState();
    }
  }

  // Aktualisiert die Anzeige im Header (zeigt den aktiven Bearbeiter)
  function updateUIState() {
    const savedUser = localStorage.getItem('schlauchmanagement_current_user');
    const badgeEl = document.getElementById('role_status_badge');
    const logoutItem = document.getElementById('logout_menu_item');

    if (savedUser && badgeEl) {
      const userObj = JSON.parse(savedUser);
      badgeEl.textContent = `~/${userObj.kurzform}`;
      badgeEl.className = 'px-2 py-0.5 rounded-md text-[10px] font-mono font-bold bg-emerald-600 text-white';
      badgeEl.classList.remove('hidden');
      if (logoutItem) logoutItem.classList.remove('hidden');
    }
  }

  // Öffnet das Login-Modal
  function openAuthModal() {
    const modal = document.getElementById('auth_modal');
    if (modal) modal.classList.remove('hidden');
  }

  // Schließt das Login-Modal
  function closeAuthModal() {
    const modal = document.getElementById('auth_modal');
    if (modal) modal.classList.add('hidden');
  }

  // Benutzer abmelden (Logout)
  function logout() {
    localStorage.removeItem('schlauchmanagement_current_user');
    const badgeEl = document.getElementById('role_status_badge');
    const logoutItem = document.getElementById('logout_menu_item');
    if (badgeEl) {
      badgeEl.textContent = '';
      badgeEl.classList.add('hidden');
    }
    if (logoutItem) logoutItem.classList.add('hidden');
    openAuthModal();
  }

  return {
    init,
    logout
  };
})();
