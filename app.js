(() => {
  "use strict";

  const CONFIG = {
    appName: "bbz CRM",

    graph: {
      tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
      clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
      authority: "https://login.microsoftonline.com/3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
      redirectUri: "https://markusbaechler.github.io/crm-spa/",
      // FIX 3a: Scope auf ReadWrite erweitert — verhindert zweiten Login-Prompt beim Write-Layer
      scopes: ["User.Read", "Sites.ReadWrite.All"]
    },

    sharePoint: {
      siteHostname: "bbzsg.sharepoint.com",
      sitePath: "/sites/CRM"
    },

    lists: {
      firms: "CRMFirms",
      contacts: "CRMContacts",
      history: "CRMHistory",
      tasks: "CRMTasks"
    },

    defaults: {
      route: "firms",
      contactArchiveDefaultHidden: true,
      planningShowOnlyOpen: true,
      // Firma für Privatpersonen ohne Firmenbezug — exakter SP-Titel
      privateFirmTitle: "Privatpersonen"
    }
  };

  const SCHEMA = {
    firms: {
      listTitle: CONFIG.lists.firms,
      fields: {
        title: "Title",
        adresse: "Adresse",
        plz: "PLZ",
        ort: "Ort",
        land: "Land",
        hauptnummer: "Hauptnummer",
        klassifizierung: "Klassifizierung",
        vip: "VIP"
      }
    },

    contacts: {
      listTitle: CONFIG.lists.contacts,
      fields: {
        nachname: "Title",
        vorname: "Vorname",
        anrede: "Anrede",
        firma: "Firma",
        firmaLookupId: "FirmaLookupId",
        funktion: "Funktion",
        email1: "Email1",
        email2: "Email2",
        direktwahl: "Direktwahl",
        mobile: "Mobile",
        rolle: "Rolle",
        leadbbz0: "Leadbbz0",
        sgf: "SGF",
        geburtstag: "Geburtstag",
        kommentar: "Kommentar",
        event: "Event",
        eventhistory: "Eventhistory",
        archiviert: "Archiviert"
      }
    },

    history: {
      listTitle: CONFIG.lists.history,
      fields: {
        title: "Title",
        kontakt: "Nachname",
        kontaktLookupId: "NachnameLookupId",
        datum: "Datum",
        // KORREKTUR: SP-Feldname ist "Kontaktart", nicht "Typ"
        typ: "Kontaktart",
        notizen: "Notizen",
        projektbezug: "Projektbezug",
        leadbbz: "Leadbbz"
      }
    },

    tasks: {
      listTitle: CONFIG.lists.tasks,
      fields: {
        title: "Title",
        kontakt: "Name",
        kontaktLookupId: "NameLookupId",
        deadline: "Deadline",
        status: "Status",
        leadbbz: "Leadbbz"
      }
    }
  };

  const state = {
    auth: {
      msal: null,
      account: null,
      token: null,
      isAuthenticated: false,
      isReady: false
    },

    meta: {
      siteId: null,
      loading: false,
      lastError: null,
      // Choice-Werte aus SharePoint — pro Liste, pro SP-Feldname
      // Struktur: { "CRMContacts": { "Anrede": ["Herr", "Frau", ...], ... }, ... }
      choices: {},
      // ID der Firma "Privatpersonen" — wird nach enrich() automatisch gesetzt
      privateFirmId: null
    },

    data: {
      firms: [],
      contacts: [],
      history: [],
      tasks: []
    },

    enriched: {
      firms: [],
      contacts: [],
      history: [],
      tasks: [],
      events: []
    },

    filters: {
      route: CONFIG.defaults.route,
      firms: { search: "", klassifizierung: "", vip: "", onlyPrivat: false, sortBy: "title", sortDir: "asc", radarMode: false },
      contacts: { search: "", archiviertAusblenden: CONFIG.defaults.contactArchiveDefaultHidden, sortBy: "fullName", sortDir: "asc" },
      planning: { search: "", onlyOpen: CONFIG.defaults.planningShowOnlyOpen, groupBy: "none", sortBy: "deadline", sortDir: "asc", segment: "", leadbbz: "", faelligkeit: "" },
      history: { search: "", kontaktart: "", leadbbz: "", groupBy: "date", zeitfenster: "", radarMode: false },
      events: { search: "", onlyWithOpenTasks: false, sortBy: "contactName", sortDir: "asc", segment: "", selectedEvent: "" }
    },

    selection: {
      firmId: null,
      contactId: null
    },

    // Modal-State fuer Write-Layer
    modal: null
  };

  const helpers = {
    escapeHtml(value) {
      return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
    },

    bool(value) {
      if (typeof value === "boolean") return value;
      if (typeof value === "number") return value === 1;
      if (typeof value === "string") {
        const v = value.trim().toLowerCase();
        return ["true", "1", "ja", "yes"].includes(v);
      }
      return false;
    },

    isEmpty(value) {
      return value === null || value === undefined || value === "";
    },

    toArray(value) {
      if (Array.isArray(value)) return value;
      if (value === null || value === undefined || value === "") return [];
      if (typeof value === "string") {
        if (value.includes(";#")) return value.split(";#").map(v => v.trim()).filter(Boolean);
        if (value.includes(",")) return value.split(",").map(v => v.trim()).filter(Boolean);
        return [value.trim()].filter(Boolean);
      }
      return [value];
    },

    normalizeChoiceList(value) {
      return helpers.toArray(value).filter(Boolean);
    },

    toDate(value) {
      if (!value) return null;
      const d = new Date(value);
      return Number.isNaN(d.getTime()) ? null : d;
    },

    formatDate(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      return d.toLocaleDateString("de-CH", { day: "2-digit", month: "2-digit", year: "numeric" });
    },

    formatDateTime(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      return d.toLocaleString("de-CH", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
    },

    // FIX 1: fehlende Hilfsfunktion fuer <input type="date"> — gibt YYYY-MM-DD zurueck
    // Wichtig: Lokale Datum-Komponenten verwenden (nicht toISOString = UTC),
    // sonst verschiebt sich das Datum in Zeitzonen wie CH (UTC+1) um einen Tag
    toDateInput(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      return `${y}-${m}-${day}`;
    },

    todayStart() {
      const d = new Date();
      d.setHours(0, 0, 0, 0);
      return d;
    },

    isOpenTask(status) {
      const v = String(status || "").trim().toLowerCase();
      return !["erledigt", "geschlossen", "completed", "done", "closed"].includes(v);
    },

    isOverdue(deadline) {
      const d = helpers.toDate(deadline);
      if (!d) return false;
      return d < helpers.todayStart();
    },

    compareDateAsc(a, b) {
      const ad = helpers.toDate(a), bd = helpers.toDate(b);
      if (!ad && !bd) return 0;
      if (!ad) return 1;
      if (!bd) return -1;
      return ad - bd;
    },

    compareDateDesc(a, b) {
      const ad = helpers.toDate(a), bd = helpers.toDate(b);
      if (!ad && !bd) return 0;
      if (!ad) return 1;
      if (!bd) return -1;
      return bd - ad;
    },

    textIncludes(haystack, needle) {
      return String(haystack || "").toLowerCase().includes(String(needle || "").toLowerCase());
    },

    joinNonEmpty(values, sep = " · ") {
      return values.filter(v => !helpers.isEmpty(v)).join(sep);
    },

    fullName(contact) {
      return helpers.joinNonEmpty([contact.vorname, contact.nachname], " ").trim();
    },

    firmBadgeClass(value) {
      const v = String(value || "").toUpperCase();
      if (v === "A" || v === "A-KUNDE") return "bbz-pill bbz-pill-a";
      if (v === "B" || v === "B-KUNDE") return "bbz-pill bbz-pill-b";
      if (v === "C" || v === "C-KUNDE") return "bbz-pill bbz-pill-c";
      return "bbz-pill";
    },

    // Leadbbz als farbiges Pill
    leadbbzBadgeHtml(value) {
      if (!value) return '<span class="bbz-muted">—</span>';
      return `<span class="bbz-pill bbz-pill-lead">${helpers.escapeHtml(value)}</span>`;
    },

    // Detailband-Klasse je nach Segment und VIP
    detailBandClass(firm) {
      if (!firm) return "bbz-detail-band-default";
      if (firm.vip) return "bbz-detail-band bbz-detail-band-vip";
      const v = String(firm.klassifizierung || "").toUpperCase();
      if (v.includes("A")) return "bbz-detail-band bbz-detail-band-a";
      if (v.includes("B")) return "bbz-detail-band bbz-detail-band-b";
      if (v.includes("C")) return "bbz-detail-band bbz-detail-band-c";
      return "bbz-detail-band bbz-detail-band-default";
    },

    statusClass(status, deadline) {
      if (!helpers.isOpenTask(status)) return "bbz-success";
      if (helpers.isOverdue(deadline)) return "bbz-danger";
      return "bbz-warning";
    },

    multiChoiceHtml(values) {
      const list = helpers.normalizeChoiceList(values);
      if (!list.length) return '<span class="bbz-muted">—</span>';
      return list.map(v => `<span class="bbz-chip">${helpers.escapeHtml(v)}</span>`).join("");
    },

    // Avatar-Initialen: gibt fertiges HTML-Element zurück
    // Farbe wird deterministisch aus dem Namen gehasht (0–5)
    avatarHtml(contact) {
      const first = String(contact.vorname || "").charAt(0).toUpperCase();
      const last  = String(contact.nachname || "").charAt(0).toUpperCase();
      const initials = (first + last) || "?";
      // Einfacher Hash aus Zeichencodes
      const seed = [...initials].reduce((s, c) => s + c.charCodeAt(0), 0);
      const idx  = seed % 6;
      return `<span class="bbz-avatar" data-idx="${idx}">${helpers.escapeHtml(initials)}</span>`;
    },

    // Status-Chip: gibt ein farbiges Pill-HTML zurück
    // status: Taskstatus-String, deadline: ISO-Datum
    statusChipHtml(status, deadline) {
      if (!helpers.isOpenTask(status)) {
        return `<span class="bbz-status-chip bbz-status-done">${helpers.escapeHtml(status || "Erledigt")}</span>`;
      }
      if (helpers.isOverdue(deadline)) {
        return `<span class="bbz-status-chip bbz-status-overdue">${helpers.escapeHtml(status || "Überfällig")}</span>`;
      }
      return `<span class="bbz-status-chip bbz-status-open">${helpers.escapeHtml(status || "Offen")}</span>`;
    },

    // Relatives Datum: "heute", "gestern", "vor 3 Tagen", "vor 2 Wochen"
    // Fällt nach 60 Tagen auf formatDate zurück
    relativeDate(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      const today = helpers.todayStart();
      const diffMs = today - d;
      const diffDays = Math.floor(diffMs / 86400000);
      if (diffDays < 0) {
        const futureDays = Math.abs(diffDays);
        if (futureDays === 1) return "morgen";
        if (futureDays < 7) return `in ${futureDays} Tagen`;
        if (futureDays < 14) return "nächste Woche";
        return helpers.formatDate(value);
      }
      if (diffDays === 0) return "heute";
      if (diffDays === 1) return "gestern";
      if (diffDays < 7) return `vor ${diffDays} Tagen`;
      if (diffDays < 14) return "vor 1 Woche";
      if (diffDays < 30) return `vor ${Math.floor(diffDays / 7)} Wochen`;
      if (diffDays < 60) return `vor ${Math.floor(diffDays / 30)} Monat${Math.floor(diffDays / 30) > 1 ? "en" : ""}`;
      return helpers.formatDate(value);
    },

    // Aktivitäts-Signal für Firmenliste und Pflege-Radar:
    // gibt "" | "overdue" | "never" | "cold" | "ok" zurück
    // "overdue" — offene überfällige Tasks (alle Segmente)
    // "never"   — A-Kunde, noch kein History-Eintrag
    // "cold"    — A/B-Kunde, kein Kontakt seit >360 Tagen
    // "ok"      — A/B-Kunde, letzter Kontakt <90 Tage, keine überfälligen Tasks
    // ""        — C-Kunde/keine Klassifizierung: kein Signal
    firmSignal(firm) {
      if (firm.openTasksCount > 0 && firm.tasks.some(t => t.isOpen && t.isOverdue)) return "overdue";
      const kl = String(firm.klassifizierung || "").toUpperCase();
      const isA = kl.includes("A");
      const isB = kl.includes("B");
      if (!isA && !isB) return "";
      if (firm.history.length === 0) return isA ? "never" : "";
      const last = helpers.toDate(firm.latestActivity);
      if (!last) return "never";
      const diffDays = Math.floor((helpers.todayStart() - last) / 86400000);
      if (diffDays > 360) return "cold";
      if (diffDays <= 90) return "ok";
      return "";
    },

    // Debounce: verhindert excessive DOM-Rebuilds beim Tippen in Suchfeldern
    debounce(fn, ms = 150) {
      let timer = null;
      return (...args) => {
        clearTimeout(timer);
        timer = setTimeout(() => fn(...args), ms);
      };
    },

    // Rendert ein <select> aus SP-Choices — fällt auf <input> zurück wenn keine Choices geladen
    choiceSelectHtml(name, listTitle, spFieldName, currentValue, required = false) {
      const choices = state.meta.choices?.[listTitle]?.[spFieldName] || [];
      if (!choices.length) {
        // Fallback: Freitext — tritt auf wenn Choices noch nicht geladen oder SP-Feld kein Choice
        return `<input class="bbz-input" name="${name}" value="${helpers.escapeHtml(currentValue || "")}" ${required ? "required" : ""} placeholder="Wird geladen..." />`;
      }
      return `
        <select class="bbz-select" name="${name}" ${required ? "required" : ""}>
          <option value="">— bitte wählen —</option>
          ${choices.map(c => `<option value="${helpers.escapeHtml(c)}" ${currentValue === c ? "selected" : ""}>${helpers.escapeHtml(c)}</option>`).join("")}
        </select>
      `;
    },

    // Rendert Checkboxen für Multi-Choice-Felder aus SP
    // currentValues: string[] der aktuell gesetzten Werte
    choiceMultiHtml(name, listTitle, spFieldName, currentValues) {
      const choices = state.meta.choices?.[listTitle]?.[spFieldName] || [];
      const selected = new Set(Array.isArray(currentValues) ? currentValues : []);
      if (!choices.length) {
        return `<input class="bbz-input" name="${name}" value="${helpers.escapeHtml([...selected].join(", "))}" placeholder="Wird geladen..." />`;
      }
      return `
        <div class="bbz-multi-choice">
          ${choices.map(c => `
            <label class="bbz-multi-choice-item">
              <input type="checkbox" name="${name}" value="${helpers.escapeHtml(c)}" ${selected.has(c) ? "checked" : ""} />
              <span>${helpers.escapeHtml(c)}</span>
            </label>
          `).join("")}
        </div>
      `;
    },

    ensureMsalAvailable() {
      if (!window.msal || !window.msal.PublicClientApplication) {
        throw new Error("MSAL-Bibliothek wurde nicht geladen.");
      }
    },

    validateConfig() {
      const missing = [];
      if (!CONFIG.graph.clientId) missing.push("clientId");
      if (!CONFIG.graph.tenantId) missing.push("tenantId");
      if (!CONFIG.graph.authority) missing.push("authority");
      if (!CONFIG.graph.redirectUri) missing.push("redirectUri");
      if (!CONFIG.sharePoint.siteHostname) missing.push("sharePoint.siteHostname");
      if (!CONFIG.sharePoint.sitePath) missing.push("sharePoint.sitePath");
      if (missing.length) throw new Error(`Konfiguration unvollstaendig: ${missing.join(", ")}`);
    }
  };

  const ui = {
    els: {
      viewRoot: null,
      authStatus: null,
      globalMessage: null,
      btnLogin: null,
      btnRefresh: null,
      navButtons: []
    },

    init() {
      this.els.viewRoot = document.getElementById("view-root");
      this.els.authStatus = document.getElementById("auth-status");
      this.els.globalMessage = document.getElementById("global-message");
      this.els.btnLogin = document.getElementById("btn-login");
      this.els.btnRefresh = document.getElementById("btn-refresh");
      this.els.navButtons = [...document.querySelectorAll(".bbz-nav-btn")];

      if (this.els.btnLogin) this.els.btnLogin.addEventListener("click", () => controller.handleLogin());
      if (this.els.btnRefresh) this.els.btnRefresh.addEventListener("click", () => controller.handleRefresh());

      this.els.navButtons.forEach(btn => {
        btn.addEventListener("click", () => controller.navigate(btn.dataset.route));
      });

      // Zentraler Click-Handler
      document.addEventListener("click", (event) => {
        const openFirm = event.target.closest("[data-action='open-firm']");
        if (openFirm) { controller.openFirm(openFirm.dataset.id); return; }

        const navPlanning = event.target.closest("[data-action='navigate-planning']");
        if (navPlanning) { event.preventDefault(); controller.navigate("planning"); return; }

        // KPI-Schnellfilter — setzt Filter und navigiert bei Bedarf
        const kpiFilter = event.target.closest("[data-action='kpi-filter']");
        if (kpiFilter) {
          const scope = kpiFilter.dataset.scope;
          const value = kpiFilter.dataset.value;
          if (scope === "firms-klassifizierung") {
            state.filters.firms.klassifizierung = state.filters.firms.klassifizierung === value ? "" : value;
            if (value === "") { state.filters.firms.vip = ""; state.filters.firms.onlyPrivat = false; }
            state.filters.firms.search = "";
            state.filters.route = "firms";
            state.selection.firmId = null;
          } else if (scope === "firms-radar") {
            state.filters.firms.radarMode = !state.filters.firms.radarMode;
            if (state.filters.firms.radarMode) {
              state.filters.firms.search = "";
              state.filters.firms.klassifizierung = "";
              state.filters.firms.vip = "";
              state.filters.firms.onlyPrivat = false;
            }
            state.filters.route = "firms";
            state.selection.firmId = null;
          } else if (scope === "firms-privat") {
            // Privat ist additiver Toggle — unabhängig von Segment
            state.filters.firms.onlyPrivat = !state.filters.firms.onlyPrivat;
            state.filters.firms.vip = ""; // VIP und Privat schliessen sich aus
            state.filters.firms.route = "firms";
            state.selection.firmId = null;
          } else if (scope === "contacts-mode") {
            // direkt route setzen — NICHT controller.navigate() da das _kpiMode zurücksetzt
            const newMode = state.filters.contacts._kpiMode === value ? "all" : value;
            state.filters.contacts._kpiMode = newMode;
            state.filters.route = "contacts";
            state.selection.contactId = null;
            state.modal = null;
          } else if (scope === "firms-vip") {
            // VIP ist additiver Toggle — unabhängig von Segment
            state.filters.firms.vip = state.filters.firms.vip === "yes" ? "" : "yes";
            state.filters.firms.onlyPrivat = false;
            state.filters.route = "firms";
            state.selection.firmId = null;
          } else if (scope === "history-zeitfenster") {
            state.filters.history.zeitfenster = state.filters.history.zeitfenster === value ? "" : value;
          } else if (scope === "history-kontaktart") {
            state.filters.history.kontaktart = state.filters.history.kontaktart === value ? "" : value;
          } else if (scope === "history-leadbbz") {
            state.filters.history.leadbbz = state.filters.history.leadbbz === value ? "" : value;
          } else if (scope === "history-radar") {
            state.filters.history.radarMode = !state.filters.history.radarMode;
            if (state.filters.history.radarMode) {
              state.filters.history.search = "";
              state.filters.history.zeitfenster = "";
            }
          } else if (scope === "planning-faelligkeit") {
            state.filters.planning.faelligkeit = state.filters.planning.faelligkeit === value ? "" : value;
          } else if (scope === "planning-segment") {
            state.filters.planning.segment = state.filters.planning.segment === value ? "" : value;
          } else if (scope === "planning-leadbbz") {
            state.filters.planning.leadbbz = state.filters.planning.leadbbz === value ? "" : value;
          } else if (scope === "events-segment") {
            state.filters.events.segment = state.filters.events.segment === value ? "" : value;
          } else if (scope === "events-selected") {
            state.filters.events.selectedEvent = state.filters.events.selectedEvent === value ? "" : value;
          } else if (scope === "navigate") {
            controller.navigate(value);
            return;
          }
          controller.render();
          return;
        }

        const openContact = event.target.closest("[data-action='open-contact']");
        if (openContact) { controller.openContact(openContact.dataset.id); return; }

        const backToFirms = event.target.closest("[data-action='back-to-firms']");
        if (backToFirms) { controller.navigate("firms"); return; }

        const backToContacts = event.target.closest("[data-action='back-to-contacts']");
        if (backToContacts) { controller.navigate("contacts"); return; }

        const openForm = event.target.closest("[data-action='open-contact-form']");
        if (openForm) {
          const itemId = openForm.dataset.itemId ? Number(openForm.dataset.itemId) : null;
          const firmId = openForm.dataset.firmId ? Number(openForm.dataset.firmId) : null;
          controller.openContactForm(itemId, firmId);
          return;
        }

        // FIX 2a: Modal schliessen via Button oder Backdrop-Klick
        const closeModal = event.target.closest("[data-close-modal]");
        if (closeModal) { controller.closeModal(); return; }

        const backdrop = event.target.closest(".bbz-modal-backdrop");
        if (backdrop && !event.target.closest(".bbz-modal")) { controller.closeModal(); return; }

        // Kontakt löschen
        const deleteContact = event.target.closest("[data-action='delete-contact']");
        if (deleteContact) {
          controller.handleDeleteContact(deleteContact.dataset.id, deleteContact.dataset.name);
          return;
        }

        // Firma bearbeiten
        const openFirmForm = event.target.closest("[data-action='open-firm-form']");
        if (openFirmForm) {
          controller.openFirmForm(openFirmForm.dataset.id);
          return;
        }

        // History-Formular öffnen
        const openHistoryForm = event.target.closest("[data-action='open-history-form']");
        if (openHistoryForm) {
          const contactId = openHistoryForm.dataset.contactId || null;
          const firmId    = openHistoryForm.dataset.firmId    || null;
          // Guard: Firma ohne Kontakte — kein Modal öffnen
          if (!contactId && firmId) {
            const firm = dataModel.getFirmById(Number(firmId));
            if (firm && firm.contacts.length === 0) {
              ui.setMessage(`"${firm.title}" hat noch keine Kontakte. Bitte zuerst einen Kontakt erfassen.`, "error");
              return;
            }
          }
          controller.openHistoryForm(contactId ? Number(contactId) : null, firmId ? Number(firmId) : null, null);
          return;
        }

        // History-Eintrag bearbeiten
        const editHistory = event.target.closest("[data-action='edit-history']");
        if (editHistory) {
          controller.openHistoryForm(null, null, Number(editHistory.dataset.id));
          return;
        }

        // History-Eintrag löschen
        const deleteHistory = event.target.closest("[data-action='delete-history']");
        if (deleteHistory) {
          controller.handleDeleteHistory(deleteHistory.dataset.id, deleteHistory.dataset.title);
          return;
        }

        // Task-Formular öffnen
        const openTaskForm = event.target.closest("[data-action='open-task-form']");
        if (openTaskForm) {
          const contactId = openTaskForm.dataset.contactId || null;
          const firmId    = openTaskForm.dataset.firmId    || null;
          // Guard: Firma ohne Kontakte — kein Modal öffnen
          if (!contactId && firmId) {
            const firm = dataModel.getFirmById(Number(firmId));
            if (firm && firm.contacts.length === 0) {
              ui.setMessage(`"${firm.title}" hat noch keine Kontakte. Bitte zuerst einen Kontakt erfassen.`, "error");
              return;
            }
          }
          controller.openTaskForm(contactId ? Number(contactId) : null, firmId ? Number(firmId) : null, null);
          return;
        }

        // Task bearbeiten
        const editTask = event.target.closest("[data-action='edit-task']");
        if (editTask) {
          controller.openTaskForm(null, null, Number(editTask.dataset.id));
          return;
        }

        // Task löschen
        const deleteTask = event.target.closest("[data-action='delete-task']");
        if (deleteTask) {
          controller.handleDeleteTask(deleteTask.dataset.id, deleteTask.dataset.title);
          return;
        }

        // Spalten-Sortierung (Planung)
        const setSort = event.target.closest("[data-action='set-sort']");
        if (setSort) {
          const col = setSort.dataset.col;
          const scope = setSort.dataset.scope || "planning";
          const f = scope === "firms" ? state.filters.firms : scope === "contacts" ? state.filters.contacts : state.filters.planning;
          if (f.sortBy === col) {
            f.sortDir = f.sortDir === "asc" ? "desc" : "asc";
          } else {
            f.sortBy = col;
            f.sortDir = "asc";
          }
          controller.render();
          return;
        }

        // Firma löschen
        const deleteFirm = event.target.closest("[data-action='delete-firm']");
        if (deleteFirm) {
          if (Number(deleteFirm.dataset.contacts) > 0) {
            ui.setMessage("Diese Firma hat noch Kontakte und kann nicht gelöscht werden.", "error");
            return;
          }
          controller.handleDeleteFirm(deleteFirm.dataset.id, deleteFirm.dataset.name);
          return;
        }

        // History: Notizen aufklappen/zuklappen (kein Re-Render)
        const expandBtn = event.target.closest("[data-action='toggle-expand']");
        if (expandBtn) {
          const card = expandBtn.closest(".bbz-timeline-item");
          if (card) {
            card.classList.toggle("bbz-expanded");
            expandBtn.textContent = card.classList.contains("bbz-expanded") ? "weniger" : "mehr";
          }
          return;
        }

        // History Pflege-Radar: Firma-Filter setzen
        const radarFirm = event.target.closest("[data-action='history-firma-filter']");
        if (radarFirm) {
          const firmTitle = radarFirm.dataset.firmTitle || "";
          // Toggle: nochmals klicken = Filter aufheben
          state.filters.history.search = state.filters.history.search === firmTitle ? "" : firmTitle;
          controller.render();
          return;
        }

        // Batch-Event-Dialog öffnen
        const openBatchEvent = event.target.closest("[data-action='open-batch-event']");
        if (openBatchEvent) {
          const eventName = openBatchEvent.dataset.eventName || "";
          const mode = openBatchEvent.dataset.mode || "anmelden";
          controller.openBatchEventForm(eventName, mode);
          return;
        }

        // Hilfsfunktion: Zähler, Submit-Button und Alle-Checkbox synchronisieren
        const syncBatchUI = (sel) => {
          const payload = state.modal?.payload;
          if (!payload) return;
          const preview = payload.previewContacts || [];
          const max = preview.length;
          const form = document.querySelector("[data-modal-form='batch-event']");
          const isEH = form?.dataset.mode === "eventhistory";
          const cat  = form?.dataset.eventName || "";

          const submitBtn = form?.querySelector("button[type='submit']");
          if (submitBtn) {
            submitBtn.textContent = isEH
              ? `+ ${sel.length} × Eventhistory «${cat}» setzen`
              : `+ ${sel.length} × Event «${cat}» setzen`;
            submitBtn.disabled = sel.length === 0;
          }
          const counter = document.querySelector("[data-batch-counter]");
          if (counter) counter.textContent = `${sel.length} von ${max} ausgewählt${max >= 200 ? " (max. 200 — Filter verfeinern)" : ""}`;
          const allChecked = max > 0 && preview.every(c => sel.includes(c.id));
          const allCb = document.querySelector("input[data-action='batch-toggle-all']");
          if (allCb) allCb.checked = allChecked;
        };

        // Batch-Event-Auswahl: einzelne Checkbox
        const batchToggle = event.target.closest("[data-action='batch-toggle-contact']");
        if (batchToggle) {
          const cid = Number(batchToggle.dataset.contactId);
          if (!state.modal?.payload?.selected) return;
          const sel = state.modal.payload.selected;
          const idx = sel.indexOf(cid);
          if (idx === -1) sel.push(cid); else sel.splice(idx, 1);
          // data-action ist direkt auf dem <input> — batchToggle IST die Checkbox
          batchToggle.checked = sel.includes(cid);
          batchToggle.closest("tr")?.classList.toggle("bbz-row-ok", sel.includes(cid));
          syncBatchUI(sel);
          return;
        }

        // Batch-Event: Alle/Keine togglen
        const batchToggleAll = event.target.closest("[data-action='batch-toggle-all']");
        if (batchToggleAll && state.modal?.payload) {
          const preview = state.modal.payload.previewContacts || [];
          const allIds = preview.map(c => c.id);
          const allSelected = allIds.every(id => state.modal.payload.selected.includes(id));
          state.modal.payload.selected = allSelected ? [] : [...allIds];
          const newSel = state.modal.payload.selected;
          document.querySelectorAll("input[data-action='batch-toggle-contact']").forEach(cb => {
            cb.checked = newSel.includes(Number(cb.dataset.contactId));
            cb.closest("tr")?.classList.toggle("bbz-row-ok", cb.checked);
          });
          syncBatchUI(newSel);
          return;
        }

        // KEIN separater Handler für [data-modal-submit] nötig:
        // Der Button hat type="submit" und löst den nativen Form-Submit aus,
        // der vom submit-Listener unten abgefangen wird.
        // Ein zusätzlicher dispatchEvent hier würde double-submit verursachen.
      });

      // isPrivat-Label: dynamisch aktualisieren wenn Firma im Kontaktformular wechselt
      document.addEventListener("change", (event) => {
        const firmSelect = event.target.closest("[data-modal-form='contact'] select[name='firmaLookupId']");
        if (firmSelect && state.meta.privateFirmId !== null) {
          const isPrivat = String(firmSelect.value) === String(state.meta.privateFirmId);
          const label = firmSelect.closest(".bbz-modal")?.querySelector("label[data-kommentar-label]");
          if (label) label.textContent = isPrivat ? "Adresse / Notizen (Privatperson — Adresse hier erfassen)" : "Kommentar";
        }
      }, true); // capture: true — vor dem bbz change-listener feuern

      // FIX 2c: Zentraler Form-Submit-Handler — Guard gegen Double-Submit
      document.addEventListener("submit", (event) => {
        const form = event.target.closest("[data-modal-form]");
        if (form) {
          event.preventDefault();
          if (state.meta.loading) return;
          const formType = form.dataset.modalForm;
          if (formType === "firm") {
            controller.handleFirmModalSubmit(form, form.dataset.mode, form.dataset.itemId || null);
          } else if (formType === "history") {
            controller.handleHistoryModalSubmit(form);
          } else if (formType === "task") {
            controller.handleTaskModalSubmit(form);
          } else if (formType === "batch-event") {
            controller.handleBatchEventSubmit(form);
          } else {
            controller.handleModalSubmit(form, form.dataset.mode, form.dataset.itemId || null);
          }
        }
      });

      const debouncedRender = helpers.debounce(() => controller.render(), 150);

      // Browser-Back/Forward — State aus history.state wiederherstellen
      window.addEventListener("popstate", (event) => {
        const s = event.state;
        if (!s) {
          // Kein State (z.B. erster Eintrag) — zur Startseite
          state.filters.route = CONFIG.defaults.route;
          state.selection.firmId = null;
          state.selection.contactId = null;
        } else {
          state.filters.route = s.route || CONFIG.defaults.route;
          state.selection.firmId = s.firmId || null;
          state.selection.contactId = s.contactId || null;
        }
        state.modal = null;
        window.scrollTo(0, 0);
        controller.render();
      });

      // Initialen State setzen damit der erste Back-Schritt korrekt funktioniert
      history.replaceState(
        { route: state.filters.route, firmId: null, contactId: null },
        "",
        `#${state.filters.route}`
      );

      document.addEventListener("input", (event) => {
        const el = event.target;
        if (el.matches("[data-filter='firms-search']")) { state.filters.firms.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='contacts-search']")) { state.filters.contacts.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='planning-search']")) { state.filters.planning.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='history-search']")) { state.filters.history.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='events-search']")) { state.filters.events.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='batch-search']") && state.modal?.payload) { state.modal.payload.filterSearch = el.value; state.modal.payload.selected = []; debouncedRender(); }
        if (el.matches("[data-filter='batch-eventhistory-category-text']") && state.modal?.payload) { state.modal.payload.selectedHistoryCategory = el.value; state.modal.payload.selected = []; debouncedRender(); }
      });

      document.addEventListener("change", (event) => {
        const el = event.target;
        if (el.matches("[data-filter='firms-klassifizierung']")) { state.filters.firms.klassifizierung = el.value; controller.render(); }
        if (el.matches("[data-filter='firms-vip']")) { state.filters.firms.vip = el.value; controller.render(); }
        if (el.matches("[data-filter='firms-sortdir']")) { state.filters.firms.sortDir = el.value; controller.render(); }
        if (el.matches("[data-filter='contacts-archiviert']")) { state.filters.contacts.archiviertAusblenden = el.checked; controller.render(); }
        if (el.matches("[data-filter='planning-open']")) { state.filters.planning.onlyOpen = el.checked; controller.render(); }
        if (el.matches("[data-filter='planning-groupby']")) { state.filters.planning.groupBy = el.value; controller.render(); }
        if (el.matches("[data-filter='planning-sortdir']")) { state.filters.planning.sortDir = el.value; controller.render(); }
        if (el.matches("[data-filter='history-kontaktart']")) { state.filters.history.kontaktart = el.value; controller.render(); }
        if (el.matches("[data-filter='history-leadbbz']")) { state.filters.history.leadbbz = el.value; controller.render(); }
        if (el.matches("[data-filter='history-groupby']")) { state.filters.history.groupBy = el.value; controller.render(); }
        if (el.matches("[data-filter='history-zeitfenster']")) { state.filters.history.zeitfenster = el.value; controller.render(); }
        if (el.matches("[data-filter='events-open']")) { state.filters.events.onlyWithOpenTasks = el.checked; controller.render(); }
        if (el.matches("[data-filter='events-sortby']")) { state.filters.events.sortBy = el.value; controller.render(); }
        // Batch-Event-Modal Filter
        if (el.matches("[data-filter='batch-segment']") && state.modal?.payload) { state.modal.payload.filterSegment = el.value; state.modal.payload.selected = []; controller.render(); }
        if (el.matches("[data-filter='batch-leadbbz']") && state.modal?.payload) { state.modal.payload.filterLeadbbz = el.value; state.modal.payload.selected = []; controller.render(); }
        if (el.matches("[data-filter='batch-sgf']") && state.modal?.payload) { state.modal.payload.filterSgf = el.value; state.modal.payload.selected = []; controller.render(); }
        if (el.matches("[data-filter='batch-eventhistory-category']") && state.modal?.payload) { state.modal.payload.selectedHistoryCategory = el.value; state.modal.payload.selected = []; controller.render(); }
        if (el.matches("[data-action='task-status-change']")) {
          controller.handleTaskStatusChange(Number(el.dataset.taskId), el.value);
        }
      });
    },

    setLoading(isLoading) {
      state.meta.loading = isLoading;
      this.renderShell();
    },

    setMessage(message, type = "info") {
      const el = this.els.globalMessage;
      if (!el) return;
      if (!message) { el.className = "bbz-banner"; el.textContent = ""; return; }
      const cls = { success: "bbz-banner bbz-banner-success show", warning: "bbz-banner bbz-banner-warning show", error: "bbz-banner bbz-banner-error show", info: "bbz-banner bbz-banner-info show" };
      el.className = cls[type] || cls.info;
      el.textContent = message;
    },

    renderShell() {
      // Desktop nav active state
      this.els.navButtons.forEach(btn => {
        btn.classList.toggle("active", btn.dataset.route === state.filters.route);
      });

      // Mobile bottom nav active state — direkt synchronisieren, kein MutationObserver nötig
      document.querySelectorAll(".bbz-bottom-btn").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.route === state.filters.route);
      });

      if (state.auth.isAuthenticated && state.auth.account) {
        this.els.authStatus.innerHTML = `<span class="bbz-auth-dot"></span><span>Angemeldet: ${helpers.escapeHtml(state.auth.account.username || state.auth.account.name || "")}</span>`;
      } else if (state.auth.isReady) {
        this.els.authStatus.innerHTML = `<span class="bbz-auth-dot" style="background:#94a3b8;"></span><span>Nicht angemeldet</span>`;
      } else {
        this.els.authStatus.innerHTML = `<span class="bbz-auth-dot" style="background:#f59e0b;"></span><span>Authentifizierung wird initialisiert ...</span>`;
      }

      if (this.els.btnLogin) {
        this.els.btnLogin.textContent = state.auth.isAuthenticated ? "Angemeldet" : "Anmelden";
        this.els.btnLogin.disabled = state.meta.loading || !state.auth.isReady;
      }
      if (this.els.btnRefresh) {
        this.els.btnRefresh.disabled = state.meta.loading || !state.auth.isReady;
      }
    },

    renderView(html) {
      if (!this.els.viewRoot) return;
      // Fokus + Cursor-Position bei Suchfeldern vor dem Re-Render merken
      const active = document.activeElement;
      const isSearchInput = active && active.matches("[data-filter$='-search']");
      const savedFilter = isSearchInput ? active.dataset.filter : null;
      const savedStart  = isSearchInput ? active.selectionStart : null;
      const savedEnd    = isSearchInput ? active.selectionEnd   : null;

      this.els.viewRoot.innerHTML = html;

      // Fokus + Cursor wiederherstellen
      if (savedFilter) {
        const restored = this.els.viewRoot.querySelector(`[data-filter="${savedFilter}"]`);
        if (restored) {
          restored.focus();
          try { restored.setSelectionRange(savedStart, savedEnd); } catch (_) {}
        }
      }
    },

    loadingBlock(text = "Daten werden geladen ...") {
      return `<section class="bbz-section"><div class="bbz-section-body"><div style="display:flex;align-items:center;gap:10px;"><div class="bbz-loader"></div><div style="font-size:13px;color:var(--muted);">${helpers.escapeHtml(text)}</div></div></div></section>`;
    },

    emptyBlock(text = "Keine Daten vorhanden.", action = null, actionLabel = null) {
      if (action && actionLabel) {
        return `<div class="bbz-empty">${helpers.escapeHtml(text)}<br><button class="bbz-button bbz-button-secondary" style="margin-top:10px;height:32px;font-size:13px;" data-action="${helpers.escapeHtml(action)}">${helpers.escapeHtml(actionLabel)}</button></div>`;
      }
      return `<div class="bbz-empty">${helpers.escapeHtml(text)}</div>`;
    },

    kv(label, value) {
      return `<div class="bbz-kv"><div class="bbz-kv-label">${helpers.escapeHtml(label)}</div><div class="bbz-kv-value">${value || '<span class="bbz-muted">—</span>'}</div></div>`;
    },

    // Wrapper für KV-Gruppen — gibt eine section mit kompakten Rows zurück
    kvSection(title, rows) {
      return `<section class="bbz-section"><div class="bbz-section-header"><div class="bbz-section-title">${helpers.escapeHtml(title)}</div></div><div class="bbz-section-body">${rows.join("")}</div></section>`;
    }
  };

  const api = {
    async initAuth() {
      helpers.ensureMsalAvailable();
      helpers.validateConfig();

      state.auth.isReady = false;
      state.auth.msal = null;

      const msalInstance = new window.msal.PublicClientApplication({
        auth: {
          clientId: CONFIG.graph.clientId,
          authority: CONFIG.graph.authority,
          redirectUri: CONFIG.graph.redirectUri
        },
        cache: { cacheLocation: "localStorage" }
      });

      await msalInstance.initialize();
      state.auth.msal = msalInstance;

      try {
        const redirectResponse = await state.auth.msal.handleRedirectPromise();
        if (redirectResponse?.account) {
          state.auth.account = redirectResponse.account;
          state.auth.isAuthenticated = true;
        }
      } catch (error) {
        console.warn("handleRedirectPromise Fehler:", error);
        state.meta.lastError = error;
        // Nicht fatal — App kann trotzdem mit Cache-Account weitermachen
      }

      // Accounts aus Cache nachladen falls kein Redirect-Response
      if (!state.auth.account) {
        const accounts = state.auth.msal.getAllAccounts();
        if (accounts.length > 0) {
          // Tenant-Match bevorzugen — verhindert falschen Account bei Multi-Tenant-Umgebungen
          state.auth.account = accounts.find(a => a.tenantId === CONFIG.graph.tenantId) || accounts[0];
          state.auth.isAuthenticated = true;
        }
      }

      state.auth.isReady = true;
    },

    async login() {
      if (!state.auth.msal) throw new Error("MSAL ist nicht initialisiert.");

      const loginResponse = await state.auth.msal.loginPopup({
        scopes: CONFIG.graph.scopes,
        prompt: "select_account"
      });

      if (!loginResponse?.account) throw new Error("Keine Kontoinformation aus dem Login erhalten.");

      state.auth.account = loginResponse.account;
      state.auth.isAuthenticated = true;
      await this.acquireToken();
    },

    // FIX 3b: robusteres Token-Handling mit Account-Fallback
    async acquireToken() {
      if (!state.auth.msal) throw new Error("MSAL ist nicht initialisiert.");

      // Account aus Cache nachladen falls leer
      if (!state.auth.account) {
        const accounts = state.auth.msal.getAllAccounts();
        if (accounts.length > 0) {
          // Tenant-Match bevorzugen — verhindert falschen Account bei Multi-Tenant-Umgebungen
          state.auth.account = accounts.find(a => a.tenantId === CONFIG.graph.tenantId) || accounts[0];
          state.auth.isAuthenticated = true;
        } else {
          throw new Error("Kein angemeldetes Konto gefunden.");
        }
      }

      try {
        const tokenResponse = await state.auth.msal.acquireTokenSilent({
          account: state.auth.account,
          scopes: CONFIG.graph.scopes,
          forceRefresh: false
        });
        if (!tokenResponse?.accessToken) throw new Error("Kein Token aus acquireTokenSilent erhalten.");
        state.auth.token = tokenResponse.accessToken;
        return state.auth.token;
      } catch (silentError) {
        console.warn("Silent token fehlgeschlagen, versuche Popup:", silentError);
        const tokenResponse = await state.auth.msal.acquireTokenPopup({
          account: state.auth.account,
          scopes: CONFIG.graph.scopes
        });
        if (!tokenResponse?.accessToken) throw new Error("Kein Token aus acquireTokenPopup erhalten.");
        state.auth.token = tokenResponse.accessToken;
        return state.auth.token;
      }
    },

    async graphRequest(path, options = {}) {
      // Token immer frisch via acquireToken — nicht auf gecachten state.auth.token verlassen
      const token = await this.acquireToken();
      const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
        method: options.method || "GET",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(options.headers || {}) },
        body: options.body ? JSON.stringify(options.body) : undefined
      });

      if (!response.ok) {
        let detail = "";
        try {
          // Vollständigen Body lesen — gibt bei 400 den exakten SP-Feldnamen
          detail = await response.text();
          console.error(`Graph ${response.status} auf ${options.method || "GET"} ${path}:`, detail);
        } catch { detail = response.statusText; }
        throw new Error(`Graph ${response.status}: ${detail}`);
      }

      if (response.status === 204) return null;
      return await response.json();
    },

    async getSiteId() {
      if (state.meta.siteId) return state.meta.siteId;
      const siteRef = `${CONFIG.sharePoint.siteHostname}:${CONFIG.sharePoint.sitePath}`;
      const data = await this.graphRequest(`/sites/${siteRef}`);
      state.meta.siteId = data.id;
      return state.meta.siteId;
    },

    async getListItems(listTitle) {
      const siteId = await this.getSiteId();
      const data = await this.graphRequest(`/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items?expand=fields&top=5000`);
      return data.value || [];
    },

    async loadAll() {
      if (!state.auth.isAuthenticated) throw new Error("Nicht angemeldet — loadAll() abgebrochen.");
      const [firms, contacts, history, tasks] = await Promise.all([
        this.getListItems(SCHEMA.firms.listTitle),
        this.getListItems(SCHEMA.contacts.listTitle),
        this.getListItems(SCHEMA.history.listTitle),
        this.getListItems(SCHEMA.tasks.listTitle)
      ]);

      state.data.firms = firms.map(item => normalizer.firm(item));
      state.data.contacts = contacts.map(item => normalizer.contact(item));
      state.data.history = history.map(item => normalizer.history(item));
      state.data.tasks = tasks.map(item => normalizer.task(item));

      dataModel.enrich();
    },

    // Liest alle Choice-Felder aller relevanten Listen aus SharePoint
    // Schreibt in state.meta.choices[listTitle][spFieldName] = ["Wert1", "Wert2", ...]
    // Wird bei loadAll() und handleRefresh() mitgeladen — SP ist Single Source of Truth
    async loadColumnChoices() {
      const lists = [
        CONFIG.lists.firms,
        CONFIG.lists.contacts,
        CONFIG.lists.history,
        CONFIG.lists.tasks
      ];

      const siteId = await this.getSiteId();

      await Promise.all(lists.map(async (listTitle) => {
        try {
          const data = await this.graphRequest(
            `/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/columns`
          );

          const choicesForList = {};
          for (const col of (data.value || [])) {
            if (col.choice && Array.isArray(col.choice.choices) && col.choice.choices.length > 0) {
              choicesForList[col.name] = col.choice.choices;
            }
          }
          state.meta.choices[listTitle] = choicesForList;
        } catch (err) {
          // Nicht-fatal: Choices bleiben leer, Formular fällt auf Freitext zurück
          console.warn(`loadColumnChoices fehlgeschlagen für ${listTitle}:`, err);
          state.meta.choices[listTitle] = {};
        }
      }));
    },

    // Write-Layer — POST (neues Item anlegen)
    async postItem(listTitle, fields) {
      const siteId = await this.getSiteId();
      return await this.graphRequest(
        `/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items`,
        { method: "POST", body: { fields } }
      );
    },

    // Write-Layer — PATCH (bestehendes Item aktualisieren)
    async patchItem(listTitle, itemId, fields) {
      const siteId = await this.getSiteId();
      return await this.graphRequest(
        `/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items/${itemId}/fields`,
        { method: "PATCH", body: fields }
      );
    },

    // Write-Layer — DELETE
    async deleteItem(listTitle, itemId) {
      const siteId = await this.getSiteId();
      return await this.graphRequest(
        `/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items/${itemId}`,
        { method: "DELETE" }
      );
    }
  };

  const normalizer = {
    getField(item, fieldName) { return item?.fields?.[fieldName]; },
    itemId(item) { return Number(item?.id) || null; },

    firm(item) {
      const f = SCHEMA.firms.fields;
      return {
        id: this.itemId(item),
        title: this.getField(item, f.title) || "",
        adresse: this.getField(item, f.adresse) || "",
        plz: this.getField(item, f.plz) || "",
        ort: this.getField(item, f.ort) || "",
        land: this.getField(item, f.land) || "",
        hauptnummer: this.getField(item, f.hauptnummer) || "",
        klassifizierung: this.getField(item, f.klassifizierung) || "",
        vip: helpers.bool(this.getField(item, f.vip))
      };
    },

    contact(item) {
      const f = SCHEMA.contacts.fields;
      return {
        id: this.itemId(item),
        nachname: this.getField(item, f.nachname) || "",
        vorname: this.getField(item, f.vorname) || "",
        anrede: this.getField(item, f.anrede) || "",
        firmaRaw: this.getField(item, f.firma),
        firmaLookupId: Number(this.getField(item, f.firmaLookupId)) || null,
        funktion: this.getField(item, f.funktion) || "",
        email1: this.getField(item, f.email1) || "",
        email2: this.getField(item, f.email2) || "",
        direktwahl: this.getField(item, f.direktwahl) || "",
        mobile: this.getField(item, f.mobile) || "",
        rolle: this.getField(item, f.rolle) || "",
        leadbbz0: this.getField(item, f.leadbbz0) || "",
        sgf: helpers.normalizeChoiceList(this.getField(item, f.sgf)),
        geburtstag: this.getField(item, f.geburtstag) || "",
        kommentar: this.getField(item, f.kommentar) || "",
        event: helpers.normalizeChoiceList(this.getField(item, f.event)),
        // FIX: eventhistory konsistent als Array normalisieren (wie sgf und event)
        eventhistory: helpers.normalizeChoiceList(this.getField(item, f.eventhistory)),
        archiviert: helpers.bool(this.getField(item, f.archiviert))
      };
    },

    history(item) {
      const f = SCHEMA.history.fields;
      return {
        id: this.itemId(item),
        title: this.getField(item, f.title) || "",
        kontaktRaw: this.getField(item, f.kontakt),
        kontaktLookupId: Number(this.getField(item, f.kontaktLookupId)) || null,
        datum: this.getField(item, f.datum) || "",
        typ: this.getField(item, f.typ) || "",
        notizen: this.getField(item, f.notizen) || "",
        projektbezug: this.getField(item, f.projektbezug) || "",
        leadbbz: this.getField(item, f.leadbbz) || ""
      };
    },

    task(item) {
      const f = SCHEMA.tasks.fields;
      return {
        id: this.itemId(item),
        title: this.getField(item, f.title) || "",
        kontaktRaw: this.getField(item, f.kontakt),
        kontaktLookupId: Number(this.getField(item, f.kontaktLookupId)) || null,
        deadline: this.getField(item, f.deadline) || "",
        status: this.getField(item, f.status) || "",
        leadbbz: this.getField(item, f.leadbbz) || ""
      };
    }
  };

  const dataModel = {
    enrich() {
      const firmById = new Map(state.data.firms.map(f => [f.id, f]));
      const contactById = new Map(state.data.contacts.map(c => [c.id, c]));

      const contacts = state.data.contacts.map(contact => {
        const firm = firmById.get(contact.firmaLookupId) || null;
        return { ...contact, fullName: helpers.fullName(contact), firmId: firm?.id || contact.firmaLookupId || null, firmTitle: firm?.title || contact.firmaRaw || "", firm };
      });

      const history = state.data.history.map(entry => {
        const contact = contactById.get(entry.kontaktLookupId) || null;
        const firm = contact ? firmById.get(contact.firmaLookupId) || null : null;
        return { ...entry, contactId: contact?.id || entry.kontaktLookupId || null, contactName: contact ? helpers.fullName(contact) : (entry.kontaktRaw || ""), firmId: firm?.id || null, firmTitle: firm?.title || "", projektbezugBool: helpers.bool(entry.projektbezug) };
      });

      const tasks = state.data.tasks.map(task => {
        const contact = contactById.get(task.kontaktLookupId) || null;
        const firm = contact ? firmById.get(contact.firmaLookupId) || null : null;
        return { ...task, contactId: contact?.id || task.kontaktLookupId || null, contactName: contact ? helpers.fullName(contact) : (task.kontaktRaw || ""), firmId: firm?.id || null, firmTitle: firm?.title || "", isOpen: helpers.isOpenTask(task.status), isOverdue: helpers.isOverdue(task.deadline) };
      });

      const firms = state.data.firms.map(firm => {
        const firmContacts = contacts.filter(c => c.firmId === firm.id);
        const firmContactIds = new Set(firmContacts.map(c => c.id));
        const firmTasks = tasks.filter(t => firmContactIds.has(t.contactId));
        const firmHistory = history.filter(h => firmContactIds.has(h.contactId));
        const openTasks = firmTasks.filter(t => t.isOpen);
        const nextDeadlineTask = openTasks.filter(t => helpers.toDate(t.deadline)).sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline))[0] || null;
        const latestHistory = [...firmHistory].sort((a, b) => helpers.compareDateDesc(a.datum, b.datum))[0] || null;

        return {
          ...firm,
          contactsCount: firmContacts.length,
          contacts: firmContacts.sort((a, b) => a.fullName.localeCompare(b.fullName, "de")),
          tasks: firmTasks.sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline)),
          history: firmHistory.sort((a, b) => helpers.compareDateDesc(a.datum, b.datum)),
          openTasksCount: openTasks.length,
          nextDeadline: nextDeadlineTask?.deadline || "",
          latestActivity: latestHistory?.datum || ""
        };
      });

      const eventMap = new Map();
      contacts.forEach(contact => {
        const contactTasks = tasks.filter(t => t.contactId === contact.id);
        const contactHistory = history.filter(h => h.contactId === contact.id).sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));
        const latestH = contactHistory[0] || null;
        const openTasks = contactTasks.filter(t => t.isOpen);

        contact.event.forEach(eventName => {
          const key = String(eventName || "").trim();
          if (!key) return;
          if (!eventMap.has(key)) eventMap.set(key, { name: key, contacts: [], contactCount: 0, openTasksCount: 0 });
          eventMap.get(key).contacts.push({
            contactId: contact.id,
            contactName: contact.fullName || contact.nachname,
            firmId: contact.firmId,
            firmTitle: contact.firmTitle,
            rolle: contact.rolle,
            funktion: contact.funktion,
            eventhistory: contact.eventhistory,
            segment: contact.firm ? String(contact.firm.klassifizierung || "").toUpperCase() : "",
            leadbbz: contact.leadbbz0 || "",
            sgf: contact.sgf || [],
            latestHistoryDate: latestH?.datum || "",
            latestHistoryType: latestH?.typ || "",
            latestHistoryText: latestH?.notizen || "",
            openTasksCount: openTasks.length,
            email1: contact.email1
          });
        });
      });

      const eventChoicesOrder = state.meta.choices?.[CONFIG.lists.contacts]?.["Event"] || [];
      const eventOrderIndex = (name) => {
        const idx = eventChoicesOrder.indexOf(name);
        return idx === -1 ? 9999 : idx; // nicht in SP-Choices → ans Ende
      };

      const events = [...eventMap.values()]
        .map(group => ({ ...group, contactCount: group.contacts.length, openTasksCount: group.contacts.reduce((sum, c) => sum + c.openTasksCount, 0), contacts: group.contacts.sort((a, b) => String(a.contactName).localeCompare(String(b.contactName), "de")) }))
        .sort((a, b) => {
          const ia = eventOrderIndex(a.name), ib = eventOrderIndex(b.name);
          if (ia !== ib) return ia - ib;
          return a.name.localeCompare(b.name, "de"); // Fallback alphabetisch
        });

      state.enriched.contacts = contacts.sort((a, b) => a.fullName.localeCompare(b.fullName, "de"));
      state.enriched.history = history.sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));
      state.enriched.tasks = tasks.sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline));
      state.enriched.firms = firms.sort((a, b) => a.title.localeCompare(b.title, "de"));
      state.enriched.events = events;

      // privateFirmId nach jedem enrich() neu auflösen — robust gegen SP-ID-Änderungen
      const privateFirm = state.data.firms.find(
        f => String(f.title).trim() === CONFIG.defaults.privateFirmTitle
      );
      state.meta.privateFirmId = privateFirm?.id || null;
    },

    getFirmById(id) { return state.enriched.firms.find(f => String(f.id) === String(id)) || null; },
    getContactById(id) { return state.enriched.contacts.find(c => String(c.id) === String(id)) || null; }
  };

  const views = {
    // metaType: "" | "alert" | "warn" | "ok"
    kpiBlock(label, value, meta = "", metaType = "") {
      const metaClass = metaType === "alert" ? "bbz-kpi-meta-alert"
        : metaType === "warn" ? "bbz-kpi-meta-warn"
        : metaType === "ok"   ? "bbz-kpi-meta-ok"
        : "bbz-kpi-meta";
      return `<div class="bbz-kpi"><div class="bbz-kpi-label">${helpers.escapeHtml(label)}</div><div class="bbz-kpi-value">${helpers.escapeHtml(String(value))}</div>${meta ? `<div class="${metaClass}">${helpers.escapeHtml(meta)}</div>` : ""}</div>`;
    },

    miniItem(title, meta) {
      return `<div class="bbz-mini-item"><div class="bbz-mini-title">${title}</div><div class="bbz-mini-meta">${meta}</div></div>`;
    },

    renderRoute() {
      if (state.meta.loading) return ui.loadingBlock();

      let viewHtml = "";
      switch (state.filters.route) {
        case "firms": viewHtml = state.selection.firmId ? this.firmDetail() : this.firms(); break;
        case "contacts": viewHtml = state.selection.contactId ? this.contactDetail() : this.contacts(); break;
        case "planning": viewHtml = this.planning(); break;
        case "history": viewHtml = this.historyView(); break;
        case "events": viewHtml = this.events(); break;
        default: viewHtml = this.firms();
      }

      // Modal wird ueber dem View gerendert
      let modalHtml = "";
      if (state.modal?.type === "contact") modalHtml = views.renderContactForm(state.modal.mode, state.modal.payload);
      if (state.modal?.type === "firm")    modalHtml = views.renderFirmForm(state.modal.mode, state.modal.payload?.firmId);
      if (state.modal?.type === "history") modalHtml = views.renderHistoryForm(state.modal.payload);
      if (state.modal?.type === "task")    modalHtml = views.renderTaskForm(state.modal.payload);
      if (state.modal?.type === "batch-event") modalHtml = views.renderBatchEventForm(state.modal.payload);
      return viewHtml + modalHtml;
    },

    // Kontakt-Formular — FIX 1 (toDateInput) integriert, FIX 2 (Modal-Infrastruktur) verdrahtet
    renderContactForm(mode, payload = {}) {
      const itemId = Number(payload.itemId || 0) || null;
      const contact = mode === "edit" ? dataModel.getContactById(itemId) : null;
      const title = mode === "edit" ? "Kontakt bearbeiten" : "Neuer Kontakt";
      const preselectedFirmId = Number(payload.prefillFirmId || contact?.firmId || 0) || "";
      const L = CONFIG.lists.contacts;
      // Privatpersonen-Modus: wenn Firma "Privatpersonen" vorgewählt oder gesetzt
      const isPrivat = state.meta.privateFirmId !== null &&
        (String(preselectedFirmId) === String(state.meta.privateFirmId) ||
         (contact && contact.firmId === state.meta.privateFirmId));

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${title}</div>
              <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
            <form data-modal-form="contact" data-mode="${mode}" data-item-id="${itemId || ""}">
              <div class="bbz-modal-body">
                <div class="bbz-form-grid">

                  <div class="bbz-field">
                    <label>Nachname *</label>
                    <input class="bbz-input" name="nachname" required value="${helpers.escapeHtml(contact?.nachname || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Vorname</label>
                    <input class="bbz-input" name="vorname" value="${helpers.escapeHtml(contact?.vorname || "")}" />
                  </div>

                  <div class="bbz-field">
                    <label>Anrede</label>
                    ${helpers.choiceSelectHtml("anrede", L, "Anrede", contact?.anrede || "")}
                  </div>
                  <div class="bbz-field">
                    <label>Firma *</label>
                    <select class="bbz-select" name="firmaLookupId" required>
                      <option value="">— bitte wählen —</option>
                      ${state.enriched.firms.map(f => `<option value="${f.id}" ${String(preselectedFirmId) === String(f.id) ? "selected" : ""}>${helpers.escapeHtml(f.title)}</option>`).join("")}
                    </select>
                  </div>

                  <div class="bbz-field">
                    <label>Funktion</label>
                    <input class="bbz-input" name="funktion" value="${helpers.escapeHtml(contact?.funktion || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Rolle</label>
                    ${helpers.choiceSelectHtml("rolle", L, "Rolle", contact?.rolle || "")}
                  </div>

                  <div class="bbz-field">
                    <label>Email 1</label>
                    <input class="bbz-input" name="email1" value="${helpers.escapeHtml(contact?.email1 || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Email 2</label>
                    <input class="bbz-input" name="email2" value="${helpers.escapeHtml(contact?.email2 || "")}" />
                  </div>

                  <div class="bbz-field">
                    <label>Direktwahl</label>
                    <input class="bbz-input" name="direktwahl" value="${helpers.escapeHtml(contact?.direktwahl || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Mobile</label>
                    <input class="bbz-input" name="mobile" value="${helpers.escapeHtml(contact?.mobile || "")}" />
                  </div>

                  <div class="bbz-field">
                    <label>Geburtstag</label>
                    <input type="date" class="bbz-input" name="geburtstag" value="${helpers.escapeHtml(helpers.toDateInput(contact?.geburtstag || ""))}" />
                  </div>
                  <div class="bbz-field">
                    <label>Leadbbz</label>
                    ${helpers.choiceSelectHtml("leadbbz0", L, "Leadbbz0", contact?.leadbbz0 || "")}
                  </div>

                  <div class="bbz-field bbz-span-2">
                    <label>SGF <span class="bbz-field-hint">(Mehrfachauswahl)</span></label>
                    ${helpers.choiceMultiHtml("sgf", L, "SGF", contact?.sgf || [])}
                  </div>

                  <div class="bbz-field bbz-span-2">
                    <label>Event <span class="bbz-field-hint">(Mehrfachauswahl)</span></label>
                    ${helpers.choiceMultiHtml("event", L, "Event", contact?.event || [])}
                  </div>

                  <div class="bbz-field bbz-span-2">
                    <label>Eventhistory <span class="bbz-field-hint">(Mehrfachauswahl)</span></label>
                    ${helpers.choiceMultiHtml("eventhistory", L, "Eventhistory", contact?.eventhistory || [])}
                  </div>

                  <div class="bbz-field bbz-span-2">
                    <label data-kommentar-label>${isPrivat ? 'Adresse / Notizen (Privatperson — Adresse hier erfassen)' : 'Kommentar'}</label>
                    <textarea class="bbz-textarea" name="kommentar">${helpers.escapeHtml(contact?.kommentar || "")}</textarea>
                  </div>

                  <label class="bbz-checkbox">
                    <input type="checkbox" name="archiviert" ${contact?.archiviert ? "checked" : ""} />
                    Archiviert
                  </label>

                </div>
              </div>
              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" ${state.meta.loading ? "disabled" : ""}>Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderFirmForm(mode, firmId = null) {
      const firm = mode === "edit" ? dataModel.getFirmById(firmId) : null;
      const title = mode === "edit" ? "Firma bearbeiten" : "Neue Firma";
      const LF = CONFIG.lists.firms;

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${title}</div>
              <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
            <form data-modal-form="firm" data-mode="${mode}" data-item-id="${firmId || ""}">
              <div class="bbz-modal-body">
                <div class="bbz-form-grid">
                  <div class="bbz-field bbz-span-2">
                    <label>Firmenname *</label>
                    <input class="bbz-input" name="title" required value="${helpers.escapeHtml(firm?.title || "")}" />
                  </div>
                  <div class="bbz-field bbz-span-2">
                    <label>Adresse</label>
                    <input class="bbz-input" name="adresse" value="${helpers.escapeHtml(firm?.adresse || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>PLZ</label>
                    <input class="bbz-input" name="plz" value="${helpers.escapeHtml(firm?.plz || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Ort</label>
                    <input class="bbz-input" name="ort" value="${helpers.escapeHtml(firm?.ort || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Land</label>
                    <input class="bbz-input" name="land" value="${helpers.escapeHtml(firm?.land || "Schweiz")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Hauptnummer</label>
                    <input class="bbz-input" name="hauptnummer" value="${helpers.escapeHtml(firm?.hauptnummer || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Klassifizierung</label>
                    ${helpers.choiceSelectHtml("klassifizierung", LF, "Klassifizierung", firm?.klassifizierung || "")}
                  </div>
                  <div class="bbz-field">
                    <label class="bbz-checkbox" style="border:none;padding:0;margin-top:24px;">
                      <input type="checkbox" name="vip" ${firm?.vip ? "checked" : ""} />
                      VIP
                    </label>
                  </div>
                </div>
              </div>
              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" ${state.meta.loading ? "disabled" : ""}>Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderHistoryForm(payload = {}) {
      const mode = payload.mode || "create";
      const itemId = Number(payload.itemId || 0) || null;
      const entry = mode === "edit" ? state.enriched.history.find(h => h.id === itemId) || null : null;
      const prefillContactId = Number(payload.prefillContactId || entry?.contactId || 0) || "";
      const LH = CONFIG.lists.history;
      const title = mode === "edit" ? "Aktivitaet bearbeiten" : "Aktivitaet erfassen";

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${title}</div>
              <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
            <form data-modal-form="history" data-mode="${mode}" data-item-id="${itemId || ""}">
              <div class="bbz-modal-body">
                <div class="bbz-form-grid">
                  <div class="bbz-field">
                    <label>Kontakt *</label>
                    <select class="bbz-select" name="kontaktLookupId" required ${mode === "edit" ? "disabled" : ""}>
                      <option value="">— bitte waehlen —</option>
                      ${state.enriched.contacts.filter(c => !c.archiviert || (entry && c.id === entry.contactId)).map(c => `<option value="${c.id}" ${String(prefillContactId) === String(c.id) ? "selected" : ""}>${helpers.escapeHtml(c.fullName || c.nachname)}${c.firmTitle ? " — " + helpers.escapeHtml(c.firmTitle) : ""}</option>`).join("")}
                    </select>
                    ${mode === "edit" ? `<input type="hidden" name="kontaktLookupId" value="${prefillContactId}" />` : ""}
                  </div>
                  <div class="bbz-field">
                    <label>Datum *</label>
                    <input type="date" class="bbz-input" name="datum" required value="${helpers.toDateInput(entry?.datum || new Date())}" />
                  </div>
                  <div class="bbz-field">
                    <label>Kontaktart</label>
                    ${helpers.choiceSelectHtml("kontaktart", LH, "Kontaktart", entry?.typ || "")}
                  </div>
                  <div class="bbz-field">
                    <label>Leadbbz</label>
                    ${helpers.choiceSelectHtml("leadbbz", LH, "Leadbbz", entry?.leadbbz || "")}
                  </div>
                  <div class="bbz-field bbz-span-2">
                    <label>Projektbezug</label>
                    <label class="bbz-checkbox" style="border:none;padding:0;">
                      <input type="checkbox" name="projektbezug" ${entry?.projektbezugBool ? "checked" : ""} />
                      Ja, mit Projektbezug
                    </label>
                  </div>
                  <div class="bbz-field bbz-span-2">
                    <label>Notizen</label>
                    <textarea class="bbz-textarea" name="notizen" rows="4" placeholder="Was wurde besprochen?">${helpers.escapeHtml(entry?.notizen || "")}</textarea>
                  </div>
                </div>
              </div>
              <div class="bbz-modal-footer">
                <div style="flex:1;">
                  ${mode === "edit" ? `<button type="button" class="bbz-button bbz-button-secondary" style="color:var(--red);border-color:var(--red);" data-action="delete-history" data-id="${itemId}" data-title="${helpers.escapeHtml(entry?.typ || entry?.title || 'Eintrag')}">Löschen</button>` : ""}
                </div>
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" ${state.meta.loading ? "disabled" : ""}>Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderTaskForm(payload = {}) {
      const mode = payload.mode || "create";
      const itemId = Number(payload.itemId || 0) || null;
      const task = mode === "edit" ? state.enriched.tasks.find(t => t.id === itemId) || null : null;
      const prefillContactId = Number(payload.prefillContactId || task?.contactId || 0) || "";
      const LT = CONFIG.lists.tasks;
      const title = mode === "edit" ? "Aufgabe bearbeiten" : "Aufgabe erfassen";

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${title}</div>
              <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
            <form data-modal-form="task" data-mode="${mode}" data-item-id="${itemId || ""}">
              <div class="bbz-modal-body">
                <div class="bbz-form-grid">
                  <div class="bbz-field bbz-span-2">
                    <label>Titel *</label>
                    <input class="bbz-input" name="title" required value="${helpers.escapeHtml(task?.title || "")}" placeholder="Was ist zu tun?" />
                  </div>
                  <div class="bbz-field">
                    <label>Kontakt *</label>
                    <select class="bbz-select" name="kontaktLookupId" required ${mode === "edit" ? "disabled" : ""}>
                      <option value="">— bitte waehlen —</option>
                      ${state.enriched.contacts.filter(c => !c.archiviert || (task && c.id === task.contactId)).map(c => `<option value="${c.id}" ${String(prefillContactId) === String(c.id) ? "selected" : ""}>${helpers.escapeHtml(c.fullName || c.nachname)}${c.firmTitle ? " — " + helpers.escapeHtml(c.firmTitle) : ""}</option>`).join("")}
                    </select>
                    ${mode === "edit" ? `<input type="hidden" name="kontaktLookupId" value="${prefillContactId}" />` : ""}
                  </div>
                  <div class="bbz-field">
                    <label>Deadline</label>
                    <input type="date" class="bbz-input" name="deadline" value="${helpers.toDateInput(task?.deadline || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Status</label>
                    ${helpers.choiceSelectHtml("status", LT, "Status", task?.status || "")}
                  </div>
                  <div class="bbz-field">
                    <label>Leadbbz</label>
                    ${helpers.choiceSelectHtml("leadbbz", LT, "Leadbbz", task?.leadbbz || "")}
                  </div>
                </div>
              </div>
              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" ${state.meta.loading ? "disabled" : ""}>Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderBatchEventForm(payload = {}) {
      const { eventName = "", mode = "anmelden", filterSegment = "", filterLeadbbz = "", filterSgf = "", filterSearch = "", selected = [], selectedHistoryCategory = "" } = payload;
      const LC = CONFIG.lists.contacts;
      const isEventhistory = mode === "eventhistory";

      // SP-Choices für Eventhistory-Feld laden
      const eventhistoryChoices = state.meta.choices?.[CONFIG.lists.contacts]?.["Eventhistory"] || [];
      const eventChoices        = state.meta.choices?.[CONFIG.lists.contacts]?.["Event"] || [];

      // Im Eventhistory-Modus: Kategorie muss erst gewählt werden
      const activeCategory = isEventhistory ? selectedHistoryCategory : eventName;
      const categoryMissing = isEventhistory && !activeCategory;

      const allLeadbbz = [...new Set(state.enriched.contacts.map(c => c.leadbbz0).filter(Boolean))].sort();
      const allSgf     = [...new Set(state.enriched.contacts.flatMap(c => helpers.toArray(c.sgf)))].filter(Boolean).sort();

      // Kandidaten berechnen — nur wenn Kategorie bekannt
      let candidates = [];
      if (!categoryMissing) {
        const existingContactIds = isEventhistory
          ? new Set() // Eventhistory: alle Kontakte wählbar
          : new Set((state.enriched.events.find(g => g.name === activeCategory)?.contacts || []).map(c => c.contactId));

        candidates = isEventhistory
          ? state.enriched.contacts.filter(c => !c.archiviert)
          : state.enriched.contacts.filter(c => !c.archiviert && !existingContactIds.has(c.id));

        if (filterSegment) {
          const firmMap = new Map(state.enriched.firms.map(f => [f.id, f]));
          candidates = candidates.filter(c => String(firmMap.get(c.firmId)?.klassifizierung || "").toUpperCase().startsWith(filterSegment));
        }
        if (filterLeadbbz) candidates = candidates.filter(c => c.leadbbz0 === filterLeadbbz);
        if (filterSgf) candidates = candidates.filter(c => helpers.toArray(c.sgf).includes(filterSgf));
        if (filterSearch.trim()) {
          const s = filterSearch.trim().toLowerCase();
          candidates = candidates.filter(c => [c.fullName, c.firmTitle].some(v => helpers.textIncludes(v, s)));
        }
      }

      const previewContacts = candidates.slice(0, 200);
      const validSelected = categoryMissing ? [] : selected.filter(id => previewContacts.some(c => c.id === id));
      if (state.modal?.payload) {
        state.modal.payload.previewContacts = previewContacts;
        state.modal.payload.selected = validSelected;
      }
      const allChecked = previewContacts.length > 0 && previewContacts.every(c => validSelected.includes(c.id));

      const leadbbzOptions = [`<option value="">— alle Lead BBZ —</option>`, ...allLeadbbz.map(l =>
        `<option value="${helpers.escapeHtml(l)}" ${filterLeadbbz === l ? "selected" : ""}>${helpers.escapeHtml(l)}</option>`)].join("");
      const sgfOptions = [`<option value="">— alle SGF —</option>`, ...allSgf.map(s =>
        `<option value="${helpers.escapeHtml(s)}" ${filterSgf === s ? "selected" : ""}>${helpers.escapeHtml(s)}</option>`)].join("");

      const modeLabel = isEventhistory ? "Eventhistory setzen" : "Event setzen";
      const choicesForDropdown = isEventhistory ? eventhistoryChoices : [];

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal" style="max-width:780px;width:95vw;">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${isEventhistory ? "Eventhistory setzen" : `${helpers.escapeHtml(activeCategory)} — Event setzen`}</div>
              <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
            <form data-modal-form="batch-event" data-event-name="${helpers.escapeHtml(activeCategory)}" data-mode="${mode}">
              <div class="bbz-modal-body">

                ${isEventhistory ? `
                <!-- Schritt 1: Kategorie wählen -->
                <div class="bbz-field" style="margin-bottom:14px;">
                  <label style="font-size:13px;font-weight:500;display:block;margin-bottom:6px;">Eventhistory-Kategorie *</label>
                  ${eventhistoryChoices.length
                    ? `<select class="bbz-select" data-filter="batch-eventhistory-category" style="max-width:360px;">
                        <option value="">— Kategorie wählen —</option>
                        ${eventhistoryChoices.map(c => `<option value="${helpers.escapeHtml(c)}" ${selectedHistoryCategory === c ? "selected" : ""}>${helpers.escapeHtml(c)}</option>`).join("")}
                       </select>`
                    : `<input class="bbz-input" data-filter="batch-eventhistory-category-text" type="text"
                         placeholder="Kategorie eingeben (Choices nicht geladen)" value="${helpers.escapeHtml(selectedHistoryCategory)}" style="max-width:360px;" />`
                  }
                </div>
                ${categoryMissing ? `<div style="font-size:13px;color:var(--muted);padding:12px 0;">Bitte zuerst eine Kategorie wählen.</div>` : ""}
                ` : ""}

                ${!categoryMissing ? `
                <!-- Filterzeile -->
                <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-bottom:12px;">
                  <input class="bbz-input" data-filter="batch-search" type="text" placeholder="Name / Firma ..." value="${helpers.escapeHtml(filterSearch)}" style="font-size:12px;" />
                  <select class="bbz-select" data-filter="batch-segment" style="font-size:12px;">
                    <option value="">— Segment —</option>
                    ${["A","B","C"].map(v=>`<option value="${v}" ${filterSegment===v?"selected":""}>${v}</option>`).join("")}
                  </select>
                  <select class="bbz-select" data-filter="batch-leadbbz" style="font-size:12px;">${leadbbzOptions}</select>
                  <select class="bbz-select" data-filter="batch-sgf" style="font-size:12px;">${sgfOptions}</select>
                </div>

                <!-- Kontakt-Tabelle -->
                <div class="bbz-table-wrap" style="max-height:340px;overflow-y:auto;">
                  <table class="bbz-table" style="min-width:500px;">
                    <thead><tr>
                      <th style="width:32px;">
                        <input type="checkbox" data-action="batch-toggle-all" ${allChecked ? "checked" : ""} title="Alle/Keine" />
                      </th>
                      <th>Kontakt</th>
                      <th>Firma</th>
                      <th>Segment</th>
                      <th>Lead BBZ</th>
                    </tr></thead>
                    <tbody>
                      ${previewContacts.length ? previewContacts.map(c => {
                        const firmObj = state.enriched.firms.find(f => f.id === c.firmId);
                        const seg = String(firmObj?.klassifizierung || "").toUpperCase();
                        const isChecked = validSelected.includes(c.id);
                        return `<tr style="${isChecked ? "background:var(--blue-light);" : ""}">
                          <td><input type="checkbox" data-action="batch-toggle-contact" data-contact-id="${c.id}" ${isChecked ? "checked" : ""} /></td>
                          <td>${helpers.avatarHtml(c)} <span style="margin-left:6px;">${helpers.escapeHtml(c.fullName || c.nachname)}</span></td>
                          <td><span class="bbz-muted" style="font-size:12px;">${helpers.escapeHtml(c.firmTitle || "—")}</span></td>
                          <td>${seg ? `<span class="${helpers.firmBadgeClass(seg)}">${helpers.escapeHtml(seg)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                          <td>${helpers.leadbbzBadgeHtml(c.leadbbz0)}</td>
                        </tr>`;
                      }).join("") : `<tr><td colspan="5"><div class="bbz-empty" style="padding:16px;">Keine Kontakte für diese Filter.</div></td></tr>`}
                    </tbody>
                  </table>
                </div>
                <div data-batch-counter style="font-size:12px;color:var(--muted);margin-top:8px;">
                  ${validSelected.length} von ${previewContacts.length} ausgewählt
                  ${previewContacts.length === 200 ? " (max. 200 — Filter verfeinern)" : ""}
                </div>
                <input type="hidden" name="selectedIds" value="${helpers.escapeHtml(JSON.stringify(validSelected))}" />
                ` : ""}
              </div>
              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary"
                  ${state.meta.loading || validSelected.length === 0 || categoryMissing ? "disabled" : ""}>
                  ${isEventhistory
                    ? `+ ${validSelected.length} × Eventhistory «${helpers.escapeHtml(activeCategory || "?")}» setzen`
                    : `+ ${validSelected.length} × Event «${helpers.escapeHtml(activeCategory)}» setzen`}
                </button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    firms() {
      const filters = state.filters.firms;
      const filteredFirms = state.enriched.firms.filter(firm => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [firm.title, firm.ort, firm.klassifizierung, firm.hauptnummer, firm.adresse, firm.land, ...firm.contacts.map(c => c.fullName)].some(v => helpers.textIncludes(v, search));
        const klassMatch = !filters.klassifizierung || String(firm.klassifizierung || "").toUpperCase().startsWith(filters.klassifizierung.toUpperCase());
        const vipMatch = !filters.vip || (filters.vip === "yes" && firm.vip);
        const privatMatch = !filters.onlyPrivat || (state.meta.privateFirmId !== null && firm.id === state.meta.privateFirmId);
        return searchMatch && klassMatch && vipMatch && privatMatch;
      });
      const firmSortDir = filters.sortDir === "asc" ? 1 : -1;
      const rows = [...filteredFirms].sort((a, b) => {
        if (filters.sortBy === "title")          return a.title.localeCompare(b.title, "de") * firmSortDir;
        if (filters.sortBy === "klassifizierung") return String(a.klassifizierung||"").localeCompare(String(b.klassifizierung||""), "de") * firmSortDir;
        if (filters.sortBy === "vip")            return ((b.vip ? 1 : 0) - (a.vip ? 1 : 0)) * firmSortDir;
        if (filters.sortBy === "openTasksCount") return (a.openTasksCount - b.openTasksCount) * firmSortDir;
        if (filters.sortBy === "latestActivity") return helpers.compareDateDesc(a.latestActivity, b.latestActivity) * -firmSortDir;
        return 0;
      });

      // KPI-Daten
      const aCount = state.enriched.firms.filter(f => String(f.klassifizierung).toUpperCase().includes("A")).length;
      const bCount = state.enriched.firms.filter(f => String(f.klassifizierung).toUpperCase().includes("B")).length;
      const cCount = state.enriched.firms.filter(f => String(f.klassifizierung).toUpperCase().includes("C")).length;
      const allOpenTasks   = state.enriched.tasks.filter(t => t.isOpen);
      const overdueTasks   = allOpenTasks.filter(t => t.isOverdue);

      // Radar-Modus: A/B-Kunden mit Signal, priorisiert sortiert
      const signalPriority = { overdue: 0, never: 1, cold: 2, ok: 3 };
      const today = helpers.todayStart();
      const radarRows = state.enriched.firms.filter(f => {
        const kl = String(f.klassifizierung || "").toUpperCase();
        if (!kl.includes("A") && !kl.includes("B")) return false;
        const sig = helpers.firmSignal(f);
        if (!sig) return false; // "" = keine Info (90-360 Tage, kein Problem)
        const search = filters.search.trim().toLowerCase();
        return !search || helpers.textIncludes(f.title, search);
      }).sort((a, b) => {
        const pa = signalPriority[helpers.firmSignal(a)] ?? 9;
        const pb = signalPriority[helpers.firmSignal(b)] ?? 9;
        if (pa !== pb) return pa - pb;
        return helpers.compareDateAsc(a.latestActivity, b.latestActivity);
      });

      const radarNeverCount   = radarRows.filter(f => helpers.firmSignal(f) === "never").length;
      const radarColdCount    = radarRows.filter(f => helpers.firmSignal(f) === "cold").length;
      const radarOverdueCount = radarRows.filter(f => helpers.firmSignal(f) === "overdue").length;
      // On Track: A/B-Kunden mit letztem Kontakt <90 Tage, keine überfälligen Tasks
      const onTrackCount = state.enriched.firms.filter(f => helpers.firmSignal(f) === "ok").length;

      // Fokus-Bar: überfällige + diese Woche fällig
      const in7     = new Date(today); in7.setDate(in7.getDate() + 7);
      const thisWeek = allOpenTasks.filter(t => { const d = helpers.toDate(t.deadline); return d && d >= today && d <= in7; });      const focusTasks = [...overdueTasks, ...thisWeek.filter(t => !t.isOverdue)]
        .sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline))
        .slice(0, 5);

      // Fokus-Bar HTML — neues Design
      const focusBarHtml = (() => {
        if (focusTasks.length === 0) {
          return `<div class="bbz-focus-bar">
            <div class="bbz-focus-inner">
              <div class="bbz-focus-stat">
                <div class="bbz-focus-stat-label">Dringend heute</div>
                <div class="bbz-focus-number bbz-focus-number-ok">0</div>
                <div class="bbz-focus-stat-sub">alles erledigt</div>
              </div>
              <div class="bbz-focus-divider"></div>
              <div class="bbz-focus-ok">
                <div class="bbz-focus-ok-icon">
                  <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="#22d98a" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 10l4 4 8-8"/></svg>
                </div>
                <div class="bbz-focus-ok-text">Heute geschafft.</div>
                <div class="bbz-focus-ok-sub">Keine offenen Tasks mehr.</div>
              </div>
              <div class="bbz-focus-divider"></div>
              <div class="bbz-focus-cta">
                <a class="bbz-focus-all-link" data-action="navigate-planning">Zur Planung →</a>
              </div>
            </div>
          </div>`;
        }
        const taskItems = focusTasks.map(t => {
          const isOd = t.isOverdue;
          return `<div class="bbz-focus-task" data-action="navigate-planning">
            <div class="${isOd ? "bbz-focus-bar-red" : "bbz-focus-bar-amber"}"></div>
            <span class="bbz-focus-firm">${helpers.escapeHtml(t.firmTitle || t.contactName || "—")}</span>
            <span class="bbz-focus-desc">${helpers.escapeHtml(t.title)}</span>
            <span class="bbz-focus-due ${isOd ? "bbz-focus-due-red" : "bbz-focus-due-amber"}">${isOd ? "überfällig" : helpers.relativeDate(t.deadline)}</span>
          </div>`;
        }).join("");
        const urgentCount = overdueTasks.length;
        const moreCount = Math.max(0, overdueTasks.length + thisWeek.length - focusTasks.length);
        return `<div class="bbz-focus-bar">
          <div class="bbz-focus-inner">
            <div class="bbz-focus-stat">
              <div class="bbz-focus-stat-label">Dringend heute</div>
              <div class="bbz-focus-number${urgentCount > 0 ? " bbz-focus-number-alert" : ""}">${urgentCount > 0 ? urgentCount : focusTasks.length}</div>
              <div class="bbz-focus-stat-sub">${urgentCount > 0 ? "überfällige Tasks" : "Tasks diese Woche"}</div>
            </div>
            <div class="bbz-focus-divider"></div>
            <div class="bbz-focus-tasks">${taskItems}</div>
            <div class="bbz-focus-divider"></div>
            <div class="bbz-focus-cta">
              <a class="bbz-focus-all-link" data-action="navigate-planning">Alle ${allOpenTasks.length} Tasks →</a>
              ${moreCount > 0 ? `<span style="font-size:10px;color:rgba(255,255,255,0.2);">+${moreCount} weitere</span>` : ""}
            </div>
          </div>
        </div>`;
      })();

      return `
        <div>
          ${focusBarHtml}
          <div class="bbz-kpis">
            ${filters.radarMode ? `
            <!-- Radar-Modus KPIs -->
            <div class="bbz-kpi bbz-kpi-red">
              <div class="bbz-kpi-label">Nie kontaktiert</div>
              <div class="bbz-kpi-value bbz-kpi-value-red">${radarNeverCount}</div>
              <div class="bbz-kpi-meta">A-Kunden ohne History</div>
            </div>
            <div class="bbz-kpi bbz-kpi-amber">
              <div class="bbz-kpi-label">Eingeschlafen</div>
              <div class="bbz-kpi-value bbz-kpi-value-amber">${radarColdCount}</div>
              <div class="bbz-kpi-meta">>360 Tage kein Kontakt</div>
            </div>
            <div class="bbz-kpi bbz-kpi-red">
              <div class="bbz-kpi-label">Überfällige Tasks</div>
              <div class="bbz-kpi-value bbz-kpi-value-red">${radarOverdueCount}</div>
              <div class="bbz-kpi-meta">A/B-Kunden betroffen</div>
            </div>
            <div class="bbz-kpi bbz-kpi-green">
              <div class="bbz-kpi-label">On Track ✓</div>
              <div class="bbz-kpi-value bbz-kpi-value-green">${onTrackCount}</div>
              <div class="bbz-kpi-meta bbz-kpi-meta-ok">Kontakt &lt;90 Tage</div>
            </div>
            ` : `
            <!-- Firmen-Kachel: Segment-Filter + orthogonale Zusatzfilter -->
            <div class="bbz-kpi bbz-kpi-blue">
              <div class="bbz-kpi-label">Firmen</div>
              <div class="bbz-kpi-value">${state.enriched.firms.length}</div>
              <!-- Reihe 1: Segment (exklusiv) -->
              <div class="bbz-kpi-chips" style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${["A","B","C"].map(k => {
                  const cnt = state.enriched.firms.filter(f => String(f.klassifizierung||"").toUpperCase().startsWith(k)).length;
                  const active = filters.klassifizierung.toUpperCase().startsWith(k);
                  return `<button class="bbz-kpi-chip ${active?"bbz-kpi-chip-active":""}" data-action="kpi-filter" data-scope="firms-klassifizierung" data-value="${k}">${k} <span>${cnt}</span></button>`;
                }).join("")}
                <button class="bbz-kpi-chip ${!filters.klassifizierung && !filters.vip && !filters.onlyPrivat ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="firms-klassifizierung" data-value="">Alle</button>
              </div>
              <!-- Trennlinie -->
              <div style="border-top:1px solid var(--line-2);margin:6px 0 5px;"></div>
              <!-- Reihe 2: Zusatzfilter (additiv, orthogonal) -->
              <div class="bbz-kpi-chips" style="display:flex;gap:4px;flex-wrap:wrap;">
                <button class="bbz-kpi-chip ${filters.vip === "yes" ? "bbz-kpi-chip-active-gold" : ""}" data-action="kpi-filter" data-scope="firms-vip" data-value="yes">♛ <span>${state.enriched.firms.filter(f=>f.vip).length}</span></button>
                ${state.meta.privateFirmId ? `<button class="bbz-kpi-chip ${filters.onlyPrivat ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="firms-privat" data-value="yes">👤 Privat <span>${state.enriched.contacts.filter(c => c.firmId === state.meta.privateFirmId && !c.archiviert).length}</span></button>` : ""}
              </div>
            </div>
            <!-- Kontakte-Kachel mit History/Tasks/Alle Filter -->
            <div class="bbz-kpi bbz-kpi-blue bbz-kpi-clickable" data-action="kpi-filter" data-scope="navigate" data-value="contacts" style="cursor:pointer;">
              <div class="bbz-kpi-label">Kontakte</div>
              <div class="bbz-kpi-value">${state.enriched.contacts.filter(c => !c.archiviert).length}</div>
              <div class="bbz-kpi-chips" style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                <button class="bbz-kpi-chip" data-action="kpi-filter" data-scope="contacts-mode" data-value="history">History <span>${state.enriched.contacts.filter(c => !c.archiviert && state.enriched.history.some(h => h.contactId === c.id)).length}</span></button>
                <button class="bbz-kpi-chip" data-action="kpi-filter" data-scope="contacts-mode" data-value="tasks">Offene Tasks <span>${state.enriched.contacts.filter(c => !c.archiviert && state.enriched.tasks.some(t => t.contactId === c.id && t.isOpen)).length}</span></button>
                <button class="bbz-kpi-chip" data-action="kpi-filter" data-scope="contacts-mode" data-value="all">Alle</button>
              </div>
            </div>
            <!-- Offene Tasks — drei Zeitzonen als Chips, klickbar zur Planung -->
            <div class="bbz-kpi bbz-kpi-amber bbz-kpi-clickable" data-action="navigate-planning" style="cursor:pointer;">
              <div class="bbz-kpi-label">Offene Tasks</div>
              <div class="bbz-kpi-value${allOpenTasks.filter(t=>t.isOverdue).length > 0 ? ' bbz-kpi-value-amber' : ''}">${allOpenTasks.length}</div>
              <div class="bbz-kpi-chips" style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;pointer-events:none;">
                ${(() => {
                  const in30  = new Date(today); in30.setDate(in30.getDate() + 30);
                  const faellig = allOpenTasks.filter(t => { const d = helpers.toDate(t.deadline); return d ? d <= today : t.isOverdue; }).length;
                  const monat   = allOpenTasks.filter(t => { const d = helpers.toDate(t.deadline); return d && d > today && d <= in30; }).length;
                  const uebrige = allOpenTasks.length - faellig - monat;
                  return `<span class="bbz-kpi-chip" style="background:var(--red-soft);border-color:#f0b0b2;color:var(--red);">Fällig <span style="color:var(--red);">${faellig}</span></span>`
                       + `<span class="bbz-kpi-chip" style="background:#fff9eb;border-color:#f4dfab;color:var(--amber);">Monat <span style="color:var(--amber);">${monat}</span></span>`
                       + `<span class="bbz-kpi-chip">Übrige <span>${uebrige}</span></span>`;
                })()}
              </div>
            </div>`}
            <!-- Events-Kachel — immer sichtbar -->
            <div class="bbz-kpi bbz-kpi-blue bbz-kpi-clickable" data-action="kpi-filter" data-scope="navigate" data-value="events" style="cursor:pointer;">
              <div class="bbz-kpi-label">Events</div>
              <div class="bbz-kpi-value">${state.enriched.events.length}</div>
              <div class="bbz-kpi-chips" style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;pointer-events:none;">
                ${state.enriched.events.slice(0, 3).map(e =>
                  `<span class="bbz-kpi-chip">${helpers.escapeHtml(e.name)} <span>${e.contactCount}</span></span>`
                ).join("")}
                ${state.enriched.events.length > 3
                  ? `<span class="bbz-kpi-chip bbz-muted">+${state.enriched.events.length - 3}</span>`
                  : ""}
              </div>
            </div>
          </div>
          <div class="bbz-grid bbz-grid-70-30">
            <section class="bbz-section">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">${filters.radarMode ? "Pflege A/B" : "Firmen-Cockpit"}</div>
                  <div class="bbz-section-subtitle">${filters.radarMode ? `${radarRows.filter(f => helpers.firmSignal(f) !== "ok").length} mit Handlungsbedarf · ${onTrackCount} On Track ✓` : "Segment, Tasks und Fristen auf einen Blick"}</div>
                </div>
                <div style="display:flex;gap:8px;align-items:center;">
                  <!-- Tab-Bar -->
                  <div style="display:flex;border:1px solid var(--line);border-radius:9px;overflow:hidden;background:var(--panel-2);">
                    <button class="bbz-button" style="height:32px;font-size:12px;border:none;border-radius:0;${!filters.radarMode ? "background:var(--panel);color:var(--text);font-weight:700;" : "background:none;color:var(--muted);"}"
                      data-action="kpi-filter" data-scope="firms-radar" ${!filters.radarMode ? "disabled" : ""}>
                      Alle Firmen
                    </button>
                    <button class="bbz-button" style="height:32px;font-size:12px;border:none;border-radius:0;${filters.radarMode ? "background:var(--panel);color:var(--text);font-weight:700;" : "background:none;color:var(--muted);"}"
                      data-action="kpi-filter" data-scope="firms-radar" ${filters.radarMode ? "disabled" : ""}>
                      Pflege A/B ${radarRows.length > 0 ? `<span style="background:var(--red-light);color:var(--red);border-radius:999px;padding:1px 6px;font-size:11px;margin-left:4px;">${radarRows.length}</span>` : ""}
                    </button>
                  </div>
                  ${!filters.radarMode ? `<button class="bbz-button bbz-button-primary" data-action="open-firm-form">+ Firma</button>` : ""}
                </div>
              </div>
              <div class="bbz-section-body">
                <div style="margin-bottom:10px;">
                  <input class="bbz-input" style="width:100%;height:40px;font-size:14px;" data-filter="firms-search" type="text"
                    placeholder="${filters.radarMode ? "Suche nach Firma ..." : "Suche nach Firma, Ort, Ansprechpartner ..."}"
                    value="${helpers.escapeHtml(filters.search)}" />
                </div>
                ${filters.radarMode ? `
                <div class="bbz-table-wrap">
                  <table class="bbz-table">
                    <thead><tr>
                      <th></th>
                      <th>Firma</th>
                      <th>Klassifizierung</th>
                      <th>Pflege-Grund</th>
                      <th>Letzte Aktivität</th>
                      <th>Nächste Deadline</th>
                      <th>Kontakte</th>
                    </tr></thead>
                    <tbody>
                      ${radarRows.length ? radarRows.map(firm => {
                        const sig = helpers.firmSignal(firm);
                        const signalDot = sig === "overdue"
                          ? `<span class="bbz-signal bbz-signal-red" title="Überfällige Tasks"></span>`
                          : sig === "ok"
                          ? `<span class="bbz-signal bbz-signal-green" title="On Track — letzter Kontakt < 90 Tage"></span>`
                          : `<span class="bbz-signal bbz-signal-amber"></span>`;
                        const rowClass = sig === "overdue" ? "bbz-row-alert"
                          : sig === "ok" ? "bbz-row-ok"
                          : "bbz-row-cold";
                        const lastDate = helpers.toDate(firm.latestActivity);
                        const months = lastDate
                          ? (today.getFullYear() - lastDate.getFullYear()) * 12 + (today.getMonth() - lastDate.getMonth())
                          : null;
                        const grundHtml = sig === "never"
                          ? `<span style="color:var(--red);font-weight:600;">🔴 Nie kontaktiert</span>`
                          : sig === "cold"
                          ? `<span style="color:var(--amber);font-weight:600;">🟡 Seit ${months} Monat${months !== 1 ? "en" : ""} still</span>`
                          : sig === "ok"
                          ? `<span style="color:var(--green);font-weight:600;">✅ On Track — vor ${months !== null ? months + " Monat" + (months !== 1 ? "en" : "") : "kurzem"}</span>`
                          : `<span style="color:var(--muted);font-weight:600;">⚠️ ${firm.tasks.filter(t => t.isOpen && t.isOverdue).length} Task${firm.tasks.filter(t => t.isOpen && t.isOverdue).length !== 1 ? "s" : ""} überfällig</span>`;
                        return `
                          <tr class="${rowClass}">
                            <td style="width:28px;padding-right:4px;">${signalDot}</td>
                            <td><a class="bbz-link" data-action="open-firm" data-id="${firm.id}">${helpers.escapeHtml(firm.title)}</a></td>
                            <td>${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                            <td>${grundHtml}</td>
                            <td>${firm.latestActivity ? `<span title="${helpers.formatDate(firm.latestActivity)}">${helpers.relativeDate(firm.latestActivity)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                            <td class="${firm.nextDeadline && helpers.isOverdue(firm.nextDeadline) ? "bbz-danger" : ""}">${firm.nextDeadline ? helpers.relativeDate(firm.nextDeadline) : '<span class="bbz-muted">—</span>'}</td>
                            <td>${firm.contactsCount > 0 ? firm.contactsCount : `<span style="color:var(--red);">${firm.contactsCount}</span>`}</td>
                          </tr>`;
                      }).join("") : `<tr><td colspan="7">${ui.emptyBlock("Keine Pflege-Fälle gefunden.")}</td></tr>`}
                    </tbody>
                  </table>
                </div>` : `
                <div class="bbz-table-wrap">
                  <table class="bbz-table">
                    <thead><tr>
                      ${(()=>{
                        const firmSortTh = (label, col) => {
                          const active = filters.sortBy === col;
                          const icon = active ? (filters.sortDir === "asc" ? " ↑" : " ↓") : "";
                          return `<th style="cursor:pointer;user-select:none;${active?"color:var(--blue);":""}" data-action="set-sort" data-col="${col}" data-scope="firms">${label}${icon}</th>`;
                        };
                        return "<th></th>"
                          + firmSortTh("Firma","title")
                          + "<th>Ort</th>"
                          + firmSortTh("Klassifizierung","klassifizierung")
                          + firmSortTh("VIP","vip")
                          + "<th>Kontakte</th>"
                          + firmSortTh("Tasks","openTasksCount")
                          + "<th>Nächste Deadline</th>"
                          + firmSortTh("Letzte Aktivität","latestActivity");
                      })()}
                    </tr></thead>
                    <tbody>
                      ${rows.length ? rows.map(firm => {
                        const signal = helpers.firmSignal(firm);
                        const signalDot = signal === "overdue"
                          ? `<span class="bbz-signal bbz-signal-red" title="Überfällige Tasks"></span>`
                          : signal === "never"
                          ? `<span class="bbz-signal bbz-signal-amber" title="A-Kunde — noch kein Kontakt erfasst"></span>`
                          : signal === "cold"
                          ? `<span class="bbz-signal bbz-signal-amber" title="Kein Kontakt seit über 360 Tagen (A/B-Kunde)"></span>`
                          : signal === "ok"
                          ? `<span class="bbz-signal bbz-signal-green" title="On Track — letzter Kontakt < 90 Tage"></span>`
                          : `<span class="bbz-signal bbz-signal-none"></span>`;
                        const rowClass = signal === "overdue" ? "bbz-row-alert"
                          : (signal === "cold" || signal === "never") ? "bbz-row-cold"
                          : signal === "ok" ? "bbz-row-ok"
                          : "";
                        return `
                        <tr class="${rowClass}">
                          <td style="width:28px;padding-right:4px;">${signalDot}</td>
                          <td><a class="bbz-link" data-action="open-firm" data-id="${firm.id}">${helpers.escapeHtml(firm.title)}</a><div class="bbz-subtext">${helpers.escapeHtml(firm.hauptnummer || "—")}</div></td>
                          <td>${helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ")) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                          <td>${firm.vip ? '<span class="bbz-pill bbz-pill-vip">♛</span>' : '<span class="bbz-muted">—</span>'}</td>
                          <td>${firm.contactsCount}</td>
                          <td>${firm.openTasksCount > 0 ? `<span class="${overdueTasks.some(t => t.firmId === firm.id) ? "bbz-danger" : ""}">${firm.openTasksCount}</span>` : '<span class="bbz-muted">—</span>'}</td>
                          <td class="${firm.nextDeadline && helpers.isOverdue(firm.nextDeadline) ? "bbz-danger" : ""}">${firm.nextDeadline ? helpers.relativeDate(firm.nextDeadline) : '<span class="bbz-muted">—</span>'}</td>
                          <td>${firm.latestActivity ? `<span title="${helpers.formatDate(firm.latestActivity)}">${helpers.relativeDate(firm.latestActivity)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                        </tr>`; }).join("") : `<tr><td colspan="9">${ui.emptyBlock("Keine Firmen für die aktuelle Filterung gefunden.")}</td></tr>`}
                    </tbody>
                  </table>
                </div>`}
                <!-- Mobile Card List (nur sichtbar auf kleinen Screens via CSS) -->
                <div class="bbz-card-list bbz-mobile-only">
                  ${rows.length ? rows.map(firm => {
                    const signal = helpers.firmSignal(firm);
                    const sigDot = signal === "overdue"
                      ? `<span class="bbz-signal bbz-signal-red"></span>`
                      : (signal === "never" || signal === "cold")
                      ? `<span class="bbz-signal bbz-signal-amber"></span>`
                      : signal === "ok"
                      ? `<span class="bbz-signal bbz-signal-green"></span>`
                      : `<span style="width:8px;flex-shrink:0;display:inline-block;"></span>`;
                    const taskBadge = firm.openTasksCount > 0
                      ? overdueTasks.some(t => t.firmId === firm.id)
                        ? `<span class="bbz-status-chip bbz-status-overdue">${firm.openTasksCount} überfällig</span>`
                        : `<span class="bbz-status-chip bbz-status-open">${firm.openTasksCount} offen</span>`
                      : "";
                    return `<div class="bbz-list-card" data-action="open-firm" data-id="${firm.id}">
                      ${sigDot}
                      <div class="bbz-list-card-body">
                        <div class="bbz-list-card-title">${helpers.escapeHtml(firm.title)}</div>
                        <div class="bbz-list-card-sub">${helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ") || "")}${firm.latestActivity ? " · " + helpers.relativeDate(firm.latestActivity) : ""}</div>
                      </div>
                      <div class="bbz-list-card-right">
                        ${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : ""}
                        ${taskBadge}
                      </div>
                    </div>`;
                  }).join("") : ui.emptyBlock("Keine Firmen gefunden.")}
                </div>
              </div>
            </section>
            <div class="bbz-cockpit-stack">
              <section class="bbz-section">
                <div class="bbz-section-header">
                  <div><div class="bbz-section-title">Offene Tasks</div><div class="bbz-section-subtitle">Fälligkeit nach Firma</div></div>
                  <a class="bbz-link" style="font-size:12px;" data-action="navigate-planning">Alle →</a>
                </div>
                <div class="bbz-section-body">
                  ${(() => {
                    const in30  = new Date(today); in30.setDate(in30.getDate() + 30);

                    // Alle Firmen mit offenen Tasks, nach nächster Deadline sortiert
                    const firmsWithTasks = [...state.enriched.firms]
                      .filter(f => f.openTasksCount > 0)
                      .sort((a, b) => helpers.compareDateAsc(a.nextDeadline, b.nextDeadline));

                    // Zone 1: Fällig — Deadline ≤ heute (inkl. keine Deadline wenn überfällige Tasks)
                    const zoneFaellig = firmsWithTasks.filter(f => {
                      const d = helpers.toDate(f.nextDeadline);
                      return d ? d <= today : f.tasks.some(t => t.isOpen && t.isOverdue);
                    });

                    // Zone 2: Dieser Monat — Deadline 1–30 Tage
                    const zoneMonat = firmsWithTasks.filter(f => {
                      const d = helpers.toDate(f.nextDeadline);
                      return d && d > today && d <= in30;
                    });

                    // Zone 3: Übrige — Deadline > 30 Tage oder keine Deadline
                    const zoneUebrige = firmsWithTasks.filter(f => {
                      const d = helpers.toDate(f.nextDeadline);
                      return !d ? !f.tasks.some(t => t.isOpen && t.isOverdue) : d > in30;
                    });

                    const zoneHtml = (label, color, firms, emptyText) => {
                      if (firms.length === 0 && emptyText === null) return "";
                      return `
                        <div class="bbz-zone" style="margin-bottom:12px;">
                          <div class="bbz-zone-label" style="color:${color};font-size:11px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;margin-bottom:6px;">
                            ${label} <span style="font-weight:400;opacity:.7;">(${firms.length})</span>
                          </div>
                          ${firms.length ? `<div class="bbz-mini-list">${firms.slice(0, 6).map(f => {
                            const d = helpers.toDate(f.nextDeadline);
                            const meta = d
                              ? (d <= today
                                  ? `${f.openTasksCount} Task${f.openTasksCount !== 1 ? "s" : ""} · <span style="color:var(--red);font-weight:600;">${helpers.relativeDate(f.nextDeadline)}</span>`
                                  : `${f.openTasksCount} Task${f.openTasksCount !== 1 ? "s" : ""} · ${helpers.relativeDate(f.nextDeadline)}`)
                              : `${f.openTasksCount} Task${f.openTasksCount !== 1 ? "s" : ""}`;
                            return `<div class="bbz-mini-item" style="display:flex;align-items:center;justify-content:space-between;gap:8px;">
                              <a class="bbz-link bbz-mini-title" data-action="navigate-planning">${helpers.escapeHtml(f.title)}</a>
                              <span class="bbz-mini-meta" style="white-space:nowrap;flex-shrink:0;">${meta}</span>
                            </div>`;
                          }).join("")}${firms.length > 6 ? `<div style="font-size:12px;color:var(--muted);padding:4px 0;">+${firms.length - 6} weitere</div>` : ""}</div>`
                          : `<div style="font-size:12px;color:var(--muted);font-style:italic;">—</div>`}
                        </div>`;
                    };

                    const hasAny = zoneFaellig.length + zoneMonat.length + zoneUebrige.length > 0;
                    if (!hasAny) return ui.emptyBlock("Keine offenen Tasks vorhanden.");

                    return zoneHtml("Fällig", "var(--red)", zoneFaellig, "")
                         + zoneHtml("Dieser Monat", "var(--amber)", zoneMonat, "")
                         + zoneHtml("Übrige", "var(--muted)", zoneUebrige, "");
                  })()}
                </div>
              </section>
            </div>
          </div>
        </div>
      `;
    },

    firmDetail() {
      const firm = dataModel.getFirmById(state.selection.firmId);
      if (!firm) return ui.emptyBlock("Die ausgewaehlte Firma wurde nicht gefunden.");
      const recentHistory = [...firm.history].slice(0, 20);
      const bandClass = helpers.detailBandClass(firm);
      return `
        <div>
          <div class="${bandClass}">
            <div class="bbz-detail-header" style="margin-bottom:0;">
              <div>
                <button class="bbz-button bbz-button-secondary" style="margin-bottom:12px;background:rgba(255,255,255,0.7);" data-action="back-to-firms">← Firmenliste</button>
                <div class="bbz-detail-title">${helpers.escapeHtml(firm.title)}</div>
                <div class="bbz-detail-subtitle">${helpers.escapeHtml(helpers.joinNonEmpty([firm.adresse, helpers.joinNonEmpty([firm.plz, firm.ort], " "), firm.land], " · ")) || "Keine Adresse erfasst"}</div>
                <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:12px;">
                  ${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : ""}
                  ${firm.vip ? `<span class="bbz-pill bbz-pill-vip">♛</span>` : ""}
                </div>
              </div>
              <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                <button class="bbz-button bbz-button-secondary" style="${firm.contactsCount > 0 ? "opacity:0.4;cursor:not-allowed;" : "color:var(--red);border-color:var(--red);"}" data-action="delete-firm" data-id="${firm.id}" data-name="${helpers.escapeHtml(firm.title)}" data-contacts="${firm.contactsCount}">Löschen</button>
                <button class="bbz-button bbz-button-secondary" data-action="open-firm-form" data-id="${firm.id}">Bearbeiten</button>
                <button class="bbz-button bbz-button-secondary" data-action="open-task-form" data-firm-id="${firm.id}">+ Task</button>
                <button class="bbz-button bbz-button-secondary" data-action="open-history-form" data-firm-id="${firm.id}">+ Aktivität</button>
                <button class="bbz-button bbz-button-primary" data-action="open-contact-form" data-firm-id="${firm.id}">+ Kontakt</button>
              </div>
            </div>
          </div>
          <div class="bbz-kpis" style="margin-top:10px;">
            ${this.kpiBlock("Kontakte", firm.contactsCount)}
            ${this.kpiBlock("Offene Tasks", firm.openTasksCount, firm.tasks.some(t => t.isOpen && t.isOverdue) ? "überfällig" : firm.openTasksCount > 0 ? "offen" : "keine offen", firm.tasks.some(t => t.isOpen && t.isOverdue) ? "alert" : "")}
            ${this.kpiBlock("Nächste Deadline", firm.nextDeadline ? helpers.relativeDate(firm.nextDeadline) : "—", firm.nextDeadline && helpers.isOverdue(firm.nextDeadline) ? "überfällig" : "", firm.nextDeadline && helpers.isOverdue(firm.nextDeadline) ? "alert" : "")}
            ${this.kpiBlock("Aktivitäten", firm.history.length, firm.latestActivity ? helpers.relativeDate(firm.latestActivity) : "noch keine")}
          </div>
          <div class="bbz-grid bbz-grid-3">
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Stammdaten</div></div>
              <div class="bbz-section-body">
                ${ui.kv("Klassifizierung", firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("VIP", firm.vip ? '<span class="bbz-pill bbz-pill-vip">♛</span>' : '<span class="bbz-muted">Nein</span>')}
                ${ui.kv("Adresse", helpers.escapeHtml(firm.adresse) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("PLZ / Ort", helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ")) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Land", helpers.escapeHtml(firm.land) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Hauptnummer", helpers.escapeHtml(firm.hauptnummer) || '<span class="bbz-muted">—</span>')}
              </div>
            </section>
            <section class="bbz-section" style="grid-column: span 2;">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Kontakte</div><div class="bbz-section-subtitle">Ansprechpartner dieser Firma</div></div></div>
              <div class="bbz-section-body">
                <!-- Desktop: Tabelle -->
                <div class="bbz-table-wrap bbz-desktop-only">
                  <table class="bbz-table">
                    <thead><tr><th></th><th>Name</th><th>Funktion</th><th>Rolle</th><th>E-Mail</th><th>Telefon</th></tr></thead>
                    <tbody>
                      ${firm.contacts.length ? firm.contacts.map(c => `
                        <tr>
                          <td style="width:36px;padding-right:0;">${helpers.avatarHtml(c)}</td>
                          <td><a class="bbz-link" data-action="open-contact" data-id="${c.id}">${helpers.escapeHtml(c.fullName || c.nachname)}</a>${c.archiviert ? ' <span class="bbz-muted" style="font-size:11px;">(archiviert)</span>' : ""}</td>
                          <td>${helpers.escapeHtml(c.funktion) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${helpers.escapeHtml(c.rolle) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${c.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(c.email1)}">${helpers.escapeHtml(c.email1)}</a>` : '<span class="bbz-muted">—</span>'}</td>
                          <td>${helpers.escapeHtml(helpers.joinNonEmpty([c.direktwahl, c.mobile], " / ")) || '<span class="bbz-muted">—</span>'}</td>
                        </tr>`).join("") : `<tr><td colspan="6">${ui.emptyBlock("Keine Kontakte vorhanden.", "open-contact-form", "+ Ersten Kontakt hinzufügen")}</td></tr>`}
                    </tbody>
                  </table>
                </div>
                <!-- Mobile: Cards -->
                <div class="bbz-card-list bbz-mobile-only">
                  ${firm.contacts.length ? firm.contacts.map(c => `
                    <div class="bbz-list-card" data-action="open-contact" data-id="${c.id}">
                      ${helpers.avatarHtml(c)}
                      <div class="bbz-list-card-body">
                        <div class="bbz-list-card-title">${helpers.escapeHtml(c.fullName || c.nachname)}${c.archiviert ? ' <span class="bbz-muted" style="font-size:10px;">(archiviert)</span>' : ""}</div>
                        <div class="bbz-list-card-sub">${helpers.escapeHtml(helpers.joinNonEmpty([c.funktion, c.rolle], " · ")) || "—"}</div>
                      </div>
                      <div class="bbz-list-card-right">
                        ${c.email1 ? `<span style="font-size:10px;color:var(--subtle);">${helpers.escapeHtml(c.email1)}</span>` : ""}
                      </div>
                    </div>`).join("") : ui.emptyBlock("Keine Kontakte vorhanden.", "open-contact-form", "+ Ersten Kontakt hinzufügen")}
                </div>
              </div>
            </section>
          </div>
          <div class="bbz-grid bbz-grid-2" style="margin-top:12px;">
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Aktivitäten</div><div class="bbz-section-subtitle">Aggregiert über alle Kontakte</div></div>
                <button class="bbz-button bbz-button-secondary" style="height:32px;font-size:13px;" data-action="open-history-form" data-firm-id="${firm.id}">+ Aktivität</button>
              </div>
              <div class="bbz-section-body">
                ${recentHistory.length ? `<div class="bbz-timeline">${recentHistory.map(h => `
                  <div class="bbz-timeline-item">
                    <div class="bbz-timeline-date">${helpers.relativeDate(h.datum) || "—"}<br><span class="bbz-muted" style="font-size:11px;">${helpers.formatDate(h.datum)}</span><br><span class="bbz-muted">${helpers.escapeHtml(h.contactName || "")}</span></div>
                    <div>
                      <div class="bbz-timeline-title">${helpers.escapeHtml(h.typ || h.title || "Eintrag")} ${h.projektbezugBool ? '<span class="bbz-chip" style="background:var(--blue-light);color:var(--blue);border-color:#a8c8e0;">Projektbezug</span>' : '<span class="bbz-chip">Allgemein</span>'}</div>
                      <div class="bbz-timeline-text">${helpers.escapeHtml(h.notizen || "—")}</div>
                      <div style="margin-top:6px;display:flex;gap:6px;">
                        <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;" data-action="edit-history" data-id="${h.id}">Bearbeiten</button>
                        <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;color:var(--red);border-color:var(--red);" data-action="delete-history" data-id="${h.id}" data-title="${helpers.escapeHtml(h.typ || h.title || 'Eintrag')}">Löschen</button>
                      </div>
                    </div>
                  </div>`).join("")}</div>`
                  : `<div class="bbz-empty">Noch keine Aktivitäten erfasst.<br><button class="bbz-button bbz-button-secondary" style="margin-top:10px;height:32px;font-size:13px;" data-action="open-history-form" data-firm-id="${firm.id}">+ Erste Aktivität erfassen</button></div>`}
              </div>
            </section>
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Aufgaben</div></div>
                <button class="bbz-button bbz-button-secondary" style="height:32px;font-size:13px;" data-action="open-task-form" data-firm-id="${firm.id}">+ Task</button>
              </div>
              <div class="bbz-section-body">
                <div class="bbz-table-wrap">
                  <table class="bbz-table">
                    <thead><tr><th>Titel</th><th>Deadline</th><th>Status</th><th>Kontakt</th><th>Aktionen</th></tr></thead>
                    <tbody>
                      ${firm.tasks.length ? firm.tasks.map(t => `
                        <tr>
                          <td>${helpers.escapeHtml(t.title) || '<span class="bbz-muted">—</span>'}</td>
                          <td class="${helpers.isOpenTask(t.status) && helpers.isOverdue(t.deadline) ? "bbz-danger" : ""}">${t.deadline ? helpers.relativeDate(t.deadline) : '<span class="bbz-muted">—</span>'}</td>
                          <td>${helpers.statusChipHtml(t.status, t.deadline)}</td>
                          <td>${t.contactId ? `<a class="bbz-link" data-action="open-contact" data-id="${t.contactId}">${helpers.escapeHtml(t.contactName || "Kontakt")}</a>` : helpers.escapeHtml(t.contactName || "—")}</td>
                          <td style="white-space:nowrap;">
                            <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 8px;margin-right:3px;" data-action="edit-task" data-id="${t.id}">Bearbeiten</button>
                            <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 8px;color:var(--red);border-color:var(--red);" data-action="delete-task" data-id="${t.id}" data-title="${helpers.escapeHtml(t.title)}">Löschen</button>
                          </td>
                        </tr>`).join("") : `<tr><td colspan="5">${ui.emptyBlock("Keine Aufgaben vorhanden.")}</td></tr>`}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          </div>
        </div>
      `;
    },

    contacts() {
      const filters = state.filters.contacts;
      const kpiMode = filters._kpiMode || "all";

      // Contacts mit History / offenen Tasks für KPI-Counts
      const contactsWithHistory = state.enriched.contacts.filter(c => !c.archiviert && state.enriched.history.some(h => h.contactId === c.id));
      const contactsWithOpenTasks = state.enriched.contacts.filter(c => !c.archiviert && state.enriched.tasks.some(t => t.contactId === c.id && t.isOpen));

      const filteredContacts = state.enriched.contacts.filter(c => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [c.fullName, c.firmTitle, c.funktion, c.rolle, c.email1, c.email2, c.direktwahl, c.mobile, c.kommentar, ...c.sgf, ...c.event].some(v => helpers.textIncludes(v, search));
        const archivMatch = !filters.archiviertAusblenden || !c.archiviert;
        const modeMatch = kpiMode === "history"
          ? state.enriched.history.some(h => h.contactId === c.id)
          : kpiMode === "tasks"
          ? state.enriched.tasks.some(t => t.contactId === c.id && t.isOpen)
          : true;
        return searchMatch && archivMatch && modeMatch;
      });
      const cSortDir = filters.sortDir === "asc" ? 1 : -1;
      const rows = [...filteredContacts].sort((a, b) => {
        if (filters.sortBy === "fullName")  return String(a.fullName||"").localeCompare(String(b.fullName||""), "de") * cSortDir;
        if (filters.sortBy === "firmTitle") return String(a.firmTitle||"").localeCompare(String(b.firmTitle||""), "de") * cSortDir;
        if (filters.sortBy === "rolle")     return String(a.rolle||"").localeCompare(String(b.rolle||""), "de") * cSortDir;
        if (filters.sortBy === "leadbbz0")  return String(a.leadbbz0||"").localeCompare(String(b.leadbbz0||""), "de") * cSortDir;
        return 0;
      });
      const cTh = (label, col) => {
        const active = filters.sortBy === col;
        const icon = active ? (filters.sortDir === "asc" ? " ↑" : " ↓") : "";
        return `<th style="cursor:pointer;user-select:none;${active?"color:var(--blue);":""}" data-action="set-sort" data-col="${col}" data-scope="contacts">${label}${icon}</th>`;
      };

      // Counts für KPI-Chips
      const totalActive   = state.enriched.contacts.filter(c => !c.archiviert).length;
      const withHistory   = state.enriched.contacts.filter(c => !c.archiviert && state.enriched.history.some(h => h.contactId === c.id)).length;
      const withOpenTasks = state.enriched.contacts.filter(c => !c.archiviert && state.enriched.tasks.some(t => t.contactId === c.id && t.isOpen)).length;
      const allOpenTasks  = state.enriched.tasks.filter(t => t.isOpen).length;
      const overdueTasks  = state.enriched.tasks.filter(t => t.isOpen && t.isOverdue).length;

      return `
        <div>
          <div class="bbz-kpis">
            <!-- Kontakte mit Schnellfilter -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Kontakte</div>
              <div class="bbz-kpi-value">${totalActive}</div>
              <div class="bbz-kpi-chips" style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                <button class="bbz-kpi-chip ${kpiMode==="history"?"bbz-kpi-chip-active":""}" data-action="kpi-filter" data-scope="contacts-mode" data-value="history">Mit History <span>${withHistory}</span></button>
                <button class="bbz-kpi-chip ${kpiMode==="tasks"?"bbz-kpi-chip-active":""}" data-action="kpi-filter" data-scope="contacts-mode" data-value="tasks">Offene Tasks <span>${withOpenTasks}</span></button>
                <button class="bbz-kpi-chip ${kpiMode==="all"||!kpiMode?"bbz-kpi-chip-active":""}" data-action="kpi-filter" data-scope="contacts-mode" data-value="all">Alle</button>
              </div>
            </div>
            <!-- Sichtbar nach Filter -->
            ${this.kpiBlock("Angezeigt", rows.length, rows.length < totalActive ? `von ${totalActive}` : "alle aktiven")}
            <!-- Offene Tasks — klickbar zur Planung -->
            <div class="bbz-kpi bbz-kpi-clickable" data-action="navigate-planning" style="cursor:pointer;">
              <div class="bbz-kpi-label">Offene Tasks</div>
              <div class="bbz-kpi-value">${allOpenTasks}</div>
              ${overdueTasks > 0
                ? `<div class="bbz-kpi-meta-alert">${overdueTasks} überfällig — zur Planung →</div>`
                : `<div class="bbz-kpi-meta-ok">keine überfällig — zur Planung →</div>`}
            </div>
            <!-- Zurück zum Cockpit -->
            <div class="bbz-kpi bbz-kpi-clickable" data-action="kpi-filter" data-scope="navigate" data-value="firms" style="cursor:pointer;">
              <div class="bbz-kpi-label">Firmen-Cockpit</div>
              <div class="bbz-kpi-value">${state.enriched.firms.length}</div>
              <div class="bbz-kpi-meta">← zurück zum Cockpit</div>
            </div>
          </div>
          <section class="bbz-section">
          <div class="bbz-section-header">
            <div><div class="bbz-section-title">Kontakte</div><div class="bbz-section-subtitle">${kpiMode === "history" ? "Mit History-Einträgen" : kpiMode === "tasks" ? "Mit offenen Tasks" : "Operative Ansprechpartner über alle Firmen"}</div></div>
            <div style="display:flex;align-items:center;gap:8px;">
              <button class="bbz-dense-toggle" onclick="window.bbzToggleDense && window.bbzToggleDense()" title="Kompakte Ansicht">⇕ Kompakt</button>
              <button class="bbz-button bbz-button-secondary" data-action="open-history-form">+ Aktivität</button>
              <button class="bbz-button bbz-button-primary" data-action="open-contact-form">+ Kontakt</button>
            </div>
          </div>
          <div class="bbz-section-body">
            <div class="bbz-filters-3">
              <input class="bbz-input" data-filter="contacts-search" type="text" placeholder="Suche nach Name, Firma, Funktion, Rolle, E-Mail ..." value="${helpers.escapeHtml(filters.search)}" />
              <label class="bbz-checkbox"><input type="checkbox" data-filter="contacts-archiviert" ${filters.archiviertAusblenden ? "checked" : ""} /> Archivierte ausblenden</label>
              <div></div>
            </div>
            <div class="bbz-table-wrap">
              <table class="bbz-table">
                <thead><tr>${cTh("Name","fullName")}${cTh("Firma","firmTitle")}<th>Funktion</th>${cTh("Rolle","rolle")}${cTh("Lead BBZ","leadbbz0")}<th>E-Mail</th><th>Telefon</th><th>Archiviert</th></tr></thead>
                <tbody>
                  ${rows.length ? rows.map(c => `
                    <tr>
                      <td><span class="bbz-td-name">${helpers.avatarHtml(c)}<a class="bbz-link" data-action="open-contact" data-id="${c.id}">${helpers.escapeHtml(c.fullName || c.nachname)}</a></span></td>
                      <td>${c.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${c.firmId}">${helpers.escapeHtml(c.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(c.funktion) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(c.rolle) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(c.leadbbz0) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${c.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(c.email1)}">${helpers.escapeHtml(c.email1)}</a>` : '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(helpers.joinNonEmpty([c.direktwahl, c.mobile], " / ")) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${c.archiviert ? '<span class="bbz-danger">Ja</span>' : '<span class="bbz-muted">Nein</span>'}</td>
                    </tr>`).join("") : `<tr><td colspan="8">${ui.emptyBlock("Keine Kontakte fuer die aktuelle Filterung gefunden.")}</td></tr>`}
                </tbody>
              </table>
            </div>
          </div>
          </section>
        </div>
      `;
    },

    contactDetail() {
      const contact = dataModel.getContactById(state.selection.contactId);
      if (!contact) return ui.emptyBlock("Der ausgewaehlte Kontakt wurde nicht gefunden.");
      const contactHistory = state.enriched.history.filter(h => h.contactId === contact.id).sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));
      const contactTasks = state.enriched.tasks.filter(t => t.contactId === contact.id).sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline));
      const isPrivat = state.meta.privateFirmId !== null && contact.firmId === state.meta.privateFirmId;
      // Band-Farbe von der Firma erben
      const firm = contact.firmId ? dataModel.getFirmById(contact.firmId) : null;
      const bandClass = firm ? helpers.detailBandClass(firm) : "bbz-detail-band bbz-detail-band-default";
      const seed = [...(contact.vorname.charAt(0) + contact.nachname.charAt(0))].reduce((s, c) => s + c.charCodeAt(0), 0);
      const avatarIdx = seed % 6;
      const initials = (contact.vorname.charAt(0) + contact.nachname.charAt(0)).toUpperCase() || "?";

      return `
        <div>
          <div class="${bandClass}" style="margin-bottom:10px;">
            <button class="bbz-button bbz-button-secondary" style="margin-bottom:12px;background:rgba(255,255,255,0.7);" data-action="back-to-contacts">← Kontaktliste</button>
            <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:16px;flex-wrap:wrap;">
              <div style="display:flex;align-items:center;gap:14px;">
                <div class="bbz-avatar-lg" data-idx="${avatarIdx}">${helpers.escapeHtml(initials)}</div>
                <div>
                  <div class="bbz-detail-title">${helpers.escapeHtml(contact.fullName || contact.nachname)}</div>
                  <div class="bbz-detail-subtitle">
                    ${isPrivat
                      ? `<span class="bbz-pill" style="font-size:12px;">Privatperson</span>`
                      : contact.firmId
                        ? `<a class="bbz-link" data-action="open-firm" data-id="${contact.firmId}">${helpers.escapeHtml(contact.firmTitle || "Firma")}</a>`
                        : "Keine Firma verknüpft"
                    }
                    ${contact.funktion ? ` · ${helpers.escapeHtml(contact.funktion)}` : ""}
                    ${contact.rolle ? ` · ${helpers.escapeHtml(contact.rolle)}` : ""}
                  </div>
                  <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:8px;">
                    ${contact.leadbbz0 ? helpers.leadbbzBadgeHtml(contact.leadbbz0) : ""}
                    ${contact.archiviert ? '<span class="bbz-pill" style="background:var(--red-soft);color:var(--red);border-color:#f0b0b2;">Archiviert</span>' : ""}
                  </div>
                </div>
              </div>
              <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                ${contact.email1 ? `<a class="bbz-button bbz-button-secondary" href="mailto:${helpers.escapeHtml(contact.email1)}">✉ Mail</a>` : ""}
                <button class="bbz-button bbz-button-secondary" style="color:var(--red);border-color:var(--red);" data-action="delete-contact" data-id="${contact.id}" data-name="${helpers.escapeHtml(contact.fullName || contact.nachname)}">Löschen</button>
                <button class="bbz-button bbz-button-secondary" data-action="open-contact-form" data-item-id="${contact.id}">Bearbeiten</button>
                <button class="bbz-button bbz-button-secondary" data-action="open-task-form" data-contact-id="${contact.id}">+ Task</button>
                <button class="bbz-button bbz-button-primary" data-action="open-history-form" data-contact-id="${contact.id}">+ Aktivität</button>
              </div>
            </div>
          </div>
          <div class="bbz-kpis">
            ${this.kpiBlock("Tasks", contactTasks.length)}
            ${this.kpiBlock("Offen", contactTasks.filter(t => t.isOpen).length, contactTasks.some(t => t.isOpen && t.isOverdue) ? "überfällig" : contactTasks.filter(t => t.isOpen).length > 0 ? "offen" : "keine offen", contactTasks.some(t => t.isOpen && t.isOverdue) ? "alert" : "")}
            ${this.kpiBlock("Aktivitäten", contactHistory.length)}
            ${this.kpiBlock("Letzter Kontakt", contactHistory[0]?.datum ? helpers.relativeDate(contactHistory[0].datum) : "—", contactHistory.length === 0 ? "noch kein Kontakt" : "", contactHistory.length === 0 ? "warn" : "")}
          </div>
          <div class="bbz-grid bbz-grid-3">
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Stammdaten</div></div>
              <div class="bbz-section-body">
                ${ui.kv("Anrede", helpers.escapeHtml(contact.anrede) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Vorname", helpers.escapeHtml(contact.vorname) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Nachname", helpers.escapeHtml(contact.nachname) || '<span class="bbz-muted">—</span>')}
                ${isPrivat
                  ? ui.kv("Adresse / Notizen", helpers.escapeHtml(contact.kommentar) || '<span class="bbz-muted">—</span>')
                  : ui.kv("Firma", contact.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${contact.firmId}">${helpers.escapeHtml(contact.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>')
                }
                ${ui.kv("Funktion", helpers.escapeHtml(contact.funktion) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Rolle", helpers.escapeHtml(contact.rolle) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Email 1", contact.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(contact.email1)}">${helpers.escapeHtml(contact.email1)}</a>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("Email 2", contact.email2 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(contact.email2)}">${helpers.escapeHtml(contact.email2)}</a>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("Direktwahl", helpers.escapeHtml(contact.direktwahl) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Mobile", helpers.escapeHtml(contact.mobile) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Geburtstag", helpers.formatDate(contact.geburtstag) || '<span class="bbz-muted">—</span>')}
              </div>
            </section>
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">CRM-Kontext</div></div>
              <div class="bbz-section-body">
                ${ui.kv("Lead BBZ", helpers.leadbbzBadgeHtml(contact.leadbbz0))}
                ${ui.kv("SGF", helpers.multiChoiceHtml(contact.sgf))}
                ${ui.kv("Event", helpers.multiChoiceHtml(contact.event))}
                ${ui.kv("Eventhistory", helpers.multiChoiceHtml(contact.eventhistory))}
                ${isPrivat ? "" : ui.kv("Kommentar", helpers.escapeHtml(contact.kommentar) || '<span class="bbz-muted">—</span>')}
              </div>
            </section>
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Aufgaben</div>
                <button class="bbz-button bbz-button-secondary" style="height:28px;font-size:12px;" data-action="open-task-form" data-contact-id="${contact.id}">+ Task</button>
              </div>
              <div class="bbz-section-body">
                ${contactTasks.length ? contactTasks.map(t => `
                  <div style="display:flex;align-items:center;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--line-2);">
                    <div>
                      <div style="font-size:13px;font-weight:600;">${helpers.escapeHtml(t.title)}</div>
                      <div style="font-size:12px;color:var(--muted);margin-top:2px;">${t.deadline ? helpers.relativeDate(t.deadline) : "Keine Deadline"}</div>
                    </div>
                    ${helpers.statusChipHtml(t.status, t.deadline)}
                  </div>`).join("") : `<div class="bbz-empty">Noch kein Task erfasst.<br><button class="bbz-button bbz-button-secondary" style="margin-top:10px;height:32px;font-size:13px;" data-action="open-task-form" data-contact-id="${contact.id}">+ Ersten Task erstellen</button></div>`}
              </div>
            </section>
          </div>
          <div class="bbz-grid bbz-grid-2" style="margin-top:12px;">
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Aktivitäten</div></div>
                <button class="bbz-button bbz-button-primary" style="height:32px;font-size:13px;" data-action="open-history-form" data-contact-id="${contact.id}">+ Aktivität</button>
              </div>
              <div class="bbz-section-body">
                ${contactHistory.length ? `<div class="bbz-timeline">${contactHistory.map(h => `
                  <div class="bbz-timeline-item">
                    <div class="bbz-timeline-date">${helpers.relativeDate(h.datum) || "—"}<br><span class="bbz-muted" style="font-size:11px;">${helpers.formatDate(h.datum)}</span></div>
                    <div>
                      <div class="bbz-timeline-title">${helpers.escapeHtml(h.typ || h.title || "Eintrag")} ${h.projektbezugBool ? '<span class="bbz-chip" style="background:var(--blue-light);color:var(--blue);border-color:#a8c8e0;">Projektbezug</span>' : '<span class="bbz-chip">Allgemein</span>'}</div>
                      <div class="bbz-timeline-text">${helpers.escapeHtml(h.notizen || "—")}</div>
                      <div style="margin-top:6px;display:flex;gap:6px;">
                        <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;" data-action="edit-history" data-id="${h.id}">Bearbeiten</button>
                        <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;color:var(--red);border-color:var(--red);" data-action="delete-history" data-id="${h.id}" data-title="${helpers.escapeHtml(h.typ || h.title || 'Eintrag')}">Löschen</button>
                      </div>
                    </div>
                  </div>`).join("")}</div>` : ui.emptyBlock("Noch keine Aktivitäten erfasst.")}
              </div>
            </section>
          </div>
        </div>
      `;
    },

    planning() {
      const filters = state.filters.planning;
      const statusChoices = state.meta.choices?.[CONFIG.lists.tasks]?.["Status"] || [];
      const today = helpers.todayStart();
      const in7   = new Date(today); in7.setDate(in7.getDate() + 7);
      const in30  = new Date(today); in30.setDate(in30.getDate() + 30);
      const in365 = new Date(today); in365.setDate(in365.getDate() + 365);

      const baseRows = state.enriched.tasks.filter(t => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [t.title, t.status, t.contactName, t.firmTitle, t.leadbbz].some(v => helpers.textIncludes(v, search));
        if (!searchMatch) return false;
        if (filters.onlyOpen && !t.isOpen) return false;

        // Fälligkeits-Filter
        if (filters.faelligkeit) {
          const d = helpers.toDate(t.deadline);
          if (filters.faelligkeit === "overdue" && !(t.isOpen && t.isOverdue)) return false;
          if (filters.faelligkeit === "week"    && !(d && d > today && d <= in7))  return false;
          if (filters.faelligkeit === "month"   && !(d && d > in7  && d <= in30))  return false;
          if (filters.faelligkeit === "rest"    && !(d && d > in30 && d <= in365)) return false;
        }

        // Segment-Filter
        if (filters.segment) {
          const firm = t.firmId ? state.enriched.firms.find(f => f.id === t.firmId) : null;
          const kl = String(firm?.klassifizierung || "").toUpperCase();
          if (!kl.startsWith(filters.segment.toUpperCase())) return false;
        }

        // Lead BBZ-Filter
        if (filters.leadbbz && t.leadbbz !== filters.leadbbz) return false;

        return true;
      });

      // Zähler für Chips
      const cntOverdue = state.enriched.tasks.filter(t => t.isOpen && t.isOverdue).length;
      const cntWeek    = state.enriched.tasks.filter(t => { const d = helpers.toDate(t.deadline); return t.isOpen && d && d > today && d <= in7; }).length;
      const cntMonth   = state.enriched.tasks.filter(t => { const d = helpers.toDate(t.deadline); return t.isOpen && d && d > in7 && d <= in30; }).length;
      const cntRest    = state.enriched.tasks.filter(t => { const d = helpers.toDate(t.deadline); return t.isOpen && d && d > in30 && d <= in365; }).length;

      const allLeadbbz = [...new Set(state.enriched.tasks.map(t => t.leadbbz).filter(Boolean))].sort();

      const chipF = (label, val, cnt, style = "") => {
        const active = filters.faelligkeit === val;
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}" style="${style}" data-action="kpi-filter" data-scope="planning-faelligkeit" data-value="${val}">${label} <span>${cnt}</span></button>`;
      };
      const chipS = (label, val) => {
        const cnt = val === "" ? state.enriched.tasks.length : state.enriched.tasks.filter(t => {
          const firm = t.firmId ? state.enriched.firms.find(f => f.id === t.firmId) : null;
          return String(firm?.klassifizierung || "").toUpperCase().startsWith(val);
        }).length;
        const active = filters.segment === val;
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="planning-segment" data-value="${val}">${label === "Alle" ? "Alle" : label} <span>${cnt}</span></button>`;
      };
      const chipL = (l) => {
        const cnt = state.enriched.tasks.filter(t => t.leadbbz === l).length;
        const active = filters.leadbbz === l;
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="planning-leadbbz" data-value="${helpers.escapeHtml(l)}">${helpers.escapeHtml(l)} <span>${cnt}</span></button>`;
      };

      // Multi-column sort
      const dir = filters.sortDir === "asc" ? 1 : -1;
      const sorted = [...baseRows].sort((a, b) => {
        if (filters.sortBy === "deadline") {
          const ad = helpers.toDate(a.deadline), bd = helpers.toDate(b.deadline);
          if (!ad && !bd) return 0;
          if (!ad) return 1;
          if (!bd) return -1;
          return (ad - bd) * dir;
        }
        if (filters.sortBy === "leadbbz") {
          return String(a.leadbbz || "").localeCompare(String(b.leadbbz || ""), "de") * dir;
        }
        if (filters.sortBy === "firma") {
          return String(a.firmTitle || "").localeCompare(String(b.firmTitle || ""), "de") * dir;
        }
        return 0;
      });

      // Gruppierung
      const groupBy = filters.groupBy;
      let groups = [];
      if (groupBy === "none") {
        groups = [{ key: null, label: null, rows: sorted }];
      } else {
        const map = new Map();
        sorted.forEach(t => {
          const key = (groupBy === "status" ? t.status : t.leadbbz) || "—";
          if (!map.has(key)) map.set(key, []);
          map.get(key).push(t);
        });
        groups = [...map.entries()]
          .sort((a, b) => a[0].localeCompare(b[0], "de"))
          .map(([key, rows]) => ({ key, label: key, rows }));
      }

      // Sort-Header helper
      const th = (label, col) => {
        const active = filters.sortBy === col;
        const icon = active ? (filters.sortDir === "asc" ? " ↑" : " ↓") : "";
        return `<th style="cursor:pointer;user-select:none;${active ? "color:var(--blue);" : ""}" data-action="set-sort" data-col="${col}">${label}${icon}</th>`;
      };

      const renderTaskRow = (t) => {
        const statusCell = statusChoices.length
          ? `<select class="bbz-select" style="height:30px;font-size:12px;" data-action="task-status-change" data-task-id="${t.id}">${statusChoices.map(s => `<option value="${helpers.escapeHtml(s)}" ${t.status === s ? "selected" : ""}>${helpers.escapeHtml(s)}</option>`).join("")}</select>`
          : helpers.statusChipHtml(t.status, t.deadline);
        return `
          <tr>
            <td>${t.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${t.firmId}">${helpers.escapeHtml(t.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
            <td>${t.contactId ? `<a class="bbz-link" data-action="open-contact" data-id="${t.contactId}">${helpers.escapeHtml(t.contactName || "Kontakt")}</a>` : helpers.escapeHtml(t.contactName || "—")}</td>
            <td>${helpers.escapeHtml(t.title) || '<span class="bbz-muted">—</span>'}</td>
            <td class="${helpers.isOpenTask(t.status) && helpers.isOverdue(t.deadline) ? "bbz-danger" : ""}">${t.deadline ? helpers.relativeDate(t.deadline) : '<span class="bbz-muted">—</span>'}</td>
            <td>${statusCell}</td>
            <td>${helpers.escapeHtml(t.leadbbz) || '<span class="bbz-muted">—</span>'}</td>
            <td style="white-space:nowrap;">
              <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;margin-right:4px;" data-action="edit-task" data-id="${t.id}">Bearbeiten</button>
              <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;color:var(--red);border-color:var(--red);" data-action="delete-task" data-id="${t.id}" data-title="${helpers.escapeHtml(t.title)}">Löschen</button>
            </td>
          </tr>`;
      };

      const tableHead = `<thead><tr>
        ${th("Firma", "firma")}
        <th>Kontaktperson</th>
        <th>Titel</th>
        ${th("Deadline", "deadline")}
        <th>Status</th>
        ${th("Leadbbz", "leadbbz")}
        <th>Aktionen</th>
      </tr></thead>`;

      const tableBody = groups.map(g => `
        ${g.label ? `<tr><td colspan="7" style="background:#f1f5fb;font-size:12px;font-weight:700;color:var(--text);padding:7px 12px;">${helpers.escapeHtml(g.label)} <span style="font-weight:400;color:var(--muted);">(${g.rows.length})</span></td></tr>` : ""}
        ${g.rows.length ? g.rows.map(renderTaskRow).join("") : `<tr><td colspan="7">${ui.emptyBlock("Keine Tasks.")}</td></tr>`}
      `).join("");

      return `
        <div>
          <div class="bbz-kpis">
            <!-- Kachel 1: Fälligkeit -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Tasks gesamt</div>
              <div class="bbz-kpi-value">${state.enriched.tasks.length}</div>
              <div style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${chipF("Überfällig", "overdue", cntOverdue, cntOverdue > 0 ? "background:var(--red-soft);border-color:#f0b0b2;color:var(--red);" : "")}
                ${chipF("Woche", "week", cntWeek, cntWeek > 0 ? "background:#fff9eb;border-color:#f4dfab;color:var(--amber);" : "")}
                ${chipF("Monat", "month", cntMonth)}
                ${chipF("Übrige", "rest", cntRest)}
                ${filters.faelligkeit ? `<button class="bbz-kpi-chip" data-action="kpi-filter" data-scope="planning-faelligkeit" data-value="">Alle</button>` : ""}
              </div>
            </div>
            <!-- Kachel 2: Kundenklassifizierung -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Kundenklassifizierung</div>
              <div class="bbz-kpi-value">${filters.segment || "—"}</div>
              <div style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${["A","B","C"].map(k => chipS(k, k)).join("")}
                <button class="bbz-kpi-chip ${!filters.segment ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="planning-segment" data-value="">Alle</button>
              </div>
            </div>
            <!-- Kachel 3: Lead BBZ -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Lead BBZ</div>
              <div class="bbz-kpi-value">${filters.leadbbz ? helpers.escapeHtml(filters.leadbbz) : "—"}</div>
              <div style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${allLeadbbz.map(chipL).join("")}
                ${filters.leadbbz ? `<button class="bbz-kpi-chip" data-action="kpi-filter" data-scope="planning-leadbbz" data-value="">Alle</button>` : ""}
              </div>
            </div>
            <!-- Kachel 4: Sichtbar -->
            ${this.kpiBlock("Sichtbar", baseRows.length, baseRows.length < state.enriched.tasks.length ? `von ${state.enriched.tasks.length} gefiltert` : "alle", baseRows.length < state.enriched.tasks.length ? "warn" : "")}
          </div>
          <section class="bbz-section">
            <div class="bbz-section-header">
              <div><div class="bbz-section-title">Planung</div><div class="bbz-section-subtitle">Aufgaben mit Fokus auf offen und überfällig</div></div>
              <button class="bbz-button bbz-button-primary" data-action="open-task-form">+ Task</button>
            </div>
            <div class="bbz-section-body">
              <div class="bbz-filters-3" style="grid-template-columns:2fr 1fr 1fr;">
                <input class="bbz-input" data-filter="planning-search" type="text" placeholder="Suche nach Titel, Firma, Kontakt, Status ..." value="${helpers.escapeHtml(filters.search)}" />
                <select class="bbz-select" data-filter="planning-groupby">
                  <option value="none"    ${filters.groupBy === "none"    ? "selected" : ""}>Keine Gruppierung</option>
                  <option value="status"  ${filters.groupBy === "status"  ? "selected" : ""}>Gruppe: Status</option>
                  <option value="leadbbz" ${filters.groupBy === "leadbbz" ? "selected" : ""}>Gruppe: Leadbbz</option>
                </select>
                <label class="bbz-checkbox"><input type="checkbox" data-filter="planning-open" ${filters.onlyOpen ? "checked" : ""} /> Nur offene Tasks</label>
              </div>
              <div class="bbz-table-wrap">
                <table class="bbz-table" style="min-width:1060px;">
                  ${tableHead}
                  <tbody>${baseRows.length ? tableBody : `<tr><td colspan="7">${ui.emptyBlock("Keine Tasks fuer die aktuelle Filterung gefunden.")}</td></tr>`}</tbody>
                </table>
              </div>
            </div>
          </section>
        </div>
      `;
    },

    historyView() {
      const filters = state.filters.history;
      const today   = helpers.todayStart();

      // ── Filterfunktion ──────────────────────────────────────────────────────
      const applyFilters = (h) => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [h.contactName, h.firmTitle, h.typ, h.notizen, h.leadbbz].some(v => helpers.textIncludes(v, search));
        const artMatch    = !filters.kontaktart || h.typ === filters.kontaktart;
        const leadMatch   = !filters.leadbbz    || h.leadbbz === filters.leadbbz;
        if (!searchMatch || !artMatch || !leadMatch) return false;
        if (filters.zeitfenster === "today") {
          const d = helpers.toDate(h.datum);
          return d && d >= today;
        }
        if (filters.zeitfenster === "week") {
          const d = helpers.toDate(h.datum);
          const vor7 = new Date(today); vor7.setDate(vor7.getDate() - 7);
          return d && d >= vor7;
        }
        if (filters.zeitfenster === "month") {
          const d = helpers.toDate(h.datum);
          const vor30 = new Date(today); vor30.setDate(vor30.getDate() - 30);
          return d && d >= vor30;
        }
        return true;
      };

      const rows = state.enriched.history.filter(applyFilters);

      // ── KPI-Zahlen ──────────────────────────────────────────────────────────
      const totalEntries = state.enriched.history.length;
      const mitProjekt   = state.enriched.history.filter(h => h.projektbezugBool).length;
      const cntToday     = state.enriched.history.filter(h => { const d = helpers.toDate(h.datum); return d && d >= today; }).length;
      const vor7         = new Date(today); vor7.setDate(vor7.getDate() - 7);
      const cntWeek      = state.enriched.history.filter(h => { const d = helpers.toDate(h.datum); return d && d >= vor7; }).length;
      const vor30        = new Date(today); vor30.setDate(vor30.getDate() - 30);
      const cntMonth     = state.enriched.history.filter(h => { const d = helpers.toDate(h.datum); return d && d >= vor30; }).length;

      // Vorwoche-Vergleich für Trend
      const vor14 = new Date(today); vor14.setDate(vor14.getDate() - 14);
      const cntVorwoche = state.enriched.history.filter(h => { const d = helpers.toDate(h.datum); return d && d >= vor14 && d < vor7; }).length;
      const trendDiff   = cntWeek - cntVorwoche;
      const trendHtml   = trendDiff > 0
        ? `<span style="color:var(--green);font-weight:700;">↑ +${trendDiff} vs. Vorwoche</span>`
        : trendDiff < 0
        ? `<span style="color:var(--red);font-weight:700;">↓ ${trendDiff} vs. Vorwoche</span>`
        : `<span style="color:var(--muted);">= Vorwoche</span>`;

      // ── KPI-Kacheln ─────────────────────────────────────────────────────────
      const chipZeit = (label, value, count) => {
        const active = filters.zeitfenster === value;
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="history-zeitfenster" data-value="${value}">${label} <span>${count}</span></button>`;
      };
      const chipKontaktart = (k) => {
        const cnt   = state.enriched.history.filter(h => h.typ === k).length;
        const active = filters.kontaktart === k;
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="history-kontaktart" data-value="${helpers.escapeHtml(k)}">${helpers.escapeHtml(k)} <span>${cnt}</span></button>`;
      };
      const chipLead = (l) => {
        const cnt   = state.enriched.history.filter(h => h.leadbbz === l).length;
        const active = filters.leadbbz === l;
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="history-leadbbz" data-value="${helpers.escapeHtml(l)}">${helpers.escapeHtml(l)} <span>${cnt}</span></button>`;
      };

      const allKontaktart = [...new Set(state.enriched.history.map(h => h.typ).filter(Boolean))].sort();
      const allLeadbbz    = [...new Set(state.enriched.history.map(h => h.leadbbz).filter(Boolean))].sort();

      // Aktive Filter-Badges für Filterleiste
      const activeFilterBadges = [
        filters.kontaktart ? `<span class="bbz-chip" style="background:var(--blue-light);color:var(--blue);border-color:#a8c8e0;cursor:pointer;" data-action="kpi-filter" data-scope="history-kontaktart" data-value="">${helpers.escapeHtml(filters.kontaktart)} ×</span>` : "",
        filters.leadbbz    ? `<span class="bbz-chip" style="background:#f0fdf4;color:#15803d;border-color:#86efac;cursor:pointer;" data-action="kpi-filter" data-scope="history-leadbbz" data-value="">${helpers.escapeHtml(filters.leadbbz)} ×</span>` : "",
        filters.zeitfenster ? `<span class="bbz-chip" style="cursor:pointer;" data-action="kpi-filter" data-scope="history-zeitfenster" data-value="">${filters.zeitfenster === "today" ? "Heute" : filters.zeitfenster === "week" ? "Diese Woche" : "Dieser Monat"} ×</span>` : ""
      ].filter(Boolean).join("");

      // ── Timeline-Gruppierung ─────────────────────────────────────────────────
      const renderCard = (h, hideFirm = false) => {
        const hasLongText = (h.notizen || "").length > 120;
        const tagsHtml = `
          <div style="display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end;flex-shrink:0;margin-left:8px;">
            ${h.projektbezugBool ? '<span class="bbz-chip" style="background:var(--blue-light);color:var(--blue);border-color:#a8c8e0;">Projektbezug</span>' : ""}
            ${h.leadbbz ? helpers.leadbbzBadgeHtml(h.leadbbz) : ""}
          </div>`;
        return `
          <div class="bbz-timeline-item">
            <div style="display:flex;align-items:center;gap:10px;margin-bottom:5px;">
              ${helpers.avatarHtml({ vorname: (h.contactName||"").split(" ")[0]||"", nachname: (h.contactName||"").split(" ").slice(-1)[0]||"" })}
              <div style="flex:1;min-width:0;">
                <div style="font-size:13px;font-weight:700;line-height:1.25;">
                  ${h.contactId ? `<a class="bbz-link" data-action="open-contact" data-id="${h.contactId}">${helpers.escapeHtml(h.contactName || "")}</a>` : helpers.escapeHtml(h.contactName || "—")}
                </div>
                ${!hideFirm ? `<div style="font-size:12px;color:var(--muted);margin-top:1px;">
                  ${h.firmId ? `<a class="bbz-link" style="font-weight:500;" data-action="open-firm" data-id="${h.firmId}">${helpers.escapeHtml(h.firmTitle)}</a>` : (h.firmTitle ? helpers.escapeHtml(h.firmTitle) : "")}
                </div>` : ""}
              </div>
              <div style="text-align:right;flex-shrink:0;">
                <div style="font-size:12px;font-weight:600;color:var(--text);" title="${helpers.formatDate(h.datum)}">${helpers.relativeDate(h.datum) || "—"}</div>
                <div style="font-size:11px;color:var(--muted);">${helpers.formatDate(h.datum)}</div>
              </div>
            </div>
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:5px;">
              <span style="font-size:13px;font-weight:600;color:var(--text);">${helpers.escapeHtml(h.typ || "Eintrag")}</span>
              ${tagsHtml}
            </div>
            <div class="bbz-timeline-text bbz-timeline-clamp">${helpers.escapeHtml(h.notizen || "—")}</div>
            ${hasLongText ? `<button class="bbz-button bbz-button-secondary" style="height:22px;font-size:11px;padding:0 8px;margin-top:4px;" data-action="toggle-expand">mehr</button>` : ""}
            <div style="margin-top:8px;display:flex;gap:6px;align-items:center;">
              <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;" data-action="edit-history" data-id="${h.id}">Bearbeiten</button>
              <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;" data-action="open-task-form" data-contact-id="${h.contactId || ""}" title="Task für ${helpers.escapeHtml(h.contactName || "")} erstellen">+ Task</button>
            </div>
          </div>`;
      };

      // Datum-Gruppen
      const makeDateGroups = (items) => {
        const vor7d  = new Date(today); vor7d.setDate(vor7d.getDate() - 7);
        const vor30d = new Date(today); vor30d.setDate(vor30d.getDate() - 30);
        const groups = [
          { key: "today",  label: "Heute",          items: [] },
          { key: "week",   label: "Diese Woche",     items: [] },
          { key: "month",  label: "Dieser Monat",    items: [] },
          { key: "older",  label: "Älter",           items: [] }
        ];
        items.forEach(h => {
          const d = helpers.toDate(h.datum);
          if (!d) { groups[3].items.push(h); return; }
          if (d >= today)   groups[0].items.push(h);
          else if (d >= vor7d)  groups[1].items.push(h);
          else if (d >= vor30d) groups[2].items.push(h);
          else                  groups[3].items.push(h);
        });
        return groups.filter(g => g.items.length > 0);
      };

      // Firma-Gruppen — mit Segment-Farbe für den Header
      const makeFirmGroups = (items) => {
        const map = new Map();
        items.forEach(h => {
          const key = h.firmTitle || "Ohne Firma";
          if (!map.has(key)) map.set(key, { firmId: h.firmId, items: [] });
          map.get(key).items.push(h);
        });
        return [...map.entries()]
          .sort((a, b) => a[0].localeCompare(b[0], "de"))
          .map(([label, val]) => {
            const firm = val.firmId ? state.enriched.firms.find(f => f.id === val.firmId) : null;
            return { key: label, label, firmId: val.firmId, firm, latestDatum: val.items[0]?.datum || "", items: val.items };
          });
      };

      const groups = filters.groupBy === "firm" ? makeFirmGroups(rows) : makeDateGroups(rows);

      // Firma-Gruppen-Header: prominent mit Segment-Farbe und Aktions-Button
      const firmGroupHeader = (g) => {
        const firm = g.firm;
        const kl   = String(firm?.klassifizierung || "").toUpperCase();
        const borderColor = kl.includes("A") ? "var(--blue)" : kl.includes("B") ? "#d97706" : "#64748b";
        const bgColor     = kl.includes("A") ? "#f0f7ff" : kl.includes("B") ? "#fffdf0" : "#f8fafc";
        return `
          <div class="bbz-history-firm-header" style="border-left:4px solid ${borderColor};background:${bgColor};">
            <div style="display:flex;align-items:center;gap:8px;flex:1;min-width:0;">
              <span class="bbz-history-group-label" style="font-size:14px;">
                ${g.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${g.firmId}">${helpers.escapeHtml(g.label)}</a>` : helpers.escapeHtml(g.label)}
              </span>
              ${firm?.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : ""}
              ${firm?.vip ? `<span class="bbz-pill bbz-pill-vip">♛</span>` : ""}
              <span class="bbz-history-group-count">${g.items.length} Eintrag${g.items.length !== 1 ? "e" : ""}</span>
              ${g.latestDatum ? `<span class="bbz-history-group-meta">zuletzt ${helpers.relativeDate(g.latestDatum)}</span>` : ""}
            </div>
            <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;flex-shrink:0;"
              data-action="open-history-form" data-firm-id="${g.firmId || ""}">+ Aktivität</button>
          </div>`;
      };

      // Aktiv gefilterte Firma — für Guard im Header und Empty State
      const activeFirm = filters.search
        ? state.enriched.firms.find(f => f.title === filters.search) || null
        : null;
      const activeFirmHasNoContacts = activeFirm && activeFirm.contacts.length === 0;

      // Empty State — kontextabhängig
      const emptyStateHtml = activeFirmHasNoContacts
        ? `<div class="bbz-empty">
             <strong>${helpers.escapeHtml(activeFirm.title)}</strong> hat noch keine Kontakte.<br>
             Bitte zuerst einen Kontakt erfassen, bevor eine Aktivität hinzugefügt werden kann.<br>
             <button class="bbz-button bbz-button-secondary" style="margin-top:10px;height:32px;font-size:13px;"
               data-action="open-contact-form" data-firm-id="${activeFirm.id}">+ Kontakt erfassen</button>
           </div>`
        : ui.emptyBlock("Keine Aktivitäten für die aktuelle Filterung gefunden.", "open-history-form", "+ Erste Aktivität erfassen");

      const timelineHtml = groups.length ? groups.map(g => `
        <div class="bbz-history-group">
          ${filters.groupBy === "firm"
            ? firmGroupHeader(g)
            : `<div class="bbz-history-group-header">
                <span class="bbz-history-group-label">${helpers.escapeHtml(g.label)}</span>
                <span class="bbz-history-group-count">${g.items.length} Eintrag${g.items.length !== 1 ? "e" : ""}</span>
               </div>`}
          <div class="bbz-timeline">${g.items.map(h => renderCard(h, filters.groupBy === "firm")).join("")}</div>
        </div>`).join("")
        : emptyStateHtml;

      // ── Pflege-Radar Panel ───────────────────────────────────────────────────
      const abFirmen = state.enriched.firms.filter(f => {
        const kl = String(f.klassifizierung || "").toUpperCase();
        return kl.includes("A") || kl.includes("B");
      });

      const radarNever = abFirmen.filter(f => {
        const kl = String(f.klassifizierung || "").toUpperCase();
        return kl.includes("A") && f.history.length === 0;
      }).sort((a, b) => a.title.localeCompare(b.title, "de"));

      const radarCold = abFirmen.filter(f => {
        if (f.history.length === 0) return false;
        const last = helpers.toDate(f.latestActivity);
        if (!last) return false;
        const diff = Math.floor((today - last) / 86400000);
        return diff > 360;
      }).sort((a, b) => helpers.compareDateAsc(a.latestActivity, b.latestActivity));

      const radarOverdue = abFirmen.filter(f =>
        f.tasks.some(t => t.isOpen && t.isOverdue)
      ).sort((a, b) => helpers.compareDateAsc(a.nextDeadline, b.nextDeadline));

      const radarOk = abFirmen.filter(f => helpers.firmSignal(f) === "ok")
        .sort((a, b) => helpers.compareDateDesc(a.latestActivity, b.latestActivity));

      const activeRadarFirm = filters.search;

      const radarItem = (f, meta, tooltip, accentColor = "var(--muted)", showTaskBtn = false) => {
        const kl = String(f.klassifizierung || "").toUpperCase();
        const isA = kl.includes("A");
        const isActive = activeRadarFirm === f.title;
        return `
          <div class="bbz-mini-item bbz-radar-item ${isActive ? "bbz-radar-item-active" : ""}"
               style="cursor:pointer;border-left:3px solid ${accentColor};padding-left:9px;"
               title="${helpers.escapeHtml(tooltip)}">
            <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;">
              <div style="display:flex;align-items:center;gap:6px;min-width:0;"
                   data-action="history-firma-filter" data-firm-title="${helpers.escapeHtml(f.title)}">
                <span class="bbz-mini-title" style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${helpers.escapeHtml(f.title)}</span>
                <span class="${helpers.firmBadgeClass(f.klassifizierung)}" style="flex-shrink:0;">${isA ? "A" : "B"}</span>
              </div>
              <div style="display:flex;align-items:center;gap:6px;flex-shrink:0;">
                <span class="bbz-mini-meta" style="white-space:nowrap;color:${accentColor};font-weight:600;">${meta}</span>
                ${showTaskBtn ? `<button class="bbz-button bbz-button-secondary" style="height:22px;font-size:11px;padding:0 7px;"
                  data-action="open-task-form" data-firm-id="${f.id}">+ Task</button>` : ""}
              </div>
            </div>
          </div>`;
      };

      const radarZoneHtml = (label, color, items, emptyNull = true) => {
        if (items.length === 0 && emptyNull) return "";
        return `
          <div style="margin-bottom:12px;">
            <div style="color:${color};font-size:11px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;margin-bottom:6px;">
              ${label} <span style="font-weight:400;opacity:.7;">(${items.length})</span>
            </div>
            ${items.length ? `<div class="bbz-mini-list">${items.join("")}</div>` : ""}
          </div>`;
      };

      const hasRadarItems = radarNever.length + radarCold.length + radarOverdue.length + radarOk.length > 0;

      const radarHtml = hasRadarItems ? (
        radarZoneHtml("Nie kontaktiert", "var(--red)",
          radarNever.map(f => radarItem(f, "noch kein Kontakt", "A-Kunde — noch kein Kontakt erfasst", "var(--red)", true))) +
        radarZoneHtml("Eingeschlafen (>360 Tage)", "var(--amber)",
          radarCold.map(f => {
            const lastDate = helpers.toDate(f.latestActivity);
            const months = (today.getFullYear() - lastDate.getFullYear()) * 12
                         + (today.getMonth() - lastDate.getMonth());
            return radarItem(f, `seit ${months} Monat${months !== 1 ? "en" : ""}`,
              `Letzter Kontakt: ${helpers.formatDate(f.latestActivity)} (vor ${months} Monaten)`, "var(--amber)");
          })) +
        radarZoneHtml("Überfällige Tasks", "var(--muted)",
          radarOverdue.map(f => {
            const cnt = f.tasks.filter(t => t.isOpen && t.isOverdue).length;
            return radarItem(f, `${cnt} Task${cnt !== 1 ? "s" : ""} überfällig`,
              `${cnt} überfällige Task(s)`, "var(--muted)");
          })) +
        radarZoneHtml("On Track ✓", "var(--green)",
          radarOk.map(f => {
            const lastDate = helpers.toDate(f.latestActivity);
            const days = lastDate ? Math.floor((today - lastDate) / 86400000) : null;
            const meta = days !== null ? `vor ${days} Tag${days !== 1 ? "en" : ""}` : "kürzlich";
            return radarItem(f, meta, `On Track — letzter Kontakt ${helpers.formatDate(f.latestActivity)}`, "var(--green)");
          }), false)
      ) : `<div style="text-align:center;padding:16px 0;color:var(--subtle);font-size:13px;">Keine A/B-Kunden erfasst.</div>`;

      // Aktiver Firma-Filter: Reset-Banner
      const firmaFilterBanner = activeRadarFirm
        ? `<div style="background:var(--blue-light);border:1px solid #a8c8e0;border-radius:8px;padding:7px 11px;font-size:12px;margin-bottom:8px;display:flex;align-items:center;justify-content:space-between;">
             <span>Gefiltert: <strong>${helpers.escapeHtml(activeRadarFirm)}</strong></span>
             <button class="bbz-link" style="font-size:12px;" data-action="history-firma-filter" data-firm-title="${helpers.escapeHtml(activeRadarFirm)}">× aufheben</button>
           </div>` : "";

      return `
        <div>
          <div class="bbz-kpis">
            <!-- Aktivitäten mit Zeitfenster-Chips -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Aktivitäten</div>
              <div class="bbz-kpi-value">${totalEntries}</div>
              <div style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${chipZeit("Heute", "today", cntToday)}
                ${chipZeit("Woche", "week", cntWeek)}
                ${chipZeit("Monat", "month", cntMonth)}
                <button class="bbz-kpi-chip ${!filters.zeitfenster ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="history-zeitfenster" data-value="">Alle</button>
              </div>
              <div style="margin-top:6px;font-size:12px;">${trendHtml}</div>
            </div>
            <!-- Kontaktart-Chips -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Kontaktart</div>
              <div class="bbz-kpi-value">${filters.kontaktart ? helpers.escapeHtml(filters.kontaktart) : "—"}</div>
              <div style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${allKontaktart.map(chipKontaktart).join("")}
                ${filters.kontaktart ? `<button class="bbz-kpi-chip" data-action="kpi-filter" data-scope="history-kontaktart" data-value="">Alle</button>` : ""}
              </div>
            </div>
            <!-- Lead BBZ-Chips -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Lead BBZ</div>
              <div class="bbz-kpi-value">${filters.leadbbz ? helpers.escapeHtml(filters.leadbbz) : "—"}</div>
              <div style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${allLeadbbz.map(chipLead).join("")}
                ${filters.leadbbz ? `<button class="bbz-kpi-chip" data-action="kpi-filter" data-scope="history-leadbbz" data-value="">Alle</button>` : ""}
              </div>
            </div>
            <!-- Sichtbar -->
            ${this.kpiBlock("Sichtbar", rows.length, rows.length < totalEntries ? `von ${totalEntries} gefiltert` : "alle", rows.length < totalEntries ? "warn" : "")}
          </div>

          <!-- Mobile Tab-Bar: nur sichtbar unter 920px -->
          <div class="bbz-history-tab-bar" id="history-tab-bar">
            <button class="bbz-history-tab-btn active" id="history-tab-timeline"
              onclick="document.getElementById('history-split').classList.remove('show-radar');
                       document.getElementById('history-tab-timeline').classList.add('active');
                       document.getElementById('history-tab-radar').classList.remove('active');">
              📅 Aktivitäten
            </button>
            <button class="bbz-history-tab-btn" id="history-tab-radar"
              onclick="document.getElementById('history-split').classList.add('show-radar');
                       document.getElementById('history-tab-radar').classList.add('active');
                       document.getElementById('history-tab-timeline').classList.remove('active');">
              🔍 Pflege-Radar
            </button>
          </div>

          <div class="bbz-grid bbz-grid-70-30 bbz-history-split" id="history-split">
            <!-- Links: Timeline oder Pflege-Vollansicht -->
            <section class="bbz-section">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">${filters.radarMode ? "Pflege A/B" : "Aktivitäten"}</div>
                  <div class="bbz-section-subtitle">${filters.radarMode ? `${radarNever.length + radarCold.length + radarOverdue.length} mit Handlungsbedarf · ${radarOk.length} On Track ✓` : filters.groupBy === "firm" ? "Gruppiert nach Firma" : "Chronologische Timeline"}</div>
                </div>
                <div style="display:flex;gap:6px;align-items:center;">
                  <!-- Tab-Bar -->
                  <div style="display:flex;border:1px solid var(--line);border-radius:9px;overflow:hidden;background:var(--panel-2);">
                    <button class="bbz-button" style="height:32px;font-size:12px;border:none;border-radius:0;padding:0 10px;${!filters.radarMode ? "background:var(--panel);color:var(--text);font-weight:700;" : "background:none;color:var(--muted);"}"
                      data-action="kpi-filter" data-scope="history-radar" ${!filters.radarMode ? "disabled" : ""}>
                      Aktivitäten
                    </button>
                    <button class="bbz-button" style="height:32px;font-size:12px;border:none;border-radius:0;padding:0 10px;${filters.radarMode ? "background:var(--panel);color:var(--text);font-weight:700;" : "background:none;color:var(--muted);"}"
                      data-action="kpi-filter" data-scope="history-radar" ${filters.radarMode ? "disabled" : ""}>
                      Pflege A/B ${(radarNever.length + radarCold.length + radarOverdue.length) > 0 ? `<span style="background:var(--red-light);color:var(--red);border-radius:999px;padding:1px 6px;font-size:11px;margin-left:4px;">${radarNever.length + radarCold.length + radarOverdue.length}</span>` : ""}
                    </button>
                  </div>
                  ${!filters.radarMode ? `
                  <select class="bbz-select" style="height:32px;font-size:12px;" data-filter="history-groupby">
                    <option value="date" ${filters.groupBy === "date" ? "selected" : ""}>📅 Nach Datum</option>
                    <option value="firm" ${filters.groupBy === "firm" ? "selected" : ""}>🏢 Nach Firma</option>
                  </select>
                  <button class="bbz-button bbz-button-primary" style="height:32px;font-size:12px;"
                    ${activeFirmHasNoContacts
                      ? `disabled title="Zuerst einen Kontakt bei ${helpers.escapeHtml(activeFirm.title)} erfassen"`
                      : `data-action="open-history-form"`}>+ Aktivität</button>` : ""}
                </div>
              </div>
              <div class="bbz-section-body">
                ${filters.radarMode ? `
                <!-- Pflege-Vollansicht -->
                <div style="margin-bottom:10px;">
                  <input class="bbz-input" style="width:100%;" data-filter="history-search" type="text"
                    placeholder="Suche nach Firma ..." value="${helpers.escapeHtml(filters.search)}" />
                </div>
                <div class="bbz-table-wrap">
                  <table class="bbz-table">
                    <thead><tr>
                      <th></th>
                      <th>Firma</th>
                      <th>Klassifizierung</th>
                      <th>Pflege-Grund</th>
                      <th>Letzte Aktivität</th>
                      <th>Nächste Deadline</th>
                      <th></th>
                    </tr></thead>
                    <tbody>
                      ${(() => {
                        const signalPriority = { overdue: 0, never: 1, cold: 2 };
                        const search = filters.search.trim().toLowerCase();
                        const allRadarRows = [...radarNever, ...radarOverdue, ...radarCold.filter(f => !radarOverdue.includes(f))]
                          .filter((f, i, arr) => arr.findIndex(x => x.id === f.id) === i) // deduplicate
                          .filter(f => !search || helpers.textIncludes(f.title, search))
                          .sort((a, b) => {
                            const pa = signalPriority[helpers.firmSignal(a)] ?? 9;
                            const pb = signalPriority[helpers.firmSignal(b)] ?? 9;
                            if (pa !== pb) return pa - pb;
                            return helpers.compareDateAsc(a.latestActivity, b.latestActivity);
                          });
                        if (!allRadarRows.length) return `<tr><td colspan="7">${ui.emptyBlock("Keine Pflege-Fälle gefunden.")}</td></tr>`;
                        return allRadarRows.map(firm => {
                          const sig = helpers.firmSignal(firm);
                          const dot = sig === "overdue"
                            ? `<span class="bbz-signal bbz-signal-red" title="Überfällige Tasks"></span>`
                            : `<span class="bbz-signal bbz-signal-amber"></span>`;
                          const rowClass = sig === "overdue" ? "bbz-row-alert" : "bbz-row-cold";
                          const lastDate = helpers.toDate(firm.latestActivity);
                          const months = lastDate
                            ? (today.getFullYear() - lastDate.getFullYear()) * 12 + (today.getMonth() - lastDate.getMonth())
                            : null;
                          const grundHtml = sig === "never"
                            ? `<span style="color:var(--red);font-weight:600;">🔴 Nie kontaktiert</span>`
                            : sig === "cold"
                            ? `<span style="color:var(--amber);font-weight:600;">🟡 Seit ${months} Monat${months !== 1 ? "en" : ""} still</span>`
                            : `<span style="color:var(--muted);font-weight:600;">⚠️ ${firm.tasks.filter(t => t.isOpen && t.isOverdue).length} Task${firm.tasks.filter(t => t.isOpen && t.isOverdue).length !== 1 ? "s" : ""} überfällig</span>`;
                          const hasContacts = firm.contacts.length > 0;
                          return `
                            <tr class="${rowClass}">
                              <td style="width:28px;">${dot}</td>
                              <td><a class="bbz-link" data-action="open-firm" data-id="${firm.id}">${helpers.escapeHtml(firm.title)}</a></td>
                              <td>${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                              <td>${grundHtml}</td>
                              <td>${firm.latestActivity ? `<span title="${helpers.formatDate(firm.latestActivity)}">${helpers.relativeDate(firm.latestActivity)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                              <td class="${firm.nextDeadline && helpers.isOverdue(firm.nextDeadline) ? "bbz-danger" : ""}">${firm.nextDeadline ? helpers.relativeDate(firm.nextDeadline) : '<span class="bbz-muted">—</span>'}</td>
                              <td style="white-space:nowrap;">
                                ${sig === "never"
                                  ? `<button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;"
                                       data-action="open-task-form" data-firm-id="${firm.id}">+ Task</button>`
                                  : hasContacts
                                  ? `<button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;"
                                       data-action="open-history-form" data-firm-id="${firm.id}">+ Aktivität</button>`
                                  : ""}
                              </td>
                            </tr>`;
                        }).join("");
                      })()}
                    </tbody>
                  </table>
                </div>` : `
                <!-- Timeline -->
                <div class="bbz-filters-2" style="display:grid;grid-template-columns:1fr auto;gap:10px;margin-bottom:10px;align-items:center;">
                  <input class="bbz-input" data-filter="history-search" type="text" placeholder="Suche nach Kontakt, Firma, Notizen ..." value="${helpers.escapeHtml(filters.search)}" />
                  ${activeFilterBadges ? `<div style="display:flex;gap:4px;flex-wrap:wrap;">${activeFilterBadges}</div>` : "<div></div>"}
                </div>
                ${firmaFilterBanner}
                ${timelineHtml}`}
              </div>
            </section>

            <!-- Rechts: Pflege-Radar -->
            <section class="bbz-section">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">Pflege-Radar</div>
                  <div class="bbz-section-subtitle">${radarNever.length + radarCold.length + radarOverdue.length} Handlungsbedarf · ${radarOk.length} On Track ✓</div>
                </div>
                ${(radarNever.length + radarCold.length + radarOverdue.length) > 0
                  ? `<button class="bbz-link" style="font-size:12px;" data-action="kpi-filter" data-scope="history-radar">Alle →</button>`
                  : ""}
              </div>
              <div class="bbz-section-body">
                ${filters.radarMode ? `<div style="font-size:13px;color:var(--muted);text-align:center;padding:12px 0;">Vollansicht aktiv →</div>` : radarHtml}
              </div>
            </section>
          </div>
        </div>
      `;
    },

    events() {
      const filters = state.filters.events;
      const evtSortBy  = filters.sortBy  || "contactName";
      const evtSortDir = filters.sortDir === "asc" ? 1 : -1;
      const allGroups  = state.enriched.events;

      // Segment-Chips (auf Events-Kachel)
      const segChip = (label, val) => {
        const active = filters.segment === val;
        const cnt = val === ""
          ? allGroups.reduce((s, g) => s + g.contacts.length, 0)
          : allGroups.reduce((s, g) => s + g.contacts.filter(c => c.segment.startsWith(val)).length, 0);
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="events-segment" data-value="${helpers.escapeHtml(val)}">${helpers.escapeHtml(label || "Alle")} <span>${cnt}</span></button>`;
      };

      // Event-Filter-Chips auf Anmeldungen-Kachel — ein Chip pro Event-Gruppe
      const evtChip = (group) => {
        const active = filters.selectedEvent === group.name;
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}"
          data-action="kpi-filter" data-scope="events-selected" data-value="${helpers.escapeHtml(group.name)}"
          title="${helpers.escapeHtml(group.name)}">
          ${helpers.escapeHtml(group.name)} <span>${group.contactCount}</span>
        </button>`;
      };

      // Gruppen filtern und aufbereiten
      const groups = allGroups.map(group => {
        // selectedEvent-Filter: wenn gesetzt, nur diese Gruppe zeigen
        if (filters.selectedEvent && group.name !== filters.selectedEvent) return null;
        const filtered = group.contacts.filter(item => {
          const search = filters.search.trim().toLowerCase();
          const searchMatch = !search || [group.name, item.contactName, item.firmTitle, item.rolle, item.funktion].some(v => helpers.textIncludes(v, search));
          const segMatch = !filters.segment || item.segment.startsWith(filters.segment);
          return searchMatch && segMatch && (!filters.onlyWithOpenTasks || item.openTasksCount > 0);
        });
        if (!filtered.length) return null;
        const sorted = [...filtered].sort((a, b) => {
          if (evtSortBy === "firmTitle") return String(a.firmTitle||"").localeCompare(String(b.firmTitle||""), "de") * evtSortDir;
          return String(a.contactName||"").localeCompare(String(b.contactName||""), "de") * evtSortDir;
        });
        const cntA = group.contacts.filter(c => c.segment.startsWith("A")).length;
        const cntB = group.contacts.filter(c => c.segment.startsWith("B")).length;
        const cntC = group.contacts.filter(c => c.segment.startsWith("C")).length;
        return { ...group, contacts: sorted, cntA, cntB, cntC };
      }).filter(Boolean);

      const totalGroups   = allGroups.length;
      const totalContacts = allGroups.reduce((sum, e) => sum + e.contactCount, 0);

      return `
        <div>
          <div class="bbz-kpis">

            <!-- Kachel 1: Events mit Segment-Chips -->
            <div class="bbz-kpi bbz-kpi-blue">
              <div class="bbz-kpi-label">Events</div>
              <div class="bbz-kpi-value">${totalGroups}</div>
              <div class="bbz-kpi-chips" style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${segChip("A", "A")}${segChip("B", "B")}${segChip("C", "C")}${segChip("Alle", "")}
              </div>
            </div>

            <!-- Kachel 2: Anmeldungen mit Event-Filter-Chips -->
            <div class="bbz-kpi bbz-kpi-blue">
              <div class="bbz-kpi-label">Anmeldungen</div>
              <div class="bbz-kpi-value">${totalContacts}</div>
              <div class="bbz-kpi-chips" style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${allGroups.map(g => evtChip(g)).join("")}
                ${filters.selectedEvent ? `<button class="bbz-kpi-chip" data-action="kpi-filter" data-scope="events-selected" data-value="" style="opacity:0.6;">× Alle</button>` : ""}
              </div>
            </div>

            <!-- Kachel 3: Event Nachbearbeitung -->
            <div class="bbz-kpi bbz-kpi-blue" style="display:flex;flex-direction:column;justify-content:space-between;">
              <div>
                <div class="bbz-kpi-label">Event Nachbearbeitung</div>
                <div class="bbz-kpi-value" style="font-size:22px;margin-top:4px;letter-spacing:-0.03em;">Vergangene<br>Teilnahmen</div>
              </div>
              <button class="bbz-button bbz-button-primary" style="margin-top:10px;height:34px;font-size:12px;width:100%;"
                data-action="open-batch-event" data-event-name="" data-mode="eventhistory">
                Vergangene Eventteilnahmen pflegen
              </button>
            </div>

          </div>
          <section class="bbz-section">
            <div class="bbz-section-header">
              <div>
                <div class="bbz-section-title">Events</div>
                <div class="bbz-section-subtitle">${filters.selectedEvent ? `Gefiltert: ${helpers.escapeHtml(filters.selectedEvent)}` : "Anmeldungen nach Event-Kategorie"}</div>
              </div>
            </div>
            <div class="bbz-section-body">
              <div class="bbz-filters-3">
                <input class="bbz-input" data-filter="events-search" type="text" placeholder="Suche nach Kontakt, Firma ..." value="${helpers.escapeHtml(filters.search)}" />
                <select class="bbz-select" data-filter="events-sortby">
                  <option value="contactName" ${(filters.sortBy||"contactName") === "contactName" ? "selected" : ""}>Sortierung: Kontakt A–Z</option>
                  <option value="firmTitle"   ${(filters.sortBy||"") === "firmTitle"   ? "selected" : ""}>Sortierung: Firma A–Z</option>
                </select>
                <label class="bbz-checkbox"><input type="checkbox" data-filter="events-open" ${filters.onlyWithOpenTasks ? "checked" : ""} /> Nur mit offenen Tasks</label>
              </div>
              ${groups.length ? `<div class="bbz-cockpit-stack">${groups.map(group => {
                const segBadges = [
                  group.cntA ? `<span class="bbz-pill bbz-pill-a" style="font-size:11px;padding:1px 6px;">A ${group.cntA}</span>` : "",
                  group.cntB ? `<span class="bbz-pill bbz-pill-b" style="font-size:11px;padding:1px 6px;">B ${group.cntB}</span>` : "",
                  group.cntC ? `<span class="bbz-pill bbz-pill-c" style="font-size:11px;padding:1px 6px;">C ${group.cntC}</span>` : ""
                ].filter(Boolean).join("");
                return `
                <section class="bbz-section" style="box-shadow:none;">
                  <div class="bbz-section-header">
                    <div>
                      <div class="bbz-section-title" style="font-size:15px;font-weight:700;letter-spacing:-0.02em;display:flex;align-items:center;gap:8px;">
                        ${helpers.escapeHtml(group.name)}
                        <span style="display:flex;gap:4px;align-items:center;">${segBadges}</span>
                      </div>
                      <div class="bbz-section-subtitle">${group.contacts.length} Kontakte · ${group.contacts.reduce((sum, c) => sum + c.openTasksCount, 0)} offene Tasks</div>
                    </div>
                    <button class="bbz-button bbz-button-primary" style="height:30px;font-size:12px;"
                      data-action="open-batch-event" data-event-name="${helpers.escapeHtml(group.name)}" data-mode="anmelden">
                      + Kontakte hinzufügen
                    </button>
                  </div>
                  <div class="bbz-section-body">
                    <div class="bbz-table-wrap">
                      <table class="bbz-table">
                        <thead><tr>
                          <th></th><th>Kontakt</th><th>Firma</th><th>Segment</th><th>Funktion / Rolle</th><th>Letzte Aktivität</th><th>Tasks</th><th></th>
                        </tr></thead>
                        <tbody>
                          ${group.contacts.map(item => {
                            const segBadge = item.segment
                              ? `<span class="${helpers.firmBadgeClass(item.segment)}">${helpers.escapeHtml(item.segment)}</span>`
                              : '<span class="bbz-muted">—</span>';
                            return `
                              <tr>
                                <td style="width:36px;">${helpers.avatarHtml({ vorname: (item.contactName||"").split(" ")[0]||"", nachname: (item.contactName||"").split(" ").slice(-1)[0]||"" })}</td>
                                <td><a class="bbz-link" data-action="open-contact" data-id="${item.contactId}">${helpers.escapeHtml(item.contactName)}</a><div class="bbz-subtext">${item.email1 ? helpers.escapeHtml(item.email1) : "—"}</div></td>
                                <td>${item.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${item.firmId}">${helpers.escapeHtml(item.firmTitle||"Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
                                <td>${segBadge}</td>
                                <td>${helpers.escapeHtml(helpers.joinNonEmpty([item.funktion, item.rolle], " · "))||'<span class="bbz-muted">—</span>'}</td>
                                <td>${item.latestHistoryDate ? `<span title="${helpers.formatDate(item.latestHistoryDate)}">${helpers.relativeDate(item.latestHistoryDate)}</span>${item.latestHistoryType ? `<div class="bbz-subtext">${helpers.escapeHtml(item.latestHistoryType)}</div>` : ""}` : '<span class="bbz-muted">—</span>'}</td>
                                <td>${item.openTasksCount > 0 ? `<span class="bbz-status-chip bbz-status-open">${item.openTasksCount} offen</span>` : '<span class="bbz-muted">—</span>'}</td>
                                <td style="white-space:nowrap;"><button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;" data-action="open-history-form" data-contact-id="${item.contactId}">+ Aktivität</button></td>
                              </tr>`;
                          }).join("")}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </section>`;
              }).join("")}</div>`
              : ui.emptyBlock("Keine Event-Daten für die aktuelle Filterung gefunden.")}
            </div>
          </section>
        </div>
      `;
    }
  };

  const controller = {
    async init() {
      ui.init();
      ui.renderShell();
      ui.setMessage("");
      ui.renderView(ui.loadingBlock("Authentifizierung wird vorbereitet ..."));

      try {
        ui.setLoading(true);
        await api.initAuth();

        if (state.auth.isAuthenticated) {
          // acquireToken() wird von graphRequest() intern aufgerufen — kein separater Call nötig
          await Promise.all([api.loadAll(), api.loadColumnChoices()]);
          ui.setMessage("Anmeldung erkannt. Daten wurden geladen.", "success");
        } else {
          ui.setMessage("Bitte anmelden, um die SharePoint-Listen ueber Microsoft Graph zu laden.", "warning");
        }
      } catch (error) {
        console.error(error);
        state.meta.lastError = error;
        ui.setMessage(`Fehler beim Initialisieren: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleLogin() {
      try {
        if (!state.auth.isReady) { ui.setMessage("Authentifizierung ist noch nicht bereit. Bitte Seite neu laden.", "warning"); return; }
        ui.setLoading(true);
        ui.setMessage("");
        await api.login();
        // Choices und Daten beim Login parallel laden
        await Promise.all([api.loadAll(), api.loadColumnChoices()]);
        ui.setMessage("Anmeldung erfolgreich. Daten wurden geladen.", "success");
      } catch (error) {
        console.error(error);
        ui.setMessage(`Anmeldung fehlgeschlagen: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleRefresh() {
      if (!state.auth.isReady) { ui.setMessage("Authentifizierung ist noch nicht bereit.", "warning"); return; }
      if (!state.auth.isAuthenticated) { ui.setMessage("Bitte zuerst anmelden.", "warning"); return; }
      try {
        ui.setLoading(true);
        ui.setMessage("");
        await api.acquireToken();
        // Refresh: Choices ebenfalls neu laden — SP-Schema könnte sich geändert haben
        await Promise.all([api.loadAll(), api.loadColumnChoices()]);
        ui.setMessage("Daten erfolgreich neu geladen.", "success");
      } catch (error) {
        console.error(error);
        ui.setMessage(`Fehler beim Laden: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    // FIX 2d: Modal oeffnen
    openContactForm(itemId = null, prefillFirmId = null) {
      state.modal = {
        type: "contact",
        mode: itemId ? "edit" : "create",
        payload: { itemId, prefillFirmId }
      };
      this.render();
    },

    openFirmForm(firmId = null) {
      state.modal = {
        type: "firm",
        mode: firmId ? "edit" : "create",
        payload: { firmId: firmId ? Number(firmId) : null }
      };
      this.render();
    },

    async handleFirmModalSubmit(form, mode, itemId) {
      const fd = new FormData(form);

      if (!fd.get("title")?.trim()) {
        ui.setMessage("Firmenname ist ein Pflichtfeld.", "error");
        return;
      }

      const fields = {
        Title: fd.get("title").trim(),
        VIP:   form.querySelector("[name='vip']")?.checked ?? false,
        // Immer senden — null löscht den Wert in SP
        Adresse:        fd.get("adresse")?.trim()      || null,
        PLZ:            fd.get("plz")?.trim()           || null,
        Ort:            fd.get("ort")?.trim()            || null,
        Land:           fd.get("land")?.trim()           || null,
        Hauptnummer:    fd.get("hauptnummer")?.trim()    || null,
        Klassifizierung: fd.get("klassifizierung")      || "",
      };


      ui.setLoading(true);
      ui.setMessage("");

      try {
        if (mode === "create") {
          await api.postItem(SCHEMA.firms.listTitle, fields);
          ui.setMessage("Firma wurde erfolgreich angelegt.", "success");
        } else {
          if (!itemId) throw new Error("itemId fehlt für PATCH.");
          await api.patchItem(SCHEMA.firms.listTitle, Number(itemId), fields);
          ui.setMessage("Firma wurde erfolgreich gespeichert.", "success");
        }
        await api.loadAll();
        this.closeModal();
      } catch (error) {
        console.error("handleFirmModalSubmit Fehler:", error);
        let msg = error.message || "Unbekannter Fehler";
        if (msg.includes("400")) msg = "Fehler 400: Ungültige Felddaten.";
        if (msg.includes("403")) msg = "Fehler 403: Keine Schreibberechtigung.";
        ui.setMessage(msg, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleDeleteContact(id, name) {
      if (!confirm(`Kontakt "${name}" wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.`)) return;
      try {
        ui.setLoading(true);
        await api.deleteItem(SCHEMA.contacts.listTitle, Number(id));
        ui.setMessage(`Kontakt "${name}" wurde gelöscht.`, "success");
        state.selection.contactId = null;
        state.filters.route = "contacts";
        await api.loadAll();
      } catch (error) {
        console.error("handleDeleteContact:", error);
        ui.setMessage(`Fehler beim Löschen: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleDeleteFirm(id, name) {
      if (!confirm(`Firma "${name}" wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.`)) return;
      try {
        ui.setLoading(true);
        await api.deleteItem(SCHEMA.firms.listTitle, Number(id));
        ui.setMessage(`Firma "${name}" wurde gelöscht.`, "success");
        state.selection.firmId = null;
        state.filters.route = "firms";
        await api.loadAll();
      } catch (error) {
        console.error("handleDeleteFirm:", error);
        ui.setMessage(`Fehler beim Löschen: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    // FIX 2e: Modal schliessen
    closeModal() {
      state.modal = null;
      this.render();
    },

    // Write-Layer: Kontakt speichern (create oder edit)
    async handleModalSubmit(form, mode, itemId) {
      // FormData.entries() gibt bei gleichnamigen Checkboxen nur den letzten Wert zurück.
      // Deshalb getAll() für Multi-Choice-Felder verwenden.
      const fd = new FormData(form);

      const raw = {
        nachname:      fd.get("nachname") || "",
        vorname:       fd.get("vorname") || "",
        anrede:        fd.get("anrede") || "",
        firmaLookupId: fd.get("firmaLookupId") || "",
        funktion:      fd.get("funktion") || "",
        rolle:         fd.get("rolle") || "",
        email1:        fd.get("email1") || "",
        email2:        fd.get("email2") || "",
        direktwahl:    fd.get("direktwahl") || "",
        mobile:        fd.get("mobile") || "",
        geburtstag:    fd.get("geburtstag") || "",
        leadbbz0:      fd.get("leadbbz0") || "",
        kommentar:     fd.get("kommentar") || "",
        // Multi-Choice: getAll() sammelt alle checked Werte
        sgf:           fd.getAll("sgf"),
        event:         fd.getAll("event"),
        eventhistory:  fd.getAll("eventhistory"),
        // Checkbox Archiviert
        archiviert:    form.querySelector("[name='archiviert']")?.checked ?? false
      };

      // Pflichtfeld-Validierung
      if (!raw.nachname.trim()) {
        ui.setMessage("Nachname ist ein Pflichtfeld.", "error");
        return;
      }
      if (!raw.firmaLookupId) {
        ui.setMessage("Bitte eine Firma zuweisen.", "error");
        return;
      }

      // Pflichtfelder — immer senden
      const fields = {
        Title:         raw.nachname.trim(),
        FirmaLookupId: Number(raw.firmaLookupId),
        // Archiviert immer senden — auch false, sonst kann ein archivierter Kontakt nicht reaktiviert werden
        Archiviert:    raw.archiviert
      };

      // Einzelwahl-Choice-Felder — leer = "" (nicht null, SP Choice-Felder akzeptieren "" zuverlässiger)
      fields.Anrede   = raw.anrede   || "";
      fields.Rolle    = raw.rolle    || "";
      fields.Leadbbz0 = raw.leadbbz0 || "";

      // Optionaler Text — immer senden, leer = null zum Löschen in SP
      fields.Vorname    = raw.vorname.trim()    || null;
      fields.Funktion   = raw.funktion.trim()   || null;
      fields.Kommentar  = raw.kommentar.trim()  || null;
      fields.Email1     = raw.email1.trim()     || null;
      fields.Email2     = raw.email2.trim()     || null;
      fields.Direktwahl = raw.direktwahl.trim() || null;
      fields.Mobile     = raw.mobile.trim()     || null;

      // Datum — leer = null zum Löschen
      fields.Geburtstag = raw.geburtstag.trim() ? raw.geburtstag.trim() + "T00:00:00Z" : null;

      // Multi-Choice — @odata.type + Array (befüllen) oder @odata.type + [] (leeren)
      // BESTÄTIGT: @odata.type + Array mit Werten → ✅
      // OFFEN: @odata.type + [] zum Leeren → zu testen
      fields["SGF@odata.type"]          = "Collection(Edm.String)";
      fields["SGF"]                     = raw.sgf;
      fields["Event@odata.type"]        = "Collection(Edm.String)";
      fields["Event"]                   = raw.event;
      fields["Eventhistory@odata.type"] = "Collection(Edm.String)";
      fields["Eventhistory"]            = raw.eventhistory;



      ui.setLoading(true);
      ui.setMessage("");

      try {
        if (mode === "create") {
          // SharePoint Graph: POST akzeptiert nur Title + Lookup-Felder zuverlässig.
          // Alle weiteren Felder müssen per separatem PATCH auf die neue Item-ID geschrieben werden.
          // BESTÄTIGT: POST mit vollem fields-Objekt speichert nur Title.
          const createFields = {
            Title:         fields.Title,
            FirmaLookupId: fields.FirmaLookupId
          };
          const created = await api.postItem(SCHEMA.contacts.listTitle, createFields);
          const newItemId = created?.id || created?.fields?.id;
          if (!newItemId) throw new Error("Neue Item-ID fehlt im POST-Response.");

          // Restliche Felder per PATCH nachschreiben
          const patchFields = { ...fields };
          delete patchFields.Title;
          delete patchFields.FirmaLookupId;
          if (Object.keys(patchFields).length > 0) {
            await api.patchItem(SCHEMA.contacts.listTitle, Number(newItemId), patchFields);
          }

          ui.setMessage("Kontakt wurde erfolgreich angelegt.", "success");
        } else {
          if (!itemId) throw new Error("itemId fehlt für PATCH.");
          await api.patchItem(SCHEMA.contacts.listTitle, Number(itemId), fields);
          ui.setMessage("Kontakt wurde erfolgreich gespeichert.", "success");
        }

        await api.loadAll();
        this.closeModal();
      } catch (error) {
        console.error("handleModalSubmit Fehler:", error);

        // Vollständigen Graph-Fehlertext extrahieren für sauberes Debugging
        let msg = error.message || "Unbekannter Fehler";
        let detail = "";
        try {
          // Graph-Fehler haben oft JSON im message-String
          const match = msg.match(/\{.*\}/s);
          if (match) {
            const parsed = JSON.parse(match[0]);
            detail = parsed?.error?.message || parsed?.message || "";
          }
        } catch { /* ignore parse error */ }

        if (msg.includes("400")) msg = `Fehler 400: Ungültige Felddaten.${detail ? " " + detail : " Bitte Konsole prüfen."}`;
        if (msg.includes("403")) msg = "Fehler 403: Keine Schreibberechtigung auf diese Liste.";
        if (msg.includes("409")) msg = "Fehler 409: Konflikt — Eintrag wurde zwischenzeitlich geändert.";

        ui.setMessage(msg, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    openBatchEventForm(eventName, mode = "anmelden") {
      state.modal = {
        type: "batch-event",
        payload: {
          eventName,
          mode,
          filterSegment: "",
          filterLeadbbz: "",
          filterSgf: "",
          filterSearch: "",
          selected: [],
          previewContacts: [],
          selectedHistoryCategory: ""
        }
      };
      this.render();
    },

    async handleBatchEventSubmit(form) {
      const mode = form.dataset.mode || "anmelden";
      const isEventhistory = mode === "eventhistory";
      // Für eventhistory: aktive Kategorie aus Payload lesen (Dropdown-Auswahl)
      const eventName = isEventhistory
        ? (state.modal?.payload?.selectedHistoryCategory || "")
        : (form.dataset.eventName || "");

      let selectedIds = [];
      // Direkt aus State lesen — zuverlässiger als hidden Input,
      // da DOM-only Updates beim Checkbox-Klick keinen hidden Input pflegen
      selectedIds = state.modal?.payload?.selected || [];
      if (!selectedIds.length) {
        // Fallback: hidden Input (rückwärtskompatibel)
        try { selectedIds = JSON.parse(form.querySelector("[name='selectedIds']")?.value || "[]"); } catch { /* ignore */ }
      }

      if (!eventName) { ui.setMessage("Bitte eine Kategorie wählen.", "error"); return; }
      if (!selectedIds.length) { ui.setMessage("Keine Kontakte ausgewählt.", "error"); return; }

      ui.setLoading(true);
      ui.setMessage("");

      let ok = 0, fail = 0;
      try {
        const results = await Promise.allSettled(selectedIds.map(async cid => {
          const contact = state.enriched.contacts.find(c => c.id === cid);
          if (!contact) throw new Error(`Kontakt ${cid} nicht gefunden`);

          const currentEvent     = helpers.toArray(contact.event);
          const currentEventHist = helpers.toArray(contact.eventhistory);

          const patchFields = {};
          if (isEventhistory) {
            // Eventhistory-Feld: Kategorie additiv hinzufügen
            if (!currentEventHist.includes(eventName)) {
              patchFields["Eventhistory@odata.type"] = "Collection(Edm.String)";
              patchFields["Eventhistory"] = [...currentEventHist, eventName];
            }
          } else {
            // Event-Feld: Kategorie additiv hinzufügen
            if (!currentEvent.includes(eventName)) {
              patchFields["Event@odata.type"] = "Collection(Edm.String)";
              patchFields["Event"] = [...currentEvent, eventName];
            }
          }
          // Nichts zu tun wenn Flag bereits gesetzt
          if (Object.keys(patchFields).length === 0) return;
          await api.patchItem(SCHEMA.contacts.listTitle, Number(cid), patchFields);
        }));

        results.forEach(r => r.status === "fulfilled" ? ok++ : (fail++, console.error(r.reason)));
        await api.loadAll();
        this.closeModal();

        const fieldLabel = isEventhistory ? "Eventhistory" : "Event";
        const msg = `✓ ${fieldLabel} «${eventName}» für ${ok} Kontakt${ok !== 1 ? "e" : ""} gesetzt${fail > 0 ? ` — ${fail} Fehler (Konsole prüfen)` : ""}.`;
        ui.setMessage(msg, fail > 0 ? "error" : "success");
        if (fail === 0) setTimeout(() => ui.setMessage(""), 3000);

      } catch (error) {
        console.error("handleBatchEventSubmit:", error);
        ui.setMessage(`Fehler: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    openHistoryForm(contactId = null, firmId = null, itemId = null) {
      let prefillContactId = contactId;
      if (!prefillContactId && firmId) {
        const firm = dataModel.getFirmById(firmId);
        prefillContactId = firm?.contacts?.[0]?.id || null;
      }
      const mode = itemId ? "edit" : "create";
      state.modal = { type: "history", payload: { prefillContactId, mode, itemId } };
      this.render();
    },

    openTaskForm(contactId = null, firmId = null, itemId = null) {
      let prefillContactId = contactId;
      if (!prefillContactId && firmId) {
        const firm = dataModel.getFirmById(firmId);
        prefillContactId = firm?.contacts?.[0]?.id || null;
      }
      const mode = itemId ? "edit" : "create";
      state.modal = { type: "task", payload: { prefillContactId, mode, itemId } };
      this.render();
    },

    async handleDeleteHistory(id, title) {
      if (!confirm(`Aktivitaet "${title}" wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.`)) return;
      try {
        ui.setLoading(true);
        await api.deleteItem(SCHEMA.history.listTitle, Number(id));
        ui.setMessage(`Aktivitaet "${title}" wurde gelöscht.`, "success");
        await api.loadAll();
      } catch (error) {
        console.error("handleDeleteHistory:", error);
        ui.setMessage(`Fehler beim Löschen: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleDeleteTask(id, title) {
      if (!confirm(`Aufgabe "${title}" wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.`)) return;
      try {
        ui.setLoading(true);
        await api.deleteItem(SCHEMA.tasks.listTitle, Number(id));
        ui.setMessage(`Aufgabe "${title}" wurde gelöscht.`, "success");
        await api.loadAll();
      } catch (error) {
        console.error("handleDeleteTask:", error);
        ui.setMessage(`Fehler beim Löschen: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleHistoryModalSubmit(form) {
      const fd = new FormData(form);
      const mode = form.dataset.mode || "create";
      const itemId = Number(form.dataset.itemId || 0) || null;
      const kontaktLookupId = fd.get("kontaktLookupId") || "";
      const datum = fd.get("datum") || "";

      if (!kontaktLookupId) { ui.setMessage("Bitte einen Kontakt waehlen.", "error"); return; }
      if (!datum) { ui.setMessage("Datum ist ein Pflichtfeld.", "error"); return; }

      const kontaktart = fd.get("kontaktart") || "";
      const leadbbz = fd.get("leadbbz") || "";
      const notizen = fd.get("notizen") || "";
      const projektbezug = form.querySelector("[name='projektbezug']")?.checked ?? false;

      ui.setLoading(true);
      ui.setMessage("");
      try {
        if (mode === "edit") {
          if (!itemId) throw new Error("itemId fehlt fuer PATCH.");
          const patchFields = {
            Datum:      datum + "T00:00:00Z",
            Projektbezug: projektbezug,
            Kontaktart: kontaktart || "",
            Leadbbz:    leadbbz    || "",
            Notizen:    notizen.trim() || null,
          };
          await api.patchItem(SCHEMA.history.listTitle, itemId, patchFields);
          ui.setMessage("Aktivitaet wurde gespeichert.", "success");
        } else {
          // POST: nur Pflichtfelder, dann PATCH mit Rest
          const createFields = {
            Title: `Aktivitaet-${datum}`,
            NachnameLookupId: Number(kontaktLookupId)
          };
          const patchFields = { Datum: datum + "T00:00:00Z", Projektbezug: projektbezug };
          if (kontaktart) patchFields.Kontaktart = kontaktart;
          if (leadbbz) patchFields.Leadbbz = leadbbz;
          if (notizen.trim()) patchFields.Notizen = notizen.trim();
          const created = await api.postItem(SCHEMA.history.listTitle, createFields);
          const newId = created?.id || created?.fields?.id;
          if (!newId) throw new Error("Neue Item-ID fehlt im POST-Response.");
          await api.patchItem(SCHEMA.history.listTitle, Number(newId), patchFields);
          ui.setMessage("Aktivitaet wurde erfasst.", "success");
        }
        await api.loadAll();
        this.closeModal();
      } catch (error) {
        console.error("handleHistoryModalSubmit:", error);
        let msg = error.message || "Unbekannter Fehler";
        if (msg.includes("400")) msg = "Fehler 400: Ungueltige Felddaten. Bitte Konsole pruefen.";
        if (msg.includes("403")) msg = "Fehler 403: Keine Schreibberechtigung.";
        ui.setMessage(msg, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleTaskModalSubmit(form) {
      const fd = new FormData(form);
      const mode = form.dataset.mode || "create";
      const itemId = Number(form.dataset.itemId || 0) || null;
      const title = fd.get("title") || "";
      const kontaktLookupId = fd.get("kontaktLookupId") || "";

      if (!title.trim()) { ui.setMessage("Titel ist ein Pflichtfeld.", "error"); return; }
      if (!kontaktLookupId) { ui.setMessage("Bitte einen Kontakt waehlen.", "error"); return; }

      const deadline = fd.get("deadline") || "";
      const status = fd.get("status") || "";
      const leadbbz = fd.get("leadbbz") || "";

      ui.setLoading(true);
      ui.setMessage("");
      try {
        if (mode === "edit") {
          if (!itemId) throw new Error("itemId fehlt fuer PATCH.");
          const patchFields = {
            Title:    title.trim(),
            Deadline: deadline ? deadline + "T00:00:00Z" : null,
            Status:   status   || "",
            Leadbbz:  leadbbz  || "",
          };
          await api.patchItem(SCHEMA.tasks.listTitle, itemId, patchFields);
          ui.setMessage("Aufgabe wurde gespeichert.", "success");
        } else {
          const createFields = { Title: title.trim(), NameLookupId: Number(kontaktLookupId) };
          const patchFields = {};
          if (deadline) patchFields.Deadline = deadline + "T00:00:00Z";
          if (status) patchFields.Status = status;
          if (leadbbz) patchFields.Leadbbz = leadbbz;
          const created = await api.postItem(SCHEMA.tasks.listTitle, createFields);
          const newId = created?.id || created?.fields?.id;
          if (!newId) throw new Error("Neue Item-ID fehlt im POST-Response.");
          if (Object.keys(patchFields).length > 0) {
            await api.patchItem(SCHEMA.tasks.listTitle, Number(newId), patchFields);
          }
          ui.setMessage("Aufgabe wurde erstellt.", "success");
        }
        await api.loadAll();
        this.closeModal();
      } catch (error) {
        console.error("handleTaskModalSubmit:", error);
        let msg = error.message || "Unbekannter Fehler";
        if (msg.includes("400")) msg = "Fehler 400: Ungueltige Felddaten. Bitte Konsole pruefen.";
        if (msg.includes("403")) msg = "Fehler 403: Keine Schreibberechtigung.";
        ui.setMessage(msg, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleTaskStatusChange(taskId, newStatus) {
      if (!taskId || !newStatus) return;
      try {
        ui.setLoading(true);
        ui.setMessage("");
        await api.patchItem(SCHEMA.tasks.listTitle, taskId, { Status: newStatus });
        const isDone = !helpers.isOpenTask(newStatus);
        if (isDone) {
          ui.setMessage(`✓ Task als „${newStatus}" markiert.`, "success");
          // Auto-dismiss nach 2.5s
          setTimeout(() => { ui.setMessage(""); }, 2500);
        } else {
          ui.setMessage(`Status auf „${newStatus}" gesetzt.`, "success");
        }
        await api.loadAll();
      } catch (error) {
        console.error("handleTaskStatusChange:", error);
        ui.setMessage(`Fehler beim Status-Update: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    navigate(route) {
      state.filters.route = route;
      state.selection.firmId = null;
      state.selection.contactId = null;
      state.modal = null;
      state.filters.firms.radarMode = false;
      state.filters.history.radarMode = false;
      state.filters.events.segment = "";
      state.filters.events.selectedEvent = "";
      history.pushState({ route, firmId: null, contactId: null }, "", `#${route}`);
      window.scrollTo(0, 0);
      this.render();
    },

    openFirm(id) {
      state.selection.firmId = id;
      state.selection.contactId = null;
      state.filters.route = "firms";
      state.modal = null;
      history.pushState({ route: "firms", firmId: id, contactId: null }, "", `#firms-${id}`);
      window.scrollTo(0, 0);
      this.render();
    },

    openContact(id) {
      state.selection.contactId = id;
      state.filters.route = "contacts";
      state.modal = null;
      history.pushState({ route: "contacts", firmId: null, contactId: id }, "", `#contacts-${id}`);
      window.scrollTo(0, 0);
      this.render();
    },

    render() {
      ui.renderShell();
      ui.renderView(views.renderRoute());
    }
  };

  window._bbzApp = { state, api, helpers, SCHEMA, CONFIG, dataModel, controller };

  function startApp() { controller.init(); }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startApp, { once: true });
  } else {
    startApp();
  }
})();
