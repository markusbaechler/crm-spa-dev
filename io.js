/**
 * bbz CRM — io.js
 * Import / Export Modul (SheetJS-basiert, rein client-seitig)
 *
 * Einbindung in index.html:
 *   <script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
 *   <script src="./io.js"></script>
 *   (nach app.js einbinden, damit state / api / helpers / SCHEMA verfügbar sind)
 *
 * Öffentliche API (window.bbzIO):
 *   bbzIO.openExportModal(mode)   — mode: "contacts" | "event"
 *   bbzIO.openImportModal()
 *   bbzIO.downloadTemplate()      — leere Master-Vorlage herunterladen
 */

(() => {
  "use strict";

  // ─── Zugriff auf App-Interna (via IIFE-Scope nicht direkt möglich → window-Bridge) ──
  // app.js muss am Ende seiner IIFE folgendes registrieren:
  //   window._bbzApp = { state, api, helpers, SCHEMA, CONFIG, dataModel, controller };
  // Siehe Anleitung im README unten.

  function app() {
    const a = window._bbzApp;
    if (!a) throw new Error("bbzIO: window._bbzApp nicht gefunden. Bitte app.js anpassen (siehe io.js Anleitung).");
    return a;
  }

  // ─── Hilfsfunktionen ──────────────────────────────────────────────────────────────

  function today() {
    return new Date().toISOString().split("T")[0];
  }

  function filenameDate() {
    return today().replaceAll("-", "");
  }

  function escHtml(v) {
    return String(v ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;");
  }

  // Entfernt bestehende io-Modals sauber
  function removeModal(id) {
    document.getElementById(id)?.remove();
  }

  function showBanner(msg, type = "info") {
    const el = document.getElementById("global-message");
    if (!el) return;
    if (!msg) { el.className = "bbz-banner"; el.textContent = ""; return; }
    const cls = {
      success: "bbz-banner bbz-banner-success show",
      warning: "bbz-banner bbz-banner-warning show",
      error:   "bbz-banner bbz-banner-error show",
      info:    "bbz-banner bbz-banner-info show"
    };
    el.className = cls[type] || cls.info;
    el.textContent = msg;
    window.scrollTo(0, 0);
  }

  // ─── Felddefinitionen (Export) ────────────────────────────────────────────────────

  const CONTACT_FIELDS = [
    { key: "nachname",      label: "Nachname",        default: true },
    { key: "vorname",       label: "Vorname",          default: true },
    { key: "anrede",        label: "Anrede",           default: false },
    { key: "firmTitle",     label: "Firma",            default: true },
    { key: "funktion",      label: "Funktion",         default: true },
    { key: "email1",        label: "E-Mail 1",         default: true },
    { key: "email2",        label: "E-Mail 2",         default: false },
    { key: "direktwahl",    label: "Direktwahl",       default: true },
    { key: "mobile",        label: "Mobile",           default: true },
    { key: "rolle",         label: "Rolle",            default: false },
    { key: "leadbbz0",      label: "Lead BBZ",         default: false },
    { key: "sgf",           label: "SGF",              default: false },
    { key: "event",         label: "Events (aktuell)", default: true },
    { key: "eventhistory",  label: "Events (History)", default: false },
    { key: "geburtstag",    label: "Geburtstag",       default: false },
    { key: "kommentar",     label: "Kommentar",        default: false },
    { key: "archiviert",    label: "Archiviert",       default: false }
  ];

  // Wert eines Kontakt-Felds für den Export aufbereiten
  function getContactFieldValue(contact, key) {
    const v = contact[key];
    if (Array.isArray(v)) return v.join(", ");
    if (typeof v === "boolean") return v ? "Ja" : "Nein";
    if (key === "geburtstag" && v) {
      const d = new Date(v);
      return isNaN(d) ? v : d.toLocaleDateString("de-CH");
    }
    return v ?? "";
  }

  // ─── EXPORT MODAL ─────────────────────────────────────────────────────────────────

  /**
   * mode: "contacts" → alle / gefilterten Kontakte
   *       "event"    → Kontakte eines Events (zeigt Event-Auswahl)
   */
  function openExportModal(mode = "contacts") {
    removeModal("bbz-io-export-modal");

    const { state } = app();
    const events = state.enriched.events || [];
    const eventOptions = events.map(e =>
      `<option value="${escHtml(e.name)}">${escHtml(e.name)} (${e.contactCount})</option>`
    ).join("");

    const fieldCheckboxes = CONTACT_FIELDS.map(f => `
      <label class="bbz-multi-choice-item" style="min-width:160px">
        <input type="checkbox" name="field" value="${f.key}" ${f.default ? "checked" : ""} />
        <span>${escHtml(f.label)}</span>
      </label>
    `).join("");

    const eventSection = mode === "event" ? `
      <div class="bbz-field bbz-span-2" style="margin-bottom:4px">
        <label>Event auswählen</label>
        <select id="io-event-select" class="bbz-select">
          <option value="">— bitte wählen —</option>
          ${eventOptions}
        </select>
      </div>
    ` : "";

    const modeLabel = mode === "event" ? "Event-Kontakte exportieren" : "Kontakte exportieren";
    const scopeNote = mode === "contacts"
      ? `<p style="font-size:13px;color:var(--muted);margin:0 0 14px">Exportiert alle nicht-archivierten Kontakte. Archivierte Kontakte werden weggelassen, sofern nicht unten aktiviert.</p>`
      : `<p style="font-size:13px;color:var(--muted);margin:0 0 14px">Exportiert alle Kontakte des gewählten Events mit vollständigen Kontaktdaten.</p>`;

    const html = `
      <div id="bbz-io-export-modal" class="bbz-modal-backdrop show">
        <div class="bbz-modal" style="max-width:680px">
          <div class="bbz-modal-header">
            <span class="bbz-modal-title">📤 ${escHtml(modeLabel)}</span>
            <button class="bbz-button bbz-button-secondary" onclick="document.getElementById('bbz-io-export-modal').remove()" style="height:32px;padding:0 10px;font-size:13px">✕</button>
          </div>
          <div class="bbz-modal-body">
            ${scopeNote}
            <div class="bbz-form-grid">
              ${eventSection}
              <div class="bbz-field bbz-span-2">
                <label>Felder auswählen <span class="bbz-field-hint">— alle aktivierten Felder werden als Spalten exportiert</span></label>
                <div class="bbz-multi-choice" id="io-field-choices">
                  ${fieldCheckboxes}
                </div>
              </div>
              ${mode === "contacts" ? `
              <div class="bbz-field">
                <label>Archivierte Kontakte</label>
                <label class="bbz-checkbox" style="margin-top:4px">
                  <input type="checkbox" id="io-include-archived" />
                  <span>Archivierte einschliessen</span>
                </label>
              </div>` : ""}
              <div class="bbz-field" style="align-self:end">
                <div id="io-preview-count" style="font-size:13px;color:var(--muted);padding:8px 0">
                  — Felder wählen und Vorschau aktualisieren
                </div>
              </div>
            </div>
          </div>
          <div class="bbz-modal-footer">
            <button class="bbz-button bbz-button-secondary" onclick="document.getElementById('bbz-io-export-modal').remove()">Abbrechen</button>
            <button class="bbz-button bbz-button-primary" id="io-export-btn" onclick="window.bbzIO._doExport('${mode}')">
              ⬇ Excel herunterladen
            </button>
          </div>
        </div>
      </div>
    `;

    document.body.insertAdjacentHTML("beforeend", html);
    _updateExportPreview(mode);

    // Live-Vorschau bei Änderungen
    document.getElementById("bbz-io-export-modal").addEventListener("change", () => {
      _updateExportPreview(mode);
    });
  }

  function _updateExportPreview(mode) {
    const { state } = app();
    const el = document.getElementById("io-preview-count");
    if (!el) return;

    const selectedFields = _getSelectedFields();
    let count = 0;

    if (mode === "event") {
      const eventName = document.getElementById("io-event-select")?.value || "";
      if (!eventName) { el.textContent = "Bitte Event wählen"; return; }
      const ev = (state.enriched.events || []).find(e => e.name === eventName);
      count = ev?.contacts?.length || 0;
    } else {
      const inclArchived = document.getElementById("io-include-archived")?.checked ?? false;
      count = (state.enriched.contacts || []).filter(c => inclArchived || !c.archiviert).length;
    }

    el.textContent = `${count} Kontakt${count !== 1 ? "e" : ""} · ${selectedFields.length} Spalte${selectedFields.length !== 1 ? "n" : ""}`;
  }

  function _getSelectedFields() {
    const modal = document.getElementById("bbz-io-export-modal");
    if (!modal) return CONTACT_FIELDS.filter(f => f.default).map(f => f.key);
    return [...modal.querySelectorAll("input[name='field']:checked")].map(el => el.value);
  }

  function _doExport(mode) {
    if (!window.XLSX) { showBanner("SheetJS nicht geladen. Bitte Seite neu laden.", "error"); return; }

    const { state } = app();
    const selectedKeys = _getSelectedFields();
    if (!selectedKeys.length) { showBanner("Bitte mindestens ein Feld auswählen.", "warning"); return; }

    const fieldDefs = CONTACT_FIELDS.filter(f => selectedKeys.includes(f.key));
    const headers = fieldDefs.map(f => f.label);
    let rows = [];
    let filename = "";

    if (mode === "event") {
      const eventName = document.getElementById("io-event-select")?.value || "";
      if (!eventName) { showBanner("Bitte ein Event auswählen.", "warning"); return; }
      const ev = (state.enriched.events || []).find(e => e.name === eventName);
      if (!ev) { showBanner("Event nicht gefunden.", "error"); return; }

      // Vollständige Kontaktdaten aus enriched.contacts laden
      const contactMap = new Map((state.enriched.contacts || []).map(c => [c.id, c]));
      rows = ev.contacts.map(ec => {
        const contact = contactMap.get(ec.contactId) || ec;
        return fieldDefs.map(f => getContactFieldValue(contact, f.key));
      });
      filename = `bbzCRM_Event_${eventName.replace(/[^a-zA-Z0-9]/g, "_")}_${filenameDate()}.xlsx`;

    } else {
      const inclArchived = document.getElementById("io-include-archived")?.checked ?? false;
      const contacts = (state.enriched.contacts || []).filter(c => inclArchived || !c.archiviert);
      rows = contacts.map(c => fieldDefs.map(f => getContactFieldValue(c, f.key)));
      filename = `bbzCRM_Kontakte_${filenameDate()}.xlsx`;
    }

    if (!rows.length) { showBanner("Keine Daten für den Export gefunden.", "warning"); return; }

    const wsData = [headers, ...rows];
    const wb = window.XLSX.utils.book_new();
    const ws = window.XLSX.utils.aoa_to_sheet(wsData);

    // Spaltenbreiten automatisch anpassen
    ws["!cols"] = headers.map((h, i) => {
      const maxLen = Math.max(h.length, ...rows.map(r => String(r[i] ?? "").length));
      return { wch: Math.min(Math.max(maxLen + 2, 10), 50) };
    });

    // Header-Zeile fetten (SheetJS Pro-Feature ist nicht nötig — wir nutzen cell styles via aoa)
    window.XLSX.utils.book_append_sheet(wb, ws, "Kontakte");
    window.XLSX.writeFile(wb, filename);

    removeModal("bbz-io-export-modal");
    showBanner(`Export erfolgreich: ${rows.length} Kontakte als "${filename}" heruntergeladen.`, "success");
  }

  // ─── IMPORT MODAL ─────────────────────────────────────────────────────────────────

  /**
   * Master-Import: eine Excel-Datei mit bis zu 4 Sheets
   *   Sheet 1: "Firmen"    → CRMFirms
   *   Sheet 2: "Kontakte"  → CRMContacts (FirmaLookup via Firmenname-Matching)
   *   Sheet 3: "History"   → CRMHistory  (KontaktLookup via "Vorname Nachname")
   *   Sheet 4: "Tasks"     → CRMTasks    (KontaktLookup via "Vorname Nachname")
   *
   * Verarbeitungsreihenfolge: sequenziell (Firmen → Kontakte → History → Tasks)
   * damit LookupIds korrekt aufgelöst werden können.
   */
  function openImportModal() {
    removeModal("bbz-io-import-modal");

    const html = `
      <div id="bbz-io-import-modal" class="bbz-modal-backdrop show">
        <div class="bbz-modal" style="max-width:640px">
          <div class="bbz-modal-header">
            <span class="bbz-modal-title">📥 Master-Import</span>
            <button class="bbz-button bbz-button-secondary" onclick="document.getElementById('bbz-io-import-modal').remove()" style="height:32px;padding:0 10px;font-size:13px">✕</button>
          </div>
          <div class="bbz-modal-body">
            <p style="font-size:13px;color:var(--muted);margin:0 0 16px;line-height:1.5">
              Importiert Daten aus einer Excel-Datei mit den Sheets <strong>Firmen</strong>, <strong>Kontakte</strong>, <strong>Aktivitaeten</strong> und <strong>Tasks</strong>.
              Fehlende Sheets werden übersprungen. Bestehende Datensätze werden <strong>nicht</strong> überschrieben — der Import legt nur neue Items an.
            </p>

            <div style="background:#f8fafc;border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin-bottom:16px;font-size:13px;line-height:1.65">
              <strong>Pflichtfelder pro Sheet:</strong><br>
              📋 <strong>Firmen:</strong> Name<br>
              👤 <strong>Kontakte:</strong> Nachname, Firma (muss in Firmen-Sheet oder SP vorhanden sein)<br>
              📅 <strong>Aktivitaeten:</strong> Kontakt (Vorname Nachname), Datum (YYYY-MM-DD oder DD.MM.YYYY)<br>
              ✅ <strong>Tasks:</strong> Titel, Kontakt (Vorname Nachname)
            </div>

            <div class="bbz-field" style="margin-bottom:14px">
              <label>Excel-Datei auswählen (.xlsx)</label>
              <input type="file" id="io-import-file" accept=".xlsx,.xls" class="bbz-input" style="height:auto;padding:8px 12px;cursor:pointer" />
            </div>

            <div id="io-import-preview" style="display:none">
              <div style="font-size:12px;font-weight:700;color:var(--muted);letter-spacing:.05em;text-transform:uppercase;margin-bottom:8px">Vorschau</div>
              <div id="io-import-preview-body"></div>
            </div>

            <div id="io-import-log" style="display:none;margin-top:14px;max-height:200px;overflow-y:auto;font-size:12px;font-family:monospace;background:#f1f5f9;border-radius:10px;padding:10px 12px;line-height:1.6"></div>
          </div>
          <div class="bbz-modal-footer">
            <button class="bbz-button bbz-button-secondary" onclick="window.bbzIO.downloadTemplate()" style="margin-right:auto">⬇ Vorlage herunterladen</button>
            <button class="bbz-button bbz-button-secondary" onclick="document.getElementById('bbz-io-import-modal').remove()">Abbrechen</button>
            <button class="bbz-button bbz-button-primary" id="io-import-btn" disabled onclick="window.bbzIO._doImport()">
              ▶ Import starten
            </button>
          </div>
        </div>
      </div>
    `;

    document.body.insertAdjacentHTML("beforeend", html);

    document.getElementById("io-import-file").addEventListener("change", _handleFileSelect);
  }

  // Parsed die Datei und zeigt Vorschau
  async function _handleFileSelect(event) {
    if (!window.XLSX) { showBanner("SheetJS nicht geladen.", "error"); return; }
    const file = event.target.files?.[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const wb = window.XLSX.read(data, { type: "array", cellDates: true });

    window._bbzImportWorkbook = wb;

    const preview = document.getElementById("io-import-preview");
    const previewBody = document.getElementById("io-import-preview-body");
    const importBtn = document.getElementById("io-import-btn");

    const sheetSummary = IMPORT_SHEETS.map(s => {
      const ws = wb.Sheets[s.name];
      if (!ws) return `<div style="color:var(--muted);font-size:13px">⬜ Sheet <strong>${s.name}</strong> nicht gefunden — wird übersprungen</div>`;
      const rows = window.XLSX.utils.sheet_to_json(ws, { defval: "" });
      return `<div style="font-size:13px;margin-bottom:4px">✅ Sheet <strong>${s.name}</strong>: ${rows.length} Zeile${rows.length !== 1 ? "n" : ""} erkannt</div>`;
    }).join("");

    previewBody.innerHTML = sheetSummary;
    preview.style.display = "block";
    importBtn.disabled = false;
    document.getElementById("io-import-log").style.display = "none";
  }

  // Sheet-Definitionen (Spalten-Mapping Excel → SP-Felder)
  const IMPORT_SHEETS = [
    {
      name: "Firmen",
      required: ["Name"],
      toFields: (row) => {
        if (!String(row["Name"] || "").trim()) return null;
        const f = { Title: String(row["Name"]).trim() };
        if (row["Adresse"])        f.Adresse        = String(row["Adresse"]);
        if (row["PLZ"])            f.PLZ            = String(row["PLZ"]);
        if (row["Ort"])            f.Ort            = String(row["Ort"]);
        if (row["Land"])           f.Land           = String(row["Land"]);
        if (row["Hauptnummer"])    f.Hauptnummer    = String(row["Hauptnummer"]);
        if (row["Klassifizierung"])f.Klassifizierung= String(row["Klassifizierung"]);
        if (row["VIP"])            f.VIP            = ["ja","true","1","yes"].includes(String(row["VIP"]).toLowerCase());
        return f;
      }
    },
    {
      name: "Kontakte",
      required: ["Nachname", "Firma"],
      toFields: (row, lookups) => {
        const nachname = String(row["Nachname"] || "").trim();
        const firmaName = String(row["Firma"] || "").trim();
        if (!nachname || !firmaName) return null;

        // Lookup via Firmenname (case-insensitive)
        const firmaId = lookups.firmByName?.get(firmaName.toLowerCase());
        if (!firmaId) return null; // Firma nicht gefunden — überspringen + warnen

        const f = { Title: nachname, FirmaLookupId: firmaId };
        if (row["Vorname"])    f.Vorname    = String(row["Vorname"]).trim();
        if (row["Anrede"])     f.Anrede     = String(row["Anrede"]);
        if (row["Funktion"])   f.Funktion   = String(row["Funktion"]).trim();
        if (row["E-Mail 1"] || row["Email1"]) f.Email1 = String(row["E-Mail 1"] || row["Email1"]).trim();
        if (row["E-Mail 2"] || row["Email2"]) f.Email2 = String(row["E-Mail 2"] || row["Email2"]).trim();
        if (row["Direktwahl"]) f.Direktwahl = String(row["Direktwahl"]).trim();
        if (row["Mobile"])     f.Mobile     = String(row["Mobile"]).trim();
        if (row["Rolle"])      f.Rolle      = String(row["Rolle"]);
        if (row["Lead BBZ"] || row["Leadbbz0"]) f.Leadbbz0 = String(row["Lead BBZ"] || row["Leadbbz0"]);
        if (row["Kommentar"])  f.Kommentar  = String(row["Kommentar"]).trim();

        // Datum
        const geb = row["Geburtstag"];
        if (geb) {
          const parsed = _parseDate(geb);
          if (parsed) f.Geburtstag = parsed + "T00:00:00Z";
        }

        return { createFields: { Title: nachname, FirmaLookupId: firmaId }, patchFields: f, _missingFirma: null };
      },
      _missingFirmaWarning: (row) => `Firma "${row["Firma"]}" nicht gefunden → Kontakt "${row["Nachname"]}" übersprungen`
    },
    {
      name: "Aktivitaeten",
      required: ["Kontakt", "Datum"],
      toFields: (row, lookups) => {
        const kontaktName = String(row["Kontakt"] || "").trim();
        const datumRaw = row["Datum"];
        if (!kontaktName || !datumRaw) return null;

        const kontaktId = lookups.contactByName?.get(kontaktName.toLowerCase());
        if (!kontaktId) return null;

        const datum = _parseDate(datumRaw);
        if (!datum) return null;

        return {
          createFields: { Title: `Import-${datum}`, NachnameLookupId: kontaktId },
          patchFields: {
            Datum: datum + "T00:00:00Z",
            ...(row["Kontaktart"] ? { Kontaktart: String(row["Kontaktart"]) } : {}),
            ...(row["Lead BBZ"] || row["Leadbbz"] ? { Leadbbz: String(row["Lead BBZ"] || row["Leadbbz"]) } : {}),
            ...(row["Notizen"] ? { Notizen: String(row["Notizen"]).trim() } : {}),
            Projektbezug: ["ja","true","1"].includes(String(row["Projektbezug"] || "").toLowerCase())
          }
        };
      }
    },
    {
      name: "Tasks",
      required: ["Titel", "Kontakt"],
      toFields: (row, lookups) => {
        const title = String(row["Titel"] || "").trim();
        const kontaktName = String(row["Kontakt"] || "").trim();
        if (!title || !kontaktName) return null;

        const kontaktId = lookups.contactByName?.get(kontaktName.toLowerCase());
        if (!kontaktId) return null;

        const createFields = { Title: title, NameLookupId: kontaktId };
        const patchFields = {};
        const dl = row["Deadline"];
        if (dl) { const d = _parseDate(dl); if (d) patchFields.Deadline = d + "T00:00:00Z"; }
        if (row["Status"])   patchFields.Status  = String(row["Status"]);
        if (row["Lead BBZ"] || row["Leadbbz"]) patchFields.Leadbbz = String(row["Lead BBZ"] || row["Leadbbz"]);

        return { createFields, patchFields };
      }
    }
  ];

  // Datum-Parser: erkennt YYYY-MM-DD, DD.MM.YYYY, JS Date-Objekte
  function _parseDate(value) {
    if (!value) return null;
    // Hilfsfunktion: lokale Datum-Komponenten als YYYY-MM-DD
    // Wichtig: NICHT toISOString() verwenden — das gibt UTC zurück und verschiebt
    // in CH (UTC+1) Mitternacht-Datumswerte (z.B. aus Excel) um einen Tag zurück
    function _localDateStr(d) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      return `${y}-${m}-${day}`;
    }
    if (value instanceof Date) {
      if (isNaN(value.getTime())) return null;
      return _localDateStr(value);
    }
    const s = String(value).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
    if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) {
      const [d, m, y] = s.split(".");
      return `${y}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : _localDateStr(d);
  }

  // In-memory Log-Buffer — wird bei jedem Import neu befüllt
  let _logBuffer = [];

  function _log(msg, type = "info") {
    // UI
    const el = document.getElementById("io-import-log");
    if (el) {
      el.style.display = "block";
      const color = { ok: "#15803d", warn: "#b45309", error: "#b91c1c", info: "#475569" }[type] || "#475569";
      el.innerHTML += `<div style="color:${color}">${escHtml(msg)}</div>`;
      el.scrollTop = el.scrollHeight;
    }
    // Buffer
    const prefix = { ok: "[OK]  ", warn: "[WARN]", error: "[ERR] ", info: "[INFO]" }[type] || "[INFO]";
    _logBuffer.push(`${prefix}  ${msg}`);
  }

  function _downloadLog(filename, totalCreated, totalSkipped, hadCriticalError) {
    const ts = new Date().toLocaleString("de-CH", {
      day: "2-digit", month: "2-digit", year: "numeric",
      hour: "2-digit", minute: "2-digit", second: "2-digit"
    });

    const header = [
      "=======================================================",
      "  bbz CRM -- Import-Protokoll",
      `  Datum/Zeit  : ${ts}`,
      `  Datei       : ${window._bbzImportFilename || "unbekannt"}`,
      `  Ergebnis    : ${hadCriticalError ? "FEHLER" : "OK"}`,
      `  Angelegt    : ${totalCreated}`,
      `  Uebersprungen: ${totalSkipped}`,
      "=======================================================",
      ""
    ].join("\n");

    const body = _logBuffer.join("\n");
    const content = header + body + "\n";

    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function _doImport() {
    if (!window.XLSX) { showBanner("SheetJS nicht geladen.", "error"); return; }
    if (!window._bbzImportWorkbook) { showBanner("Keine Datei geladen.", "warning"); return; }

    const { api, state } = app();
    const wb = window._bbzImportWorkbook;

    const importBtn = document.getElementById("io-import-btn");
    importBtn.disabled = true;
    importBtn.textContent = "Import läuft ...";

    const log = document.getElementById("io-import-log");
    if (log) { log.innerHTML = ""; log.style.display = "block"; }

    _logBuffer = []; // Buffer leeren fuer neuen Import
    window._bbzImportFilename = document.getElementById("io-import-file")?.files?.[0]?.name || "unbekannt";

    let totalCreated = 0;
    let totalSkipped = 0;
    let hadCriticalError = false;

    try {
      // Schritt 1: Firma-Lookup aufbauen (bestehende SP-Firmen + Import-Firmen)
      _log("Lade aktuelle Firmenliste aus SharePoint ...", "info");
      await api.loadAll(); // frisch laden für aktuellen Stand

      const firmByName = new Map(
        (state.enriched.firms || []).map(f => [f.title.toLowerCase(), f.id])
      );

      // ── FIRMEN ────────────────────────────────────────────────────────────
      const firmenWs = wb.Sheets["Firmen"];
      if (firmenWs) {
        const rows = window.XLSX.utils.sheet_to_json(firmenWs, { defval: "" });
        _log(`→ Firmen: ${rows.length} Zeilen gefunden`, "info");
        let created = 0, skipped = 0;

        for (const row of rows) {
          const sheetDef = IMPORT_SHEETS[0];
          const fields = sheetDef.toFields(row);
          if (!fields) { skipped++; _log(`  ⚠ Firma übersprungen (Pflichtfeld fehlt): ${JSON.stringify(row)}`, "warn"); continue; }

          const nameLower = fields.Title.toLowerCase();
          if (firmByName.has(nameLower)) {
            skipped++;
            _log(`  ⬜ Firma bereits vorhanden: "${fields.Title}"`, "info");
            continue;
          }

          try {
            const created_ = await api.postItem("CRMFirms", fields);
            const newId = created_?.id || created_?.fields?.id;
            if (newId) firmByName.set(nameLower, Number(newId));
            created++;
            _log(`  ✅ Firma angelegt: "${fields.Title}"`, "ok");
          } catch (err) {
            skipped++;
            _log(`  ❌ Firma Fehler ("${fields.Title}"): ${err.message}`, "error");
          }
        }
        _log(`  Firmen: ${created} angelegt, ${skipped} übersprungen`, "info");
        totalCreated += created; totalSkipped += skipped;
      }

      // ── KONTAKTE ──────────────────────────────────────────────────────────
      const kontakteWs = wb.Sheets["Kontakte"];

      // Duplikat-Schlüssel: "vorname nachname|firmatitel" (case-insensitive)
      // Speichert { id, contact } für Smart-Merge
      const contactDupMap = new Map();
      (state.enriched.contacts || []).forEach(c => {
        const key = `${c.vorname} ${c.nachname}|${c.firmTitle}`.toLowerCase().trim();
        contactDupMap.set(key, { id: c.id, contact: c });
      });

      // Lookup-Map für History/Tasks (Vorname Nachname → id)
      const contactByName = new Map(
        (state.enriched.contacts || []).map(c => [
          `${c.vorname} ${c.nachname}`.toLowerCase().trim(), c.id
        ])
      );
      (state.enriched.contacts || []).forEach(c => {
        contactByName.set(`${c.nachname} ${c.vorname}`.toLowerCase().trim(), c.id);
        contactByName.set(c.nachname.toLowerCase().trim(), c.id);
      });

      // Felder die beim Smart-Merge befüllt werden (nur wenn im SP leer)
      const MERGE_FIELDS = [
        { row: "Vorname",    sp: "Vorname",    contact: "vorname" },
        { row: "Anrede",     sp: "Anrede",     contact: "anrede" },
        { row: "Funktion",   sp: "Funktion",   contact: "funktion" },
        { row: "E-Mail 1",   sp: "Email1",     contact: "email1" },
        { row: "E-Mail 2",   sp: "Email2",     contact: "email2" },
        { row: "Direktwahl", sp: "Direktwahl", contact: "direktwahl" },
        { row: "Mobile",     sp: "Mobile",     contact: "mobile" },
        { row: "Rolle",      sp: "Rolle",      contact: "rolle" },
        { row: "Lead BBZ",   sp: "Leadbbz0",   contact: "leadbbz0" },
        { row: "Kommentar",  sp: "Kommentar",  contact: "kommentar" }
        // Multi-Choice (SGF, Event, Eventhistory) bewusst weggelassen:
        // Merge-Semantik unklar (ergänzen vs. ersetzen) → Datenrisiko zu hoch.
      ];

      // Geburtstag separat: braucht Datumskonvertierung
      function _mergeGeburtstag(row, existingContact) {
        const importVal = String(row["Geburtstag"] || "").trim();
        const existingVal = String(existingContact.geburtstag || "").trim();
        if (!importVal || existingVal) return null; // nichts zu tun
        const parsed = _parseDate(importVal);
        return parsed ? parsed + "T00:00:00Z" : null;
      }

      let totalMerged = 0;

      if (kontakteWs) {
        const rows = window.XLSX.utils.sheet_to_json(kontakteWs, { defval: "" });
        _log(`→ Kontakte: ${rows.length} Zeilen gefunden`, "info");
        let created = 0, skipped = 0, merged = 0;

        for (const row of rows) {
          const sheetDef = IMPORT_SHEETS[1];
          const nachname  = String(row["Nachname"] || "").trim();
          const vorname   = String(row["Vorname"]  || "").trim();
          const firmaName = String(row["Firma"]    || "").trim();

          if (!nachname || !firmaName) {
            skipped++;
            _log(`  ⚠ Kontakt übersprungen (Pflichtfeld fehlt): ${nachname || "?"}`, "warn");
            continue;
          }

          // ── Duplikat-Prüfung: Vorname + Nachname + Firma ──
          const dupKey = `${vorname} ${nachname}|${firmaName}`.toLowerCase().trim();
          const existing = contactDupMap.get(dupKey);

          if (existing) {
            // Smart-Merge: nur leere Felder ergänzen
            const patch = {};
            for (const fd of MERGE_FIELDS) {
              const importVal = String(row[fd.row] || "").trim();
              const existingVal = String(existing.contact[fd.contact] || "").trim();
              if (importVal && !existingVal) {
                patch[fd.sp] = importVal;
              }
            }

            // Geburtstag prüfen und ggf. ergänzen
            const gebVal = _mergeGeburtstag(row, existing.contact);
            if (gebVal) patch["Geburtstag"] = gebVal;

            if (Object.keys(patch).length > 0) {
              try {
                await api.patchItem("CRMContacts", existing.id, patch);
                const updatedFields = Object.keys(patch).join(", ");
                _log(`  🔄 Kontakt ergänzt: "${vorname} ${nachname}" → ${updatedFields}`, "ok");
                merged++;
              } catch (err) {
                skipped++;
                _log(`  ❌ Merge Fehler ("${vorname} ${nachname}"): ${err.message}`, "error");
              }
            } else {
              skipped++;
              _log(`  ⬜ Duplikat, keine neuen Felder: "${vorname} ${nachname}" (${firmaName})`, "info");
            }

            // Kontakt für History/Tasks-Lookup registrieren (auch wenn kein Merge nötig)
            contactByName.set(`${vorname} ${nachname}`.toLowerCase().trim(), existing.id);
            continue;
          }

          // ── Neu anlegen ──
          const result = sheetDef.toFields(row, { firmByName });
          if (!result) {
            skipped++;
            if (!firmByName.has(firmaName.toLowerCase())) {
              _log(`  ⚠ ${sheetDef._missingFirmaWarning(row)}`, "warn");
            } else {
              _log(`  ⚠ Kontakt übersprungen (Firma-Lookup fehlgeschlagen): ${nachname}`, "warn");
            }
            continue;
          }

          try {
            const created_ = await api.postItem("CRMContacts", result.createFields);
            const newId = Number(created_?.id || created_?.fields?.id);
            if (!newId) throw new Error("Keine Item-ID im POST-Response.");

            const pf = { ...result.patchFields };
            delete pf.Title; delete pf.FirmaLookupId;
            if (Object.keys(pf).length > 0) {
              await api.patchItem("CRMContacts", newId, pf);
            }

            const fullName = `${vorname} ${nachname}`.toLowerCase().trim();
            contactByName.set(fullName, newId);
            contactByName.set(nachname.toLowerCase(), newId);
            contactDupMap.set(dupKey, { id: newId, contact: { vorname, nachname, firmTitle: firmaName, ...result.patchFields } });
            created++;
            _log(`  ✅ Kontakt neu angelegt: "${vorname} ${nachname}"`, "ok");
          } catch (err) {
            skipped++;
            _log(`  ❌ Kontakt Fehler ("${nachname}"): ${err.message}`, "error");
          }
        }
        totalMerged += merged;
        _log(`  Kontakte: ${created} angelegt, ${merged} ergänzt (Merge), ${skipped} übersprungen`, "info");
        totalCreated += created; totalSkipped += skipped;
      }

      // ── HISTORY ───────────────────────────────────────────────────────────
      const historyWs = wb.Sheets["Aktivitaeten"];

      // Duplikat-Schlüssel: kontaktId|datum|kontaktart
      const historyDupSet = new Set(
        (state.enriched.history || []).map(h =>
          `${h.contactId}|${h.datum ? h.datum.slice(0,10) : ""}|${String(h.typ||"").toLowerCase()}`
        )
      );

      if (historyWs) {
        const rows = window.XLSX.utils.sheet_to_json(historyWs, { defval: "" });
        _log(`→ Aktivitaeten: ${rows.length} Zeilen gefunden`, "info");
        let created = 0, skipped = 0;

        for (const row of rows) {
          const result = IMPORT_SHEETS[2].toFields(row, { contactByName });
          if (!result) {
            skipped++;
            _log(`  ⚠ Aktivitaet übersprungen: Kontakt "${row["Kontakt"]||"?"}" nicht gefunden oder Datum fehlt`, "warn");
            continue;
          }

          // Duplikat-Prüfung
          const datum = _parseDate(row["Datum"]) || "";
          const kontaktId = contactByName.get(String(row["Kontakt"]||"").trim().toLowerCase());
          const kontaktart = String(row["Kontaktart"]||"").toLowerCase();
          const histDupKey = `${kontaktId}|${datum}|${kontaktart}`;
          if (historyDupSet.has(histDupKey)) {
            skipped++;
            _log(`  ⬜ Duplikat Aktivitaet übersprungen: "${row["Kontakt"]}" am ${datum} (${row["Kontaktart"]||"—"})`, "info");
            continue;
          }

          try {
            const created_ = await api.postItem("CRMHistory", result.createFields);
            const newId = Number(created_?.id || created_?.fields?.id);
            if (!newId) throw new Error("Keine Item-ID.");
            if (Object.keys(result.patchFields).length > 0) {
              await api.patchItem("CRMHistory", newId, result.patchFields);
            }
            historyDupSet.add(histDupKey); // verhindert Duplikate innerhalb desselben Imports
            created++;
            _log(`  ✅ Aktivitaet angelegt: "${row["Kontakt"]}" am ${datum}`, "ok");
          } catch (err) {
            skipped++;
            _log(`  ❌ Aktivitaet Fehler: ${err.message}`, "error");
          }
        }
        _log(`  Aktivitaeten: ${created} angelegt, ${skipped} übersprungen`, "info");
        totalCreated += created; totalSkipped += skipped;
      }

      // ── TASKS ─────────────────────────────────────────────────────────────
      const tasksWs = wb.Sheets["Tasks"];

      // Duplikat-Schlüssel: kontaktId|titel (case-insensitive)
      const taskDupSet = new Set(
        (state.enriched.tasks || []).map(t =>
          `${t.contactId}|${String(t.title||"").toLowerCase().trim()}`
        )
      );

      if (tasksWs) {
        const rows = window.XLSX.utils.sheet_to_json(tasksWs, { defval: "" });
        _log(`→ Tasks: ${rows.length} Zeilen gefunden`, "info");
        let created = 0, skipped = 0;

        for (const row of rows) {
          const result = IMPORT_SHEETS[3].toFields(row, { contactByName });
          if (!result) {
            skipped++;
            _log(`  ⚠ Task übersprungen: Kontakt "${row["Kontakt"]||"?"}" nicht gefunden oder Titel fehlt`, "warn");
            continue;
          }

          // Duplikat-Prüfung
          const kontaktId = contactByName.get(String(row["Kontakt"]||"").trim().toLowerCase());
          const titel = String(row["Titel"]||"").toLowerCase().trim();
          const taskDupKey = `${kontaktId}|${titel}`;
          if (taskDupSet.has(taskDupKey)) {
            skipped++;
            _log(`  ⬜ Duplikat Task übersprungen: "${row["Titel"]}" für "${row["Kontakt"]}"`, "info");
            continue;
          }

          try {
            const created_ = await api.postItem("CRMTasks", result.createFields);
            const newId = Number(created_?.id || created_?.fields?.id);
            if (!newId) throw new Error("Keine Item-ID.");
            if (Object.keys(result.patchFields).length > 0) {
              await api.patchItem("CRMTasks", newId, result.patchFields);
            }
            taskDupSet.add(taskDupKey); // verhindert Duplikate innerhalb desselben Imports
            created++;
            _log(`  ✅ Task angelegt: "${row["Titel"]}" für "${row["Kontakt"]}"`, "ok");
          } catch (err) {
            skipped++;
            _log(`  ❌ Task Fehler: ${err.message}`, "error");
          }
        }
        _log(`  Tasks: ${created} angelegt, ${skipped} übersprungen`, "info");
        totalCreated += created; totalSkipped += skipped;
      }

      _log(`\n Import abgeschlossen: ${totalCreated} neu angelegt, ${totalMerged} ergaenzt (Merge), ${totalSkipped} uebersprungen.`, "ok");
      await api.loadAll();
      const logFilename = `bbzCRM_Import_Log_${new Date().toISOString().slice(0,10).replace(/-/g,"")}.txt`;
      _downloadLog(logFilename, totalCreated, totalSkipped, false);
      showBanner(`Import erfolgreich: ${totalCreated} neue Eintraege angelegt, ${totalSkipped} uebersprungen. Logfile wurde heruntergeladen.`, "success");
      importBtn.textContent = "Abgeschlossen";

    } catch (err) {
      hadCriticalError = true;
      _log(`\n Kritischer Fehler: ${err.message}`, "error");
      const errLogFilename = `bbzCRM_Import_Log_${new Date().toISOString().slice(0,10).replace(/-/g,"")}_FEHLER.txt`;
      _downloadLog(errLogFilename, totalCreated, totalSkipped, true);
      showBanner(`Importfehler: ${err.message} — Logfile wurde heruntergeladen.`, "error");
      importBtn.disabled = false;
      importBtn.textContent = "Import starten";
    }
  }

  // ─── TEMPLATE DOWNLOAD ────────────────────────────────────────────────────────────

  function downloadTemplate() {
    if (!window.XLSX) { showBanner("SheetJS nicht geladen.", "error"); return; }

    const wb = window.XLSX.utils.book_new();

    const sheets = [
      {
        name: "Firmen",
        headers: ["Name", "Adresse", "PLZ", "Ort", "Land", "Hauptnummer", "Klassifizierung", "VIP"],
        example: ["Muster AG", "Musterstrasse 1", "9000", "St. Gallen", "CH", "+41 71 000 00 00", "A", "Nein"]
      },
      {
        name: "Kontakte",
        headers: ["Nachname", "Vorname", "Anrede", "Firma", "Funktion", "E-Mail 1", "E-Mail 2", "Direktwahl", "Mobile", "Rolle", "Lead BBZ", "Kommentar", "Geburtstag"],
        example: ["Mustermann", "Max", "Herr", "Muster AG", "CEO", "max@muster.ch", "", "+41 71 000 00 01", "+41 79 000 00 01", "", "", "", "1980-06-15"]
      },
      {
        name: "Aktivitaeten",
        headers: ["Kontakt", "Datum", "Kontaktart", "Lead BBZ", "Notizen", "Projektbezug"],
        example: ["Max Mustermann", "2024-01-15", "Telefon", "", "Erstgespräch geführt", "Nein"]
      },
      {
        name: "Tasks",
        headers: ["Titel", "Kontakt", "Deadline", "Status", "Lead BBZ"],
        example: ["Offerte senden", "Max Mustermann", "2024-02-01", "Offen", ""]
      }
    ];

    sheets.forEach(s => {
      const ws = window.XLSX.utils.aoa_to_sheet([s.headers, s.example]);
      ws["!cols"] = s.headers.map((h, i) => {
        const maxLen = Math.max(h.length, String(s.example[i] || "").length);
        return { wch: Math.min(Math.max(maxLen + 2, 12), 40) };
      });
      window.XLSX.utils.book_append_sheet(wb, ws, s.name);
    });

    window.XLSX.writeFile(wb, `bbzCRM_Import_Vorlage.xlsx`);
    showBanner("Vorlage heruntergeladen: bbzCRM_Import_Vorlage.xlsx", "success");
  }

  // ─── ÖFFENTLICHE API ──────────────────────────────────────────────────────────────

  window.bbzIO = {
    openExportModal,
    openImportModal,
    downloadTemplate,
    // intern (aber von onclick-Attributen im Modal aufgerufen)
    _doExport,
    _doImport
  };

})();
