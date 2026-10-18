// popup.js
// Controller for the FroggIteAms Chrome Extension Popup

import {
  exportMeetingToTxt,
  exportMeetingToCsv,
  exportRequestsReportToTxt,
  sanitizeFilename
} from "./lib/exporter.js";
import { DEFAULT_AI_CONFIG } from "./lib/ai-providers.js";

let state = {
  meetings: [],
  activeMeeting: null,
  aiConfig: DEFAULT_AI_CONFIG
};

const DEMO_PHRASES = [
  { speaker: "Carlos Mendoza", text: "Articulando esta reunión dado que el ciclo actual en la nube termina mañana." },
  { speaker: "Alejandro Morales", text: "Por favor Martín, hoy antes de las 15:00 necesitamos que quede creada la solicitud de transferencia en el portal." },
  { speaker: "Martín Castro", text: "Perfecto, enviaremos los dos enlaces iniciales de autorización para la aceptación de Roberto Guzmán." },
  { speaker: "Roberto Guzmán", text: "De acuerdo, mañana a las 12:00 hrs hacemos la aprobación junto con el equipo técnico." }
];
let demoIndex = 0;

document.addEventListener("DOMContentLoaded", () => {
  initTabs();
  initActions();
  refreshState();
});

function initTabs() {
  const buttons = document.querySelectorAll(".tab-btn");
  const panels = document.querySelectorAll(".tab-panel");

  buttons.forEach((btn) => {
    btn.addEventListener("click", () => {
      buttons.forEach(b => b.classList.remove("active"));
      panels.forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(`tab-${btn.dataset.tab}`)?.classList.add("active");
    });
  });
}

function initActions() {
  document.getElementById("btn-open-dashboard").addEventListener("click", () => {
    chrome.tabs.create({ url: chrome.runtime.getURL("dashboard.html") });
  });

  document.getElementById("search-meetings").addEventListener("input", (e) => {
    renderMeetingsList(e.target.value);
  });

  document.getElementById("import-txt-input").addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const content = await file.text();
    chrome.runtime.sendMessage(
      { type: "IMPORT_TXT_MEETING", content, filename: file.name },
      () => refreshState()
    );
  });

  // Live simulation
  document.getElementById("btn-simulate-line").addEventListener("click", () => {
    const item = DEMO_PHRASES[demoIndex % DEMO_PHRASES.length];
    demoIndex++;
    chrome.runtime.sendMessage(
      {
        type: "CAPTION_ENTRY",
        speaker: item.speaker,
        text: item.text,
        meetingTitle: "Sesión Sincronización MS Teams (Demo en Vivo)"
      },
      () => refreshState()
    );
  });

  document.getElementById("btn-finish-live").addEventListener("click", () => {
    chrome.runtime.sendMessage({ type: "FINISH_MEETING" }, () => {
      refreshState();
      document.querySelector('[data-tab="history"]')?.click();
    });
  });

  // AI Provider Switcher
  const providerSelect = document.getElementById("ai-provider-select");
  providerSelect.addEventListener("change", () => {
    updateProviderBoxes(providerSelect.value);
  });

  document.querySelectorAll('input[name="gemini-mode"]').forEach((radio) => {
    radio.addEventListener("change", (e) => {
      toggleGeminiAuthMode(e.target.value);
    });
  });

  document.getElementById("btn-gemini-oauth").addEventListener("click", () => {
    const clientId = document.getElementById("gemini-client-id").value.trim();
    const statusEl = document.getElementById("ai-save-status");
    statusEl.textContent = "Iniciando flujo Google OAuth...";
    chrome.runtime.sendMessage({ type: "LAUNCH_GEMINI_OAUTH", clientId }, (res) => {
      if (res?.ok && res.accessToken) {
        document.getElementById("gemini-oauth-token").value = res.accessToken;
        statusEl.textContent = "✅ Token OAuth de Google Gemini obtenido.";
      } else {
        statusEl.textContent = `ℹ️ ${res?.error || "Ingresa tu Bearer Token manualmente abajo."}`;
      }
    });
  });

  document.getElementById("btn-save-ai").addEventListener("click", saveAiSettings);
}

function refreshState() {
  chrome.runtime.sendMessage({ type: "GET_STATE" }, (res) => {
    if (!res?.ok) return;
    state.meetings = res.meetings || [];
    state.activeMeeting = res.activeMeeting || null;
    state.aiConfig = res.aiConfig || DEFAULT_AI_CONFIG;

    renderMeetingsList(document.getElementById("search-meetings")?.value || "");
    renderLivePanel();
    populateAiForm(state.aiConfig);
  });
}

function renderMeetingsList(query = "") {
  const container = document.getElementById("meetings-list");
  const q = query.trim().toLowerCase();

  const filtered = state.meetings.filter((m) => {
    if (!q) return true;
    const inTitle = (m.title || "").toLowerCase().includes(q);
    const inParticipants = (m.participants || []).some(p => p.toLowerCase().includes(q));
    return inTitle || inParticipants;
  });

  if (filtered.length === 0) {
    container.innerHTML = `<p style="color:#94a3b8;font-size:12px;text-align:center;padding:20px;">No se encontraron reuniones.</p>`;
    return;
  }

  container.innerHTML = filtered
    .map(
      (m) => `
    <div class="meeting-card" data-id="${escapeHtml(m.id)}">
      <div class="meeting-card-title">🐸 ${escapeHtml(m.title)}</div>
      <div class="meeting-card-meta">
        <span>📅 ${escapeHtml(m.startedDisplay || "")}</span>
        <span>⏱ ${escapeHtml(m.durationDisplay || `${m.durationMinutes || 0} min`)} · 👥 ${(m.participants || []).length}</span>
      </div>
      <div class="meeting-card-actions">
        <button class="btn-xs" data-action="txt" data-id="${escapeHtml(m.id)}" title="Exportar formato Tactiq TXT">📄 TXT</button>
        <button class="btn-xs" data-action="csv" data-id="${escapeHtml(m.id)}" title="Exportar tabla CSV">📊 CSV</button>
        <button class="btn-xs accent" data-action="report" data-id="${escapeHtml(m.id)}" title="Ver / Generar Reporte de Solicitudes">📋 Solicitudes (${m.requestsReport?.requests?.length || 0})</button>
      </div>
    </div>
  `
    )
    .join("");

  container.querySelectorAll("button[data-action]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const meeting = state.meetings.find(m => m.id === btn.dataset.id);
      if (!meeting) return;
      const action = btn.dataset.action;

      if (action === "txt") {
        const content = exportMeetingToTxt(meeting, false);
        downloadBlob(content, `${sanitizeFilename(meeting.title)}.txt`, "text/plain;charset=utf-8");
      } else if (action === "csv") {
        const content = exportMeetingToCsv(meeting);
        downloadBlob(content, `${sanitizeFilename(meeting.title)}.csv`, "text/csv;charset=utf-8");
      } else if (action === "report") {
        chrome.tabs.create({
          url: chrome.runtime.getURL(`dashboard.html?meetingId=${encodeURIComponent(meeting.id)}&view=requests`)
        });
      }
    });
  });
}

function renderLivePanel() {
  const pill = document.getElementById("live-status-pill");
  const countEl = document.getElementById("live-lines-count");
  const preview = document.getElementById("live-preview-box");

  const active = state.activeMeeting;
  if (!active || !active.transcript || active.transcript.length === 0) {
    pill.textContent = "Esperando audio/subtítulos MS Teams";
    countEl.textContent = "0 líneas";
    preview.innerHTML = `<span style="color:#64748b;">Sin frases en la sesión actual. Pulsa "Simular Frase en Vivo" o únete a una llamada en MS Teams.</span>`;
    return;
  }

  pill.textContent = `🔴 En vivo: ${active.title}`;
  countEl.textContent = `${active.transcript.length} líneas`;
  preview.innerHTML = active.transcript
    .slice(-10)
    .map(
      (t) => `
    <div>
      <strong style="color:#34d399;font-family:monospace;">${escapeHtml(t.time)}</strong>
      <strong style="color:#c7d2fe;"> ${escapeHtml(t.speaker)}:</strong>
      <span> ${escapeHtml(t.text)}</span>
    </div>
  `
    )
    .join("");
}

function updateProviderBoxes(provider) {
  ["gemini", "deepseek", "openai", "custom"].forEach((p) => {
    const box = document.getElementById(`box-${p}`);
    if (box) box.classList.toggle("hidden", p !== provider);
  });
}

function toggleGeminiAuthMode(mode) {
  document.getElementById("gemini-apikey-fields").classList.toggle("hidden", mode !== "apikey");
  document.getElementById("gemini-oauth-fields").classList.toggle("hidden", mode !== "oauth");
}

function populateAiForm(cfg) {
  const provider = cfg.activeProvider || "gemini";
  document.getElementById("ai-provider-select").value = provider;
  updateProviderBoxes(provider);

  const gem = cfg.gemini || {};
  const authMode = gem.authMode || "apikey";
  const radio = document.querySelector(`input[name="gemini-mode"][value="${authMode}"]`);
  if (radio) radio.checked = true;
  toggleGeminiAuthMode(authMode);

  document.getElementById("gemini-api-key").value = gem.apiKey || "";
  document.getElementById("gemini-client-id").value = gem.oauthClientId || "";
  document.getElementById("gemini-oauth-token").value = gem.oauthToken || "";
  document.getElementById("gemini-model").value = gem.model || "gemini-2.0-flash";

  document.getElementById("deepseek-api-key").value = cfg.deepseek?.apiKey || "";
  document.getElementById("deepseek-model").value = cfg.deepseek?.model || "deepseek-chat";

  document.getElementById("openai-api-key").value = cfg.openai?.apiKey || "";
  document.getElementById("openai-model").value = cfg.openai?.model || "gpt-4o-mini";

  document.getElementById("custom-endpoint").value = cfg.custom?.endpoint || "https://api.groq.com/openai/v1/chat/completions";
  document.getElementById("custom-api-key").value = cfg.custom?.apiKey || "";
  document.getElementById("custom-model").value = cfg.custom?.model || "llama-3.3-70b-versatile";
}

function saveAiSettings() {
  const authMode = document.querySelector('input[name="gemini-mode"]:checked')?.value || "apikey";
  const newConfig = {
    activeProvider: document.getElementById("ai-provider-select").value,
    gemini: {
      authMode,
      apiKey: document.getElementById("gemini-api-key").value.trim(),
      oauthClientId: document.getElementById("gemini-client-id").value.trim(),
      oauthToken: document.getElementById("gemini-oauth-token").value.trim(),
      model: document.getElementById("gemini-model").value.trim() || "gemini-2.0-flash"
    },
    deepseek: {
      apiKey: document.getElementById("deepseek-api-key").value.trim(),
      model: document.getElementById("deepseek-model").value.trim() || "deepseek-chat",
      endpoint: "https://api.deepseek.com/chat/completions"
    },
    openai: {
      apiKey: document.getElementById("openai-api-key").value.trim(),
      model: document.getElementById("openai-model").value.trim() || "gpt-4o-mini",
      endpoint: "https://api.openai.com/v1/chat/completions"
    },
    custom: {
      endpoint: document.getElementById("custom-endpoint").value.trim(),
      apiKey: document.getElementById("custom-api-key").value.trim(),
      model: document.getElementById("custom-model").value.trim()
    }
  };

  chrome.runtime.sendMessage({ type: "SAVE_AI_CONFIG", aiConfig: newConfig }, () => {
    state.aiConfig = newConfig;
    const statusEl = document.getElementById("ai-save-status");
    statusEl.textContent = "✅ Proveedor y tokens guardados correctamente.";
    setTimeout(() => {
      statusEl.textContent = "";
    }, 3000);
  });
}

function downloadBlob(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
