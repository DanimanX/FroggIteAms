// dashboard.js
// Full Tactiq-style Meeting Hub Controller for FroggIteAms

import {
  exportMeetingToTxt,
  exportMeetingToCsv,
  exportRequestsReportToTxt,
  exportRequestsReportToCsv,
  sanitizeFilename
} from "./lib/exporter.js";
import {
  DEFAULT_AI_CONFIG,
  generateRequestsReportWithAI
} from "./lib/ai-providers.js";
import { SAMPLE_MEETING } from "./lib/sample-data.js";

let state = {
  meetings: [SAMPLE_MEETING],
  selectedMeetingId: SAMPLE_MEETING.id,
  aiConfig: DEFAULT_AI_CONFIG
};

document.addEventListener("DOMContentLoaded", () => {
  initViewTabs();
  initToolbarEvents();
  initAiModal();
  loadStateFromExtension();
});

function loadStateFromExtension() {
  const params = new URLSearchParams(window.location.search);
  const requestedId = params.get("meetingId");
  const requestedView = params.get("view");

  if (typeof chrome !== "undefined" && chrome.runtime?.sendMessage) {
    chrome.runtime.sendMessage({ type: "GET_STATE" }, (res) => {
      if (res?.ok) {
        state.meetings = res.meetings?.length ? res.meetings : [SAMPLE_MEETING];
        state.aiConfig = res.aiConfig || DEFAULT_AI_CONFIG;
        if (requestedId && state.meetings.some(m => m.id === requestedId)) {
          state.selectedMeetingId = requestedId;
        } else {
          state.selectedMeetingId = state.meetings[0].id;
        }
      }
      renderAll();
      if (requestedView) {
        document.querySelector(`.view-tab[data-view="${requestedView}"]`)?.click();
      }
    });
  } else {
    renderAll();
  }
}

function getSelectedMeeting() {
  return (
    state.meetings.find(m => m.id === state.selectedMeetingId) ||
    state.meetings[0] ||
    SAMPLE_MEETING
  );
}

function initViewTabs() {
  const tabs = document.querySelectorAll(".view-tab");
  const sections = document.querySelectorAll(".view-section");

  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach(t => t.classList.remove("active"));
      sections.forEach(s => s.classList.remove("active"));
      tab.classList.add("active");
      document.getElementById(`view-${tab.dataset.view}`)?.classList.add("active");
    });
  });
}

function initToolbarEvents() {
  document.getElementById("sidebar-search").addEventListener("input", (e) => {
    renderSidebarList(e.target.value);
  });

  document.getElementById("search-in-transcript").addEventListener("input", () => {
    renderTranscriptAndHighlights();
  });

  document.getElementById("speaker-filter").addEventListener("change", () => {
    renderTranscriptAndHighlights();
  });

  document.getElementById("file-import-txt").addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const content = await file.text();
    if (typeof chrome !== "undefined" && chrome.runtime?.sendMessage) {
      chrome.runtime.sendMessage(
        { type: "IMPORT_TXT_MEETING", content, filename: file.name },
        (res) => {
          if (res?.ok) {
            state.meetings = res.meetings;
            state.selectedMeetingId = res.meeting.id;
            renderAll();
          }
        }
      );
    }
  });

  // Main Exports
  document.getElementById("btn-export-txt").addEventListener("click", () => {
    const meeting = getSelectedMeeting();
    const includeReport = document.getElementById("chk-include-report-in-txt")?.checked || false;
    const txt = exportMeetingToTxt(meeting, includeReport);
    downloadFile(txt, `${sanitizeFilename(meeting.title)}.txt`, "text/plain;charset=utf-8");
  });

  document.getElementById("btn-export-csv").addEventListener("click", () => {
    const meeting = getSelectedMeeting();
    const csv = exportMeetingToCsv(meeting);
    downloadFile(csv, `${sanitizeFilename(meeting.title)}.csv`, "text/csv;charset=utf-8");
  });

  document.getElementById("btn-generate-report").addEventListener("click", () => {
    triggerRequestsReportGeneration();
  });

  document.getElementById("btn-rerun-ai-report").addEventListener("click", () => {
    triggerRequestsReportGeneration();
  });

  document.getElementById("btn-export-report-txt").addEventListener("click", () => {
    const meeting = getSelectedMeeting();
    const txt = exportRequestsReportToTxt(meeting);
    downloadFile(
      txt,
      `Reporte_Solicitudes_${sanitizeFilename(meeting.title)}.txt`,
      "text/plain;charset=utf-8"
    );
  });

  document.getElementById("btn-export-report-csv").addEventListener("click", () => {
    const meeting = getSelectedMeeting();
    const csv = exportRequestsReportToCsv(meeting);
    downloadFile(
      csv,
      `Reporte_Solicitudes_${sanitizeFilename(meeting.title)}.csv`,
      "text/csv;charset=utf-8"
    );
  });

  document.getElementById("chk-include-report-in-txt").addEventListener("change", () => {
    renderRawTxtPreview();
  });

  document.getElementById("btn-copy-raw-txt").addEventListener("click", async (e) => {
    const raw = document.getElementById("raw-txt-preview").textContent;
    await navigator.clipboard.writeText(raw);
    const original = e.target.textContent;
    e.target.textContent = "✅ ¡Copiado!";
    setTimeout(() => {
      e.target.textContent = original;
    }, 2000);
  });
}

async function triggerRequestsReportGeneration() {
  const meeting = getSelectedMeeting();
  const btn = document.getElementById("btn-generate-report");
  const originalLabel = btn.textContent;
  btn.textContent = "⏳ Analizando Solicitudes con IA...";
  btn.disabled = true;

  try {
    if (typeof chrome !== "undefined" && chrome.runtime?.sendMessage) {
      chrome.runtime.sendMessage(
        {
          type: "GENERATE_REQUESTS_REPORT",
          meetingId: meeting.id,
          aiConfig: state.aiConfig
        },
        (res) => {
          btn.textContent = originalLabel;
          btn.disabled = false;
          if (res?.ok && res.report) {
            meeting.requestsReport = res.report;
            renderAll();
            document.querySelector('.view-tab[data-view="requests"]')?.click();
          }
        }
      );
    } else {
      const report = await generateRequestsReportWithAI(meeting, state.aiConfig);
      meeting.requestsReport = report;
      btn.textContent = originalLabel;
      btn.disabled = false;
      renderAll();
      document.querySelector('.view-tab[data-view="requests"]')?.click();
    }
  } catch (err) {
    btn.textContent = originalLabel;
    btn.disabled = false;
    alert("Error al generar reporte: " + err.message);
  }
}

function renderAll() {
  renderSidebarList(document.getElementById("sidebar-search")?.value || "");
  renderMeetingHeader();
  populateSpeakerDropdown();
  renderTranscriptAndHighlights();
  renderRequestsReportView();
  renderRawTxtPreview();
}

function renderSidebarList(filterQuery = "") {
  const container = document.getElementById("sidebar-meetings-list");
  const q = filterQuery.trim().toLowerCase();

  const filtered = state.meetings.filter((m) => {
    if (!q) return true;
    return (
      (m.title || "").toLowerCase().includes(q) ||
      (m.participants || []).some(p => p.toLowerCase().includes(q))
    );
  });

  container.innerHTML = filtered
    .map(
      (m) => `
    <div class="sidebar-meeting-item ${m.id === state.selectedMeetingId ? "active" : ""}" data-id="${escapeHtml(m.id)}">
      <div class="sidebar-item-title">🐸 ${escapeHtml(m.title)}</div>
      <div class="sidebar-item-meta">
        <span>${escapeHtml(m.startedDisplay || "")}</span>
        <span>${(m.transcript || []).length} líneas</span>
      </div>
    </div>
  `
    )
    .join("");

  container.querySelectorAll(".sidebar-meeting-item").forEach((el) => {
    el.addEventListener("click", () => {
      state.selectedMeetingId = el.dataset.id;
      renderAll();
    });
  });
}

function renderMeetingHeader() {
  const meeting = getSelectedMeeting();
  document.getElementById("main-meeting-title").textContent = meeting.title;

  const metaContainer = document.getElementById("main-meeting-meta");
  metaContainer.innerHTML = `
    <span class="meta-pill">📅 Inicio: ${escapeHtml(meeting.startedDisplay || meeting.startedAt)}</span>
    <span class="meta-pill">⏱ Duración: ${escapeHtml(meeting.durationDisplay || `${meeting.durationMinutes || 0} minutes`)}</span>
    <span class="meta-pill">👥 Participantes: ${(meeting.participants || []).length}</span>
    <span class="meta-pill">⭐ Highlights: ${(meeting.highlights || []).length}</span>
  `;

  const reqCount = meeting.requestsReport?.requests?.length || 0;
  document.getElementById("requests-badge").textContent = String(reqCount);
}

function populateSpeakerDropdown() {
  const meeting = getSelectedMeeting();
  const select = document.getElementById("speaker-filter");
  const currentVal = select.value;
  const speakers = Array.from(
    new Set([
      ...(meeting.participants || []),
      ...(meeting.transcript || []).map(t => t.speaker)
    ])
  ).filter(Boolean);

  select.innerHTML =
    `<option value="">Todos los participantes (${speakers.length})</option>` +
    speakers.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join("");

  if (speakers.includes(currentVal)) {
    select.value = currentVal;
  }
}

function renderTranscriptAndHighlights() {
  const meeting = getSelectedMeeting();
  const searchVal = (document.getElementById("search-in-transcript")?.value || "").trim().toLowerCase();
  const speakerVal = document.getElementById("speaker-filter")?.value || "";

  const highlightSet = new Set(
    (meeting.highlights || []).map(h => `${h.time}|${h.speaker}|${h.text}`)
  );

  const lines = (meeting.transcript || []).filter((line) => {
    if (speakerVal && line.speaker !== speakerVal) return false;
    if (searchVal) {
      return (
        line.text.toLowerCase().includes(searchVal) ||
        line.speaker.toLowerCase().includes(searchVal) ||
        line.time.includes(searchVal)
      );
    }
    return true;
  });

  const transcriptContainer = document.getElementById("transcript-lines-container");
  transcriptContainer.innerHTML = lines
    .map((line, i) => {
      const isHighlighted = Boolean(
        line.highlighted || highlightSet.has(`${line.time}|${line.speaker}|${line.text}`)
      );
      return `
        <div class="t-row ${isHighlighted ? "is-highlighted" : ""}" data-idx="${i}">
          <div class="t-time">${escapeHtml(line.time)}</div>
          <div class="t-content">
            <div class="t-speaker">${escapeHtml(line.speaker)}</div>
            <div class="t-text">${escapeHtml(line.text)}</div>
          </div>
          <button class="t-star" data-time="${escapeHtml(line.time)}" data-speaker="${escapeHtml(line.speaker)}" data-text="${escapeHtml(line.text)}" title="Marcar/Desmarcar como Highlight">
            ${isHighlighted ? "⭐" : "☆"}
          </button>
        </div>
      `;
    })
    .join("");

  transcriptContainer.querySelectorAll(".t-star").forEach((btn) => {
    btn.addEventListener("click", () => {
      const { time, speaker, text } = btn.dataset;
      toggleLocalHighlight(meeting, time, speaker, text);
      if (typeof chrome !== "undefined" && chrome.runtime?.sendMessage) {
        chrome.runtime.sendMessage({
          type: "TOGGLE_HIGHLIGHT",
          meetingId: meeting.id,
          time,
          speaker,
          text
        });
      }
      renderAll();
    });
  });

  // Render Highlights column
  const highlights = meeting.highlights || [];
  document.getElementById("highlights-count").textContent = String(highlights.length);
  const hContainer = document.getElementById("highlights-container");
  if (highlights.length === 0) {
    hContainer.innerHTML = `<p style="color:#64748b;font-size:12px;">Sin highlights seleccionados todavía.</p>`;
  } else {
    hContainer.innerHTML = highlights
      .map(
        (h) => `
      <div class="highlight-quote">
        <strong style="color:#34d399;font-family:monospace;">${escapeHtml(h.time)}</strong>
        <strong style="color:#fde68a;"> ${escapeHtml(h.speaker)}:</strong>
        <span> ${escapeHtml(h.text)}</span>
      </div>
    `
      )
      .join("");
  }

  // Render Participants chips
  const participants = meeting.participants || [];
  document.getElementById("participants-count").textContent = String(participants.length);
  const pContainer = document.getElementById("participants-chips");
  pContainer.innerHTML = participants
    .map(
      (p) => `<span class="participant-chip" data-speaker="${escapeHtml(p)}">👤 ${escapeHtml(p)}</span>`
    )
    .join("");

  pContainer.querySelectorAll(".participant-chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      const select = document.getElementById("speaker-filter");
      select.value = select.value === chip.dataset.speaker ? "" : chip.dataset.speaker;
      renderTranscriptAndHighlights();
    });
  });
}

function renderRequestsReportView() {
  const meeting = getSelectedMeeting();
  const report = meeting.requestsReport;

  const providerMeta = document.getElementById("report-provider-meta");
  const summaryEl = document.getElementById("report-executive-summary");
  const tbody = document.getElementById("requests-table-body");
  const agreementsList = document.getElementById("report-agreements-list");

  if (!report) {
    providerMeta.textContent = "Aún no generado para esta reunión";
    summaryEl.textContent = "Haz clic en '✨ Generar Reporte Solicitudes IA' para extraer todas las solicitudes, responsables y plazos de la reunión.";
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;color:#64748b;padding:20px;">Sin solicitudes analizadas. Pulsa "Re-analizar con IA".</td></tr>`;
    agreementsList.innerHTML = "";
    return;
  }

  providerMeta.textContent = `Analizado con: ${report.provider || "FroggIteAms AI"}`;
  summaryEl.textContent = report.summary || "";

  const requests = report.requests || [];
  tbody.innerHTML = requests
    .map(
      (req) => `
    <tr>
      <td><strong style="color:#a5b4fc;">${escapeHtml(req.id)}</strong></td>
      <td style="font-family:monospace;color:#34d399;">${escapeHtml(req.timestamp)}</td>
      <td>${escapeHtml(req.requester)}</td>
      <td><strong>${escapeHtml(req.assignee)}</strong></td>
      <td>${escapeHtml(req.description)}</td>
      <td><span class="priority-badge priority-${escapeHtml(req.priority)}">${escapeHtml(req.priority)}</span></td>
      <td>${escapeHtml(req.deadline)}</td>
      <td>
        <select class="status-select" data-req-id="${escapeHtml(req.id)}">
          <option value="Pendiente" ${req.status === "Pendiente" ? "selected" : ""}>⏳ Pendiente</option>
          <option value="Programado" ${req.status === "Programado" ? "selected" : ""}>📅 Programado</option>
          <option value="En curso" ${req.status === "En curso" ? "selected" : ""}>🔄 En curso</option>
          <option value="Completado" ${req.status === "Completado" ? "selected" : ""}>✅ Completado</option>
        </select>
      </td>
    </tr>
  `
    )
    .join("");

  tbody.querySelectorAll(".status-select").forEach((sel) => {
    sel.addEventListener("change", () => {
      const reqId = sel.dataset.reqId;
      const targetReq = requests.find(r => r.id === reqId);
      if (targetReq) targetReq.status = sel.value;
      if (typeof chrome !== "undefined" && chrome.runtime?.sendMessage) {
        chrome.runtime.sendMessage({
          type: "UPDATE_REQUEST_STATUS",
          meetingId: meeting.id,
          requestId: reqId,
          status: sel.value
        });
      }
      renderRawTxtPreview();
    });
  });

  agreementsList.innerHTML = (report.keyAgreements || [])
    .map(a => `<li>${escapeHtml(a)}</li>`)
    .join("");
}

function renderRawTxtPreview() {
  const meeting = getSelectedMeeting();
  const includeReport = document.getElementById("chk-include-report-in-txt")?.checked || false;
  document.getElementById("raw-txt-preview").textContent = exportMeetingToTxt(meeting, includeReport);
}

function toggleLocalHighlight(meeting, time, speaker, text) {
  if (!meeting.highlights) meeting.highlights = [];
  const idx = meeting.highlights.findIndex(
    h => h.time === time && h.speaker === speaker && h.text === text
  );
  const line = (meeting.transcript || []).find(
    t => t.time === time && t.speaker === speaker && t.text === text
  );
  if (idx !== -1) {
    meeting.highlights.splice(idx, 1);
    if (line) line.highlighted = false;
  } else {
    meeting.highlights.push({ time, speaker, text });
    if (line) line.highlighted = true;
  }
}

function initAiModal() {
  const backdrop = document.getElementById("ai-modal-backdrop");
  const providerSelect = document.getElementById("modal-ai-provider");

  document.getElementById("btn-open-ai-modal").addEventListener("click", () => {
    populateModalFromConfig(state.aiConfig);
    backdrop.classList.remove("hidden");
  });

  document.getElementById("btn-close-ai-modal").addEventListener("click", () => {
    backdrop.classList.add("hidden");
  });

  providerSelect.addEventListener("change", () => {
    switchModalProviderSection(providerSelect.value);
  });

  document.querySelectorAll('input[name="modal-gemini-mode"]').forEach((r) => {
    r.addEventListener("change", (e) => {
      const isOAuth = e.target.value === "oauth";
      document.getElementById("modal-gemini-apikey-wrap").classList.toggle("hidden", isOAuth);
      document.getElementById("modal-gemini-oauth-wrap").classList.toggle("hidden", !isOAuth);
    });
  });

  document.getElementById("btn-modal-oauth").addEventListener("click", () => {
    const clientId = document.getElementById("modal-gemini-client-id").value.trim();
    const fb = document.getElementById("modal-save-feedback");
    fb.textContent = "Conectando con Google OAuth 2.0...";
    if (typeof chrome !== "undefined" && chrome.runtime?.sendMessage) {
      chrome.runtime.sendMessage({ type: "LAUNCH_GEMINI_OAUTH", clientId }, (res) => {
        if (res?.ok && res.accessToken) {
          document.getElementById("modal-gemini-oauth-token").value = res.accessToken;
          fb.textContent = "✅ Token OAuth de Gemini recibido.";
        } else {
          fb.textContent = `ℹ️ ${res?.error || "Pega tu Bearer Token de Gemini manualmente."}`;
        }
      });
    }
  });

  document.getElementById("btn-save-modal-ai").addEventListener("click", () => {
    const geminiMode =
      document.querySelector('input[name="modal-gemini-mode"]:checked')?.value || "apikey";
    const newConfig = {
      activeProvider: providerSelect.value,
      gemini: {
        authMode: geminiMode,
        apiKey: document.getElementById("modal-gemini-key").value.trim(),
        oauthClientId: document.getElementById("modal-gemini-client-id").value.trim(),
        oauthToken: document.getElementById("modal-gemini-oauth-token").value.trim(),
        model: document.getElementById("modal-gemini-model").value.trim() || "gemini-2.0-flash"
      },
      deepseek: {
        apiKey: document.getElementById("modal-deepseek-key").value.trim(),
        model: document.getElementById("modal-deepseek-model").value.trim() || "deepseek-chat",
        endpoint: "https://api.deepseek.com/chat/completions"
      },
      openai: {
        apiKey: document.getElementById("modal-openai-key").value.trim(),
        model: document.getElementById("modal-openai-model").value.trim() || "gpt-4o-mini",
        endpoint: "https://api.openai.com/v1/chat/completions"
      },
      custom: {
        endpoint: document.getElementById("modal-custom-endpoint").value.trim(),
        apiKey: document.getElementById("modal-custom-key").value.trim(),
        model: document.getElementById("modal-custom-model").value.trim()
      }
    };

    state.aiConfig = newConfig;
    if (typeof chrome !== "undefined" && chrome.runtime?.sendMessage) {
      chrome.runtime.sendMessage({ type: "SAVE_AI_CONFIG", aiConfig: newConfig });
    }
    const fb = document.getElementById("modal-save-feedback");
    fb.textContent = "✅ Configuración guardada.";
    setTimeout(() => {
      fb.textContent = "";
      backdrop.classList.add("hidden");
    }, 900);
  });
}

function switchModalProviderSection(provider) {
  ["gemini", "deepseek", "openai", "custom"].forEach((p) => {
    document.getElementById(`modal-box-${p}`)?.classList.toggle("hidden", p !== provider);
  });
}

function populateModalFromConfig(cfg) {
  const p = cfg.activeProvider || "gemini";
  document.getElementById("modal-ai-provider").value = p;
  switchModalProviderSection(p);

  const gem = cfg.gemini || {};
  const mode = gem.authMode || "apikey";
  const radio = document.querySelector(`input[name="modal-gemini-mode"][value="${mode}"]`);
  if (radio) radio.checked = true;
  document.getElementById("modal-gemini-apikey-wrap").classList.toggle("hidden", mode === "oauth");
  document.getElementById("modal-gemini-oauth-wrap").classList.toggle("hidden", mode !== "oauth");

  document.getElementById("modal-gemini-key").value = gem.apiKey || "";
  document.getElementById("modal-gemini-client-id").value = gem.oauthClientId || "";
  document.getElementById("modal-gemini-oauth-token").value = gem.oauthToken || "";
  document.getElementById("modal-gemini-model").value = gem.model || "gemini-2.0-flash";

  document.getElementById("modal-deepseek-key").value = cfg.deepseek?.apiKey || "";
  document.getElementById("modal-deepseek-model").value = cfg.deepseek?.model || "deepseek-chat";

  document.getElementById("modal-openai-key").value = cfg.openai?.apiKey || "";
  document.getElementById("modal-openai-model").value = cfg.openai?.model || "gpt-4o-mini";

  document.getElementById("modal-custom-endpoint").value =
    cfg.custom?.endpoint || "https://api.groq.com/openai/v1/chat/completions";
  document.getElementById("modal-custom-key").value = cfg.custom?.apiKey || "";
  document.getElementById("modal-custom-model").value =
    cfg.custom?.model || "llama-3.3-70b-versatile";
}

function downloadFile(content, filename, mimeType) {
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
