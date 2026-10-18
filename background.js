// background.js
// Service Worker for FroggIteAms Chrome Extension
// Coordinates live Microsoft Teams transcription state, persistent meeting history,
// Google OAuth 2.0 flow for Gemini, and multi-provider AI report generation.

import { SAMPLE_MEETING } from "./lib/sample-data.js";
import {
  DEFAULT_AI_CONFIG,
  generateRequestsReportWithAI,
  refineLiveCaptionWithAI
} from "./lib/ai-providers.js";

const STORAGE_KEYS = {
  MEETINGS: "froggiteams_meetings",
  ACTIVE_MEETING: "froggiteams_active_meeting",
  AI_CONFIG: "froggiteams_ai_config"
};

// Initialize storage with sample meeting on install if empty
chrome.runtime.onInstalled.addListener(async () => {
  const data = await chrome.storage.local.get([
    STORAGE_KEYS.MEETINGS,
    STORAGE_KEYS.AI_CONFIG
  ]);

  if (!data[STORAGE_KEYS.MEETINGS] || data[STORAGE_KEYS.MEETINGS].length === 0) {
    await chrome.storage.local.set({
      [STORAGE_KEYS.MEETINGS]: [SAMPLE_MEETING]
    });
  }

  if (!data[STORAGE_KEYS.AI_CONFIG]) {
    await chrome.storage.local.set({
      [STORAGE_KEYS.AI_CONFIG]: DEFAULT_AI_CONFIG
    });
  }
});

// Listen for messages from content.js, popup.js, and dashboard.js
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  handleMessage(message, sender)
    .then(result => sendResponse({ ok: true, ...result }))
    .catch(err => {
      console.error("[FroggIteAms Background Error]", err);
      sendResponse({ ok: false, error: err.message || String(err) });
    });
  return true; // keep channel open for async response
});

async function handleMessage(message, sender) {
  switch (message.type) {
    case "GET_STATE": {
      const store = await chrome.storage.local.get([
        STORAGE_KEYS.MEETINGS,
        STORAGE_KEYS.ACTIVE_MEETING,
        STORAGE_KEYS.AI_CONFIG
      ]);
      let meetings = store[STORAGE_KEYS.MEETINGS];
      if (!meetings || meetings.length === 0) {
        meetings = [SAMPLE_MEETING];
        await chrome.storage.local.set({ [STORAGE_KEYS.MEETINGS]: meetings });
      }
      return {
        meetings,
        activeMeeting: store[STORAGE_KEYS.ACTIVE_MEETING] || null,
        aiConfig: store[STORAGE_KEYS.AI_CONFIG] || DEFAULT_AI_CONFIG
      };
    }

    case "START_MEETING": {
      const now = new Date();
      const meeting = {
        id: message.meetingId || `teams-${now.getTime()}`,
        title: message.title || "Reunión Microsoft Teams",
        startedAt: now.toISOString(),
        startedDisplay: formatTactiqStartTime(now),
        durationMinutes: 1,
        durationDisplay: "1 minutes",
        sourceUrl: message.sourceUrl || sender?.tab?.url || "https://teams.microsoft.com/",
        participants: message.participants || [],
        highlights: [],
        transcript: [],
        requestsReport: null,
        isLive: true
      };
      await chrome.storage.local.set({ [STORAGE_KEYS.ACTIVE_MEETING]: meeting });
      return { activeMeeting: meeting };
    }

    case "CAPTION_ENTRY": {
      const store = await chrome.storage.local.get([
        STORAGE_KEYS.ACTIVE_MEETING,
        STORAGE_KEYS.AI_CONFIG
      ]);
      let active = store[STORAGE_KEYS.ACTIVE_MEETING];
      const aiConfig = store[STORAGE_KEYS.AI_CONFIG] || DEFAULT_AI_CONFIG;

      if (!active) {
        const now = new Date();
        active = {
          id: `teams-${now.getTime()}`,
          title: message.meetingTitle || "Reunión Microsoft Teams en Vivo",
          startedAt: now.toISOString(),
          startedDisplay: formatTactiqStartTime(now),
          durationMinutes: 1,
          durationDisplay: "1 minutes",
          sourceUrl: sender?.tab?.url || "https://teams.microsoft.com/",
          participants: [],
          highlights: [],
          transcript: [],
          requestsReport: null,
          isLive: true
        };
      }

      if (message.meetingTitle && message.meetingTitle !== "Microsoft Teams") {
        active.title = message.meetingTitle;
      }

      const speaker = (message.speaker || "Participante").trim();
      let text = (message.text || "").trim();
      if (!text) return { activeMeeting: active };

      if (message.useAiRefinement) {
        text = await refineLiveCaptionWithAI(speaker, text, aiConfig);
      }

      if (!active.participants.includes(speaker)) {
        active.participants.push(speaker);
      }

      const elapsedMs = Math.max(0, Date.now() - new Date(active.startedAt).getTime());
      const timeStr = message.time || formatElapsedMMSS(elapsedMs);
      const durationMin = Math.max(1, Math.round(elapsedMs / 60000));
      active.durationMinutes = durationMin;
      active.durationDisplay = `${durationMin} minutes`;

      const lastEntry = active.transcript[active.transcript.length - 1];
      // Merge or update if same speaker within recent window and text is a continuation
      if (
        lastEntry &&
        lastEntry.speaker === speaker &&
        (message.isUpdate || text.startsWith(lastEntry.text.slice(0, 15)))
      ) {
        lastEntry.text = text;
        if (message.highlighted) {
          lastEntry.highlighted = true;
          syncHighlight(active, lastEntry);
        }
      } else {
        const newEntry = {
          id: `line-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
          time: timeStr,
          speaker,
          text,
          highlighted: Boolean(message.highlighted)
        };
        active.transcript.push(newEntry);
        if (newEntry.highlighted) {
          syncHighlight(active, newEntry);
        }
      }

      await chrome.storage.local.set({ [STORAGE_KEYS.ACTIVE_MEETING]: active });
      return { activeMeeting: active };
    }

    case "TOGGLE_HIGHLIGHT": {
      const store = await chrome.storage.local.get([
        STORAGE_KEYS.ACTIVE_MEETING,
        STORAGE_KEYS.MEETINGS
      ]);
      const { meetingId, time, speaker, text } = message;

      let active = store[STORAGE_KEYS.ACTIVE_MEETING];
      let meetings = store[STORAGE_KEYS.MEETINGS] || [];

      if (active && (!meetingId || active.id === meetingId)) {
        toggleHighlightOnMeeting(active, time, speaker, text);
        await chrome.storage.local.set({ [STORAGE_KEYS.ACTIVE_MEETING]: active });
      }

      const idx = meetings.findIndex(m => m.id === meetingId);
      if (idx !== -1) {
        toggleHighlightOnMeeting(meetings[idx], time, speaker, text);
        await chrome.storage.local.set({ [STORAGE_KEYS.MEETINGS]: meetings });
      }

      return { activeMeeting: active, meetings };
    }

    case "FINISH_MEETING": {
      const store = await chrome.storage.local.get([
        STORAGE_KEYS.ACTIVE_MEETING,
        STORAGE_KEYS.MEETINGS,
        STORAGE_KEYS.AI_CONFIG
      ]);
      const active = store[STORAGE_KEYS.ACTIVE_MEETING];
      let meetings = store[STORAGE_KEYS.MEETINGS] || [];
      const aiConfig = store[STORAGE_KEYS.AI_CONFIG] || DEFAULT_AI_CONFIG;

      if (!active) {
        return { meetings };
      }

      active.isLive = false;
      if (active.transcript.length > 0 && !active.requestsReport) {
        try {
          active.requestsReport = await generateRequestsReportWithAI(active, aiConfig);
        } catch (e) {
          console.warn("Could not auto-generate report on finish:", e);
        }
      }

      const existingIdx = meetings.findIndex(m => m.id === active.id);
      if (existingIdx !== -1) {
        meetings[existingIdx] = active;
      } else {
        meetings.unshift(active);
      }

      await chrome.storage.local.set({
        [STORAGE_KEYS.MEETINGS]: meetings,
        [STORAGE_KEYS.ACTIVE_MEETING]: null
      });

      return { savedMeeting: active, meetings };
    }

    case "GENERATE_REQUESTS_REPORT": {
      const store = await chrome.storage.local.get([
        STORAGE_KEYS.MEETINGS,
        STORAGE_KEYS.ACTIVE_MEETING,
        STORAGE_KEYS.AI_CONFIG
      ]);
      const meetings = store[STORAGE_KEYS.MEETINGS] || [];
      const active = store[STORAGE_KEYS.ACTIVE_MEETING];
      const aiConfig = message.aiConfig || store[STORAGE_KEYS.AI_CONFIG] || DEFAULT_AI_CONFIG;

      let target = meetings.find(m => m.id === message.meetingId);
      let isFromActive = false;
      if (!target && active && active.id === message.meetingId) {
        target = active;
        isFromActive = true;
      }
      if (!target) {
        throw new Error("Reunión no encontrada para generar el reporte de solicitudes.");
      }

      const report = await generateRequestsReportWithAI(target, aiConfig);
      target.requestsReport = report;

      // Also merge suggested highlights if any
      if (Array.isArray(report.suggestedHighlights) && report.suggestedHighlights.length > 0 && (!target.highlights || target.highlights.length === 0)) {
        target.highlights = report.suggestedHighlights;
      }

      if (isFromActive) {
        await chrome.storage.local.set({ [STORAGE_KEYS.ACTIVE_MEETING]: target });
      } else {
        await chrome.storage.local.set({ [STORAGE_KEYS.MEETINGS]: meetings });
      }

      return { report, meeting: target, meetings };
    }

    case "SAVE_AI_CONFIG": {
      const newConfig = message.aiConfig || DEFAULT_AI_CONFIG;
      await chrome.storage.local.set({ [STORAGE_KEYS.AI_CONFIG]: newConfig });
      return { aiConfig: newConfig };
    }

    case "LAUNCH_GEMINI_OAUTH": {
      // Supports either chrome.identity.getAuthToken or launchWebAuthFlow with user-provided Client ID
      const clientId = message.clientId?.trim();
      if (clientId) {
        const redirectUri = chrome.identity.getRedirectURL("oauth2");
        const scope = encodeURIComponent(
          "https://www.googleapis.com/auth/generative-language.retriever https://www.googleapis.com/auth/cloud-platform"
        );
        const authUrl = `https://accounts.google.com/o/oauth2/v2/auth?client_id=${encodeURIComponent(clientId)}&response_type=token&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${scope}&prompt=consent`;

        const responseUrl = await chrome.identity.launchWebAuthFlow({
          url: authUrl,
          interactive: true
        });
        const hashParams = new URLSearchParams(new URL(responseUrl).hash.slice(1));
        const accessToken = hashParams.get("access_token");
        if (!accessToken) {
          throw new Error("No se recibió access_token en el flujo OAuth.");
        }
        return { accessToken, redirectUri };
      } else {
        // Attempt native Chrome profile identity token
        const tokenResult = await chrome.identity.getAuthToken({ interactive: true });
        const token = typeof tokenResult === "string" ? tokenResult : tokenResult?.token;
        if (!token) {
          throw new Error("Para usar Google OAuth, ingresa tu Google OAuth Client ID o pega tu Bearer Token de Gemini.");
        }
        return { accessToken: token };
      }
    }

    case "IMPORT_TXT_MEETING": {
      const parsedMeeting = parseTactiqTxtFile(message.content, message.filename);
      const store = await chrome.storage.local.get([STORAGE_KEYS.MEETINGS]);
      const meetings = store[STORAGE_KEYS.MEETINGS] || [];
      meetings.unshift(parsedMeeting);
      await chrome.storage.local.set({ [STORAGE_KEYS.MEETINGS]: meetings });
      return { meeting: parsedMeeting, meetings };
    }

    case "DELETE_MEETING": {
      const store = await chrome.storage.local.get([STORAGE_KEYS.MEETINGS]);
      const meetings = (store[STORAGE_KEYS.MEETINGS] || []).filter(m => m.id !== message.meetingId);
      await chrome.storage.local.set({ [STORAGE_KEYS.MEETINGS]: meetings });
      return { meetings };
    }

    case "UPDATE_REQUEST_STATUS": {
      const store = await chrome.storage.local.get([STORAGE_KEYS.MEETINGS]);
      const meetings = store[STORAGE_KEYS.MEETINGS] || [];
      const meeting = meetings.find(m => m.id === message.meetingId);
      if (meeting?.requestsReport?.requests) {
        const req = meeting.requestsReport.requests.find(r => r.id === message.requestId);
        if (req) {
          req.status = message.status;
          await chrome.storage.local.set({ [STORAGE_KEYS.MEETINGS]: meetings });
        }
      }
      return { meeting, meetings };
    }

    default:
      return {};
  }
}

function syncHighlight(meeting, entry) {
  if (!meeting.highlights) meeting.highlights = [];
  const exists = meeting.highlights.some(
    h => h.time === entry.time && h.speaker === entry.speaker && h.text === entry.text
  );
  if (!exists) {
    meeting.highlights.push({
      time: entry.time,
      speaker: entry.speaker,
      text: entry.text
    });
  }
}

function toggleHighlightOnMeeting(meeting, time, speaker, text) {
  if (!meeting.highlights) meeting.highlights = [];
  const line = (meeting.transcript || []).find(
    t => t.time === time && t.speaker === speaker && t.text === text
  );
  const hIdx = meeting.highlights.findIndex(
    h => h.time === time && h.speaker === speaker && h.text === text
  );

  if (hIdx !== -1) {
    meeting.highlights.splice(hIdx, 1);
    if (line) line.highlighted = false;
  } else {
    meeting.highlights.push({ time, speaker, text });
    if (line) line.highlighted = true;
  }
}

function formatElapsedMMSS(ms) {
  const totalSec = Math.floor(ms / 1000);
  const mm = String(Math.floor(totalSec / 60)).padStart(2, "0");
  const ss = String(totalSec % 60).padStart(2, "0");
  return `${mm}:${ss}`;
}

function formatTactiqStartTime(d) {
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  let hours = d.getHours();
  const minutes = String(d.getMinutes()).padStart(2, "0");
  const seconds = String(d.getSeconds()).padStart(2, "0");
  const ampm = hours >= 12 ? "p.m." : "a.m.";
  hours = hours % 12;
  if (hours === 0) hours = 12;
  const hoursStr = String(hours).padStart(2, "0");
  return `${day}/${month}/${year}, ${hoursStr}:${minutes}:${seconds} ${ampm}`;
}

function parseTactiqTxtFile(rawText, filename = "Reunión Importada") {
  const lines = String(rawText || "").split(/\r?\n/);
  let title = filename.replace(/\.txt$/i, "");
  let startedDisplay = formatTactiqStartTime(new Date());
  let durationDisplay = "30 minutes";
  let participants = [];
  const highlights = [];
  const transcript = [];

  let section = "header";
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) continue;

    if (line.startsWith("# ")) {
      title = line.slice(2).trim();
      continue;
    }
    if (line.startsWith("Meeting started:")) {
      startedDisplay = line.replace("Meeting started:", "").trim();
      continue;
    }
    if (line.startsWith("Duration:")) {
      durationDisplay = line.replace("Duration:", "").trim();
      continue;
    }
    if (line.startsWith("Participants:")) {
      participants = line
        .replace("Participants:", "")
        .split(",")
        .map(s => s.trim())
        .filter(Boolean);
      continue;
    }
    if (line.startsWith("## Highlights")) {
      section = "highlights";
      continue;
    }
    if (line.startsWith("## Transcript")) {
      section = "transcript";
      continue;
    }

    if (section === "highlights" && line.startsWith(">")) {
      const m = line.replace(/^>\s*/, "").match(/^(\d{2}:\d{2})\s+([^:]+):\s*(.+)$/);
      if (m) {
        highlights.push({ time: m[1], speaker: m[2].trim(), text: m[3].trim() });
      }
    } else if (section === "transcript") {
      const m = line.match(/^(\d{2}:\d{2})\s+([^:]+):\s*(.+)$/);
      if (m) {
        const isHighlighted = highlights.some(
          h => h.time === m[1] && h.speaker === m[2].trim()
        );
        transcript.push({
          time: m[1],
          speaker: m[2].trim(),
          text: m[3].trim(),
          highlighted: isHighlighted
        });
      }
    }
  }

  return {
    id: `import-${Date.now()}`,
    title,
    startedAt: new Date().toISOString(),
    startedDisplay,
    durationMinutes: parseInt(durationDisplay, 10) || 30,
    durationDisplay,
    sourceUrl: "https://teams.microsoft.com/imported",
    participants: participants.length > 0
      ? participants
      : Array.from(new Set(transcript.map(t => t.speaker))),
    highlights,
    transcript,
    requestsReport: null
  };
}
