// content.js
// FroggIteAms In-Page Live Transcriber Widget for Microsoft Teams
// Captures real-time conversations via:
// 1. MS Teams Live Captions DOM MutationObserver (supports both classic & new Teams Web UI)
// 2. Web Speech API (webkitSpeechRecognition) fallback/hybrid mode for direct microphone/audio capture
// 3. Interactive Tactiq-style floating sidebar widget with 1-click Highlight, AI Requests Report, and Export

(function initFroggIteAmsContentScript() {
  if (window.__FROGGITEAMS_INITIALIZED__) return;
  window.__FROGGITEAMS_INITIALIZED__ = true;

  let isCapturing = true;
  let isMinimized = false;
  let useWebSpeechFallback = false;
  let speechRecognition = null;
  let meetingStartedMs = Date.now();
  let currentEntries = [];
  let seenKeys = new Set();

  const iconUrl = chrome.runtime.getURL("assets/icon-48.png");

  // Create Floating Tactiq-like Widget in MS Teams
  const root = document.createElement("div");
  root.id = "froggiteams-widget-root";
  root.innerHTML = `
    <div class="frogg-panel" id="frogg-panel">
      <div class="frogg-header" id="frogg-drag-handle">
        <div class="frogg-brand">
          <img src="${iconUrl}" alt="FroggIteAms" class="frogg-logo" />
          <div class="frogg-title-wrap">
            <span class="frogg-title">FroggIteAms</span>
            <span class="frogg-status" id="frogg-status-badge">
              <span class="frogg-dot live"></span> En Vivo · MS Teams
            </span>
          </div>
        </div>
        <div class="frogg-header-actions">
          <button class="frogg-icon-btn" id="frogg-btn-mic" title="Activar/Desactivar Dictado de Voz WebSpeech">🎙️</button>
          <button class="frogg-icon-btn" id="frogg-btn-dashboard" title="Abrir Menú Histórico y Reportes (Tactiq Hub)">📊</button>
          <button class="frogg-icon-btn" id="frogg-btn-minimize" title="Minimizar/Expandir">_</button>
        </div>
      </div>

      <div class="frogg-body" id="frogg-body">
        <div class="frogg-meeting-bar">
          <input type="text" id="frogg-meeting-title" class="frogg-meeting-input" value="${detectTeamsMeetingTitle()}" placeholder="Título de la reunión..." />
          <span class="frogg-timer" id="frogg-timer">00:00</span>
        </div>

        <div class="frogg-stream" id="frogg-stream">
          <div class="frogg-empty" id="frogg-empty">
            <p>🐸 <strong>Escuchando subtítulos de MS Teams...</strong></p>
            <p class="frogg-hint">Activa <em>Subtítulos en vivo (Live Captions)</em> en Teams o pulsa 🎙️ para transcribir con micrófono o simular entrada.</p>
          </div>
        </div>

        <div class="frogg-quick-input">
          <input type="text" id="frogg-manual-speaker" class="frogg-speaker-input" placeholder="Hablante" value="Tú" />
          <input type="text" id="frogg-manual-text" class="frogg-text-input" placeholder="Nota rápida o simula frase..." />
          <button id="frogg-btn-add-line" class="frogg-btn-small" title="Agregar a transcripción">+</button>
        </div>

        <div class="frogg-footer">
          <button class="frogg-btn frogg-btn-secondary" id="frogg-btn-pause">⏸ Pausar</button>
          <button class="frogg-btn frogg-btn-accent" id="frogg-btn-highlight-last">⭐ Highlight</button>
          <button class="frogg-btn frogg-btn-primary" id="frogg-btn-finish">💾 Guardar y Reporte</button>
        </div>
      </div>
    </div>
  `;

  document.body.appendChild(root);

  // Sync initial active meeting state from background
  chrome.runtime.sendMessage({ type: "GET_STATE" }, (res) => {
    if (res?.activeMeeting?.transcript?.length) {
      currentEntries = res.activeMeeting.transcript;
      if (res.activeMeeting.title) {
        document.getElementById("frogg-meeting-title").value = res.activeMeeting.title;
      }
      renderStream();
    } else {
      chrome.runtime.sendMessage({
        type: "START_MEETING",
        title: detectTeamsMeetingTitle(),
        sourceUrl: window.location.href
      });
    }
  });

  // Timer interval
  setInterval(() => {
    if (!isCapturing) return;
    const elapsed = Math.floor((Date.now() - meetingStartedMs) / 1000);
    const mm = String(Math.floor(elapsed / 60)).padStart(2, "0");
    const ss = String(elapsed % 60).padStart(2, "0");
    const timerEl = document.getElementById("frogg-timer");
    if (timerEl) timerEl.textContent = `${mm}:${ss}`;
  }, 1000);

  // Event Listeners
  document.getElementById("frogg-btn-minimize").addEventListener("click", () => {
    isMinimized = !isMinimized;
    const body = document.getElementById("frogg-body");
    body.style.display = isMinimized ? "none" : "flex";
  });

  document.getElementById("frogg-btn-dashboard").addEventListener("click", () => {
    window.open(chrome.runtime.getURL("dashboard.html"), "_blank");
  });

  document.getElementById("frogg-btn-pause").addEventListener("click", (e) => {
    isCapturing = !isCapturing;
    const statusBadge = document.getElementById("frogg-status-badge");
    if (isCapturing) {
      e.target.textContent = "⏸ Pausar";
      statusBadge.innerHTML = `<span class="frogg-dot live"></span> En Vivo · MS Teams`;
    } else {
      e.target.textContent = "▶ Reanudar";
      statusBadge.innerHTML = `<span class="frogg-dot paused"></span> En Pausa`;
    }
  });

  document.getElementById("frogg-btn-highlight-last").addEventListener("click", () => {
    if (currentEntries.length === 0) return;
    const last = currentEntries[currentEntries.length - 1];
    last.highlighted = !last.highlighted;
    chrome.runtime.sendMessage({
      type: "TOGGLE_HIGHLIGHT",
      time: last.time,
      speaker: last.speaker,
      text: last.text
    });
    renderStream();
  });

  document.getElementById("frogg-btn-finish").addEventListener("click", () => {
    chrome.runtime.sendMessage({ type: "FINISH_MEETING" }, () => {
      window.open(chrome.runtime.getURL("dashboard.html"), "_blank");
    });
  });

  document.getElementById("frogg-btn-add-line").addEventListener("click", addManualEntry);
  document.getElementById("frogg-manual-text").addEventListener("keydown", (e) => {
    if (e.key === "Enter") addManualEntry();
  });

  document.getElementById("frogg-btn-mic").addEventListener("click", toggleWebSpeechRecognition);

  function addManualEntry() {
    const speakerInput = document.getElementById("frogg-manual-speaker");
    const textInput = document.getElementById("frogg-manual-text");
    const speaker = (speakerInput.value || "Participante").trim();
    const text = (textInput.value || "").trim();
    if (!text) return;
    pushCaptionLine(speaker, text, false);
    textInput.value = "";
  }

  function pushCaptionLine(speaker, text, isUpdate = false) {
    if (!isCapturing || !text) return;
    const meetingTitle = document.getElementById("frogg-meeting-title")?.value || detectTeamsMeetingTitle();
    const elapsedSec = Math.floor((Date.now() - meetingStartedMs) / 1000);
    const time = `${String(Math.floor(elapsedSec / 60)).padStart(2, "0")}:${String(elapsedSec % 60).padStart(2, "0")}`;

    chrome.runtime.sendMessage(
      {
        type: "CAPTION_ENTRY",
        speaker,
        text,
        time,
        meetingTitle,
        isUpdate,
        useAiRefinement: true
      },
      (res) => {
        if (res?.activeMeeting?.transcript) {
          currentEntries = res.activeMeeting.transcript;
          renderStream();
        }
      }
    );
  }

  function renderStream() {
    const stream = document.getElementById("frogg-stream");
    if (!stream) return;
    if (currentEntries.length === 0) return;

    const recent = currentEntries.slice(-40);
    stream.innerHTML = recent
      .map(
        (item) => `
      <div class="frogg-line ${item.highlighted ? "highlighted" : ""}" data-time="${escapeHtml(item.time)}" data-speaker="${escapeHtml(item.speaker)}">
        <div class="frogg-line-meta">
          <span class="frogg-line-time">${escapeHtml(item.time)}</span>
          <span class="frogg-line-speaker">${escapeHtml(item.speaker)}</span>
          <button class="frogg-star-btn" title="Destacar en Highlights">${item.highlighted ? "⭐" : "☆"}</button>
        </div>
        <div class="frogg-line-text">${escapeHtml(item.text)}</div>
      </div>
    `
      )
      .join("");

    stream.querySelectorAll(".frogg-star-btn").forEach((btn, idx) => {
      btn.addEventListener("click", () => {
        const entry = recent[idx];
        entry.highlighted = !entry.highlighted;
        chrome.runtime.sendMessage({
          type: "TOGGLE_HIGHLIGHT",
          time: entry.time,
          speaker: entry.speaker,
          text: entry.text
        });
        renderStream();
      });
    });

    stream.scrollTop = stream.scrollHeight;
  }

  // Observe MS Teams Live Captions DOM (supports Teams v2 Fluent UI & Classic Teams)
  const CAPTION_SELECTORS = [
    '[data-tid="closed-captions-renderer"]',
    '[data-tid="closed-caption-v2-window-wrapper"]',
    '[data-tid="closed-caption-item"]',
    '[class*="fui-ChatMessage"]',
    '[role="log"][aria-label*="caption" i]',
    '[role="region"][aria-label*="subtítulo" i]',
    '[role="region"][aria-label*="caption" i]'
  ];

  const observer = new MutationObserver(() => {
    if (!isCapturing) return;
    scanTeamsDomForCaptions();
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true,
    characterData: true
  });

  function scanTeamsDomForCaptions() {
    // Check explicit caption items in MS Teams
    const items = document.querySelectorAll(
      '[data-tid="closed-caption-item"], [data-tid*="caption-item"], [class*="closed-caption"] [role="listitem"]'
    );

    items.forEach((node) => {
      const authorEl =
        node.querySelector('[data-tid="author"], [class*="author"], [class*="speaker"], strong') ||
        node.previousElementSibling;
      const textEl =
        node.querySelector('[data-tid="closed-caption-text"], [class*="caption-text"], [dir="auto"]') ||
        node;

      const speaker = (authorEl?.textContent || "Participante Teams").trim();
      const text = (textEl?.textContent || "").replace(speaker, "").trim();

      if (text && text.length > 2) {
        const key = `${speaker}:${text}`;
        if (!seenKeys.has(key)) {
          seenKeys.add(key);
          pushCaptionLine(speaker, text, true);
        }
      }
    });
  }

  // Optional Web Speech API real-time microphone/audio fallback
  function toggleWebSpeechRecognition() {
    const micBtn = document.getElementById("frogg-btn-mic");
    const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRec) {
      alert("Tu navegador no soporta Web Speech API directa. Usa Subtítulos en vivo de MS Teams.");
      return;
    }

    if (useWebSpeechFallback && speechRecognition) {
      useWebSpeechFallback = false;
      speechRecognition.stop();
      micBtn.classList.remove("active");
      return;
    }

    useWebSpeechFallback = true;
    micBtn.classList.add("active");
    speechRecognition = new SpeechRec();
    speechRecognition.lang = "es-CL";
    speechRecognition.continuous = true;
    speechRecognition.interimResults = false;

    speechRecognition.onresult = (event) => {
      for (let i = event.resultIndex; i < event.results.length; i++) {
        if (event.results[i].isFinal) {
          const phrase = event.results[i][0].transcript.trim();
          const speaker = document.getElementById("frogg-manual-speaker")?.value || "Participante (Voz)";
          if (phrase) pushCaptionLine(speaker, phrase, false);
        }
      }
    };

    speechRecognition.onend = () => {
      if (useWebSpeechFallback) {
        try {
          speechRecognition.start();
        } catch {}
      }
    };

    try {
      speechRecognition.start();
    } catch (e) {
      console.warn("[FroggIteAms] WebSpeech error:", e);
    }
  }

  // Make widget draggable
  makeDraggable(root, document.getElementById("frogg-drag-handle"));

  function detectTeamsMeetingTitle() {
    const raw = document.title || "";
    const cleaned = raw
      .replace(/\|?\s*Microsoft Teams/i, "")
      .replace(/^\(\d+\)\s*/, "")
      .trim();
    return cleaned || "Reunión Microsoft Teams";
  }

  function makeDraggable(element, handle) {
    let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
    handle.onmousedown = dragMouseDown;

    function dragMouseDown(e) {
      if (e.target.tagName === "BUTTON" || e.target.tagName === "INPUT") return;
      e.preventDefault();
      pos3 = e.clientX;
      pos4 = e.clientY;
      document.onmouseup = closeDragElement;
      document.onmousemove = elementDrag;
    }

    function elementDrag(e) {
      e.preventDefault();
      pos1 = pos3 - e.clientX;
      pos2 = pos4 - e.clientY;
      pos3 = e.clientX;
      pos4 = e.clientY;
      element.style.top = (element.offsetTop - pos2) + "px";
      element.style.left = (element.offsetLeft - pos1) + "px";
      element.style.right = "auto";
      element.style.bottom = "auto";
    }

    function closeDragElement() {
      document.onmouseup = null;
      document.onmousemove = null;
    }
  }

  function escapeHtml(str) {
    return String(str || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
})();
