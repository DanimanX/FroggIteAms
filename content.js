// content.js
// FroggIteAms In-Page Live Transcriber Widget for Microsoft Teams
// Supports:
// - https://teams.cloud.microsoft/* (New MS Teams v2 Unified Cloud App)
// - https://teams.microsoft.com/*
// - https://teams.live.com/*
// Captures real-time conversations via:
// 1. MS Teams v2 & Classic Live Captions / Live Transcript DOM Scanner (with Shadow DOM traversal)
// 2. Auto-Caption Enabler button (attempts to turn on Teams Live Captions automatically)
// 3. Direct Microphone Web Speech API (webkitSpeechRecognition) continuous capture
// 4. Interactive Tactiq-style floating sidebar widget with 1-click Highlight, AI Requests Report, and Export

(function initFroggIteAmsContentScript() {
  // Always run caption scanner in all frames, but only mount the floating UI in the top frame
  const isTopFrame = window === window.top;
  if (window.__FROGGITEAMS_INITIALIZED__) return;
  window.__FROGGITEAMS_INITIALIZED__ = true;

  let isCapturing = true;
  let isMinimized = false;
  let useWebSpeechFallback = false;
  let speechRecognition = null;
  let meetingStartedMs = Date.now();
  let currentEntries = [];
  let lastSpeakerTextMap = new Map(); // speaker -> last committed text

  const iconUrl = chrome.runtime.getURL("assets/icon-48.png");

  if (isTopFrame) {
    mountFloatingWidget();
  }

  // Start continuous polling + MutationObserver for MS Teams v2 (teams.cloud.microsoft) captions
  const observer = new MutationObserver(() => {
    if (!isCapturing) return;
    scanTeamsCaptions();
  });

  if (document.body) {
    observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: true
    });
  }

  // Polling every 900ms catches Shadow DOM updates and Fluent UI virtualized lists in teams.cloud.microsoft
  setInterval(() => {
    if (!isCapturing) return;
    scanTeamsCaptions();
  }, 900);

  function mountFloatingWidget() {
    const existing = document.getElementById("froggiteams-widget-root");
    if (existing) existing.remove();

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
            <button class="frogg-icon-btn" id="frogg-btn-captions" title="Activar Subtítulos en Vivo de Teams (Live Captions)">CC</button>
            <button class="frogg-icon-btn" id="frogg-btn-mic" title="Transcribir con Micrófono Directo (WebSpeech)">🎙️</button>
            <button class="frogg-icon-btn" id="frogg-btn-dashboard" title="Abrir Menú Histórico y Reportes (Tactiq Hub)">📊</button>
            <button class="frogg-icon-btn" id="frogg-btn-minimize" title="Minimizar/Expandir">_</button>
          </div>
        </div>

        <div class="frogg-body" id="frogg-body">
          <div class="frogg-meeting-bar">
            <input type="text" id="frogg-meeting-title" class="frogg-meeting-input" value="${detectTeamsMeetingTitle()}" placeholder="Título de la reunión..." />
            <span class="frogg-timer" id="frogg-timer">00:00</span>
          </div>

          <div class="frogg-mode-banner" id="frogg-mode-banner">
            <span>💡 Para capturar a todos en Teams: activa <strong>CC (Subtítulos)</strong> o pulsa <strong>🎙️ Mic</strong> para dictado directo.</span>
          </div>

          <div class="frogg-stream" id="frogg-stream">
            <div class="frogg-empty" id="frogg-empty">
              <p>🐸 <strong>Listo en teams.cloud.microsoft</strong></p>
              <p class="frogg-hint">Haz clic en el botón <strong>CC</strong> arriba para activar los Subtítulos en vivo de Teams o en <strong>🎙️</strong> para transcribir tu micrófono al instante.</p>
            </div>
          </div>

          <div class="frogg-quick-input">
            <input type="text" id="frogg-manual-speaker" class="frogg-speaker-input" placeholder="Hablante" value="Tú" />
            <input type="text" id="frogg-manual-text" class="frogg-text-input" placeholder="Escribe nota o frase en vivo..." />
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

    // Sync initial state from background
    chrome.runtime.sendMessage({ type: "GET_STATE" }, (res) => {
      if (res?.activeMeeting?.transcript?.length) {
        currentEntries = res.activeMeeting.transcript;
        if (res.activeMeeting.title) {
          const titleInput = document.getElementById("frogg-meeting-title");
          if (titleInput) titleInput.value = res.activeMeeting.title;
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

    // Listen for live updates pushed from subframes or popup
    chrome.storage.onChanged.addListener((changes, area) => {
      if (area === "local" && changes.froggiteams_active_meeting?.newValue) {
        const updated = changes.froggiteams_active_meeting.newValue;
        currentEntries = updated.transcript || [];
        renderStream();
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

    // UI Controls
    document.getElementById("frogg-btn-minimize").addEventListener("click", () => {
      isMinimized = !isMinimized;
      const body = document.getElementById("frogg-body");
      body.style.display = isMinimized ? "none" : "flex";
    });

    document.getElementById("frogg-btn-dashboard").addEventListener("click", () => {
      window.open(chrome.runtime.getURL("dashboard.html"), "_blank");
    });

    document.getElementById("frogg-btn-captions").addEventListener("click", () => {
      tryEnableTeamsLiveCaptions();
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

    makeDraggable(root, document.getElementById("frogg-drag-handle"));
  }

  function addManualEntry() {
    const speakerInput = document.getElementById("frogg-manual-speaker");
    const textInput = document.getElementById("frogg-manual-text");
    const speaker = (speakerInput?.value || "Participante").trim();
    const text = (textInput?.value || "").trim();
    if (!text) return;
    pushCaptionLine(speaker, text, false);
    textInput.value = "";
  }

  function pushCaptionLine(speaker, text, isUpdate = false) {
    if (!isCapturing || !text) return;
    const cleanSpeaker = (speaker || "Participante Teams").replace(/\s+/g, " ").trim();
    const cleanText = text.replace(/\s+/g, " ").trim();
    if (cleanText.length < 2) return;

    const prevText = lastSpeakerTextMap.get(cleanSpeaker);
    if (prevText === cleanText) return;
    lastSpeakerTextMap.set(cleanSpeaker, cleanText);

    const meetingTitle =
      document.getElementById("frogg-meeting-title")?.value || detectTeamsMeetingTitle();
    const elapsedSec = Math.floor((Date.now() - meetingStartedMs) / 1000);
    const time = `${String(Math.floor(elapsedSec / 60)).padStart(2, "0")}:${String(
      elapsedSec % 60
    ).padStart(2, "0")}`;

    chrome.runtime.sendMessage(
      {
        type: "CAPTION_ENTRY",
        speaker: cleanSpeaker,
        text: cleanText,
        time,
        meetingTitle,
        isUpdate,
        useAiRefinement: false
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

    const recent = currentEntries.slice(-45);
    stream.innerHTML = recent
      .map(
        (item) => `
      <div class="frogg-line ${item.highlighted ? "highlighted" : ""}">
        <div class="frogg-line-meta">
          <span class="frogg-line-time">${escapeHtml(item.time)}</span>
          <span class="frogg-line-speaker">${escapeHtml(item.speaker)}</span>
          <button class="frogg-star-btn" title="Destacar en Highlights">${
            item.highlighted ? "⭐" : "☆"
          }</button>
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

  /**
   * Comprehensive DOM & Shadow DOM caption extractor for:
   * - teams.cloud.microsoft (Fluent UI v9 / React 18 virtualized captions & transcript pane)
   * - teams.microsoft.com
   * - teams.live.com
   */
  function scanTeamsCaptions() {
    const candidateSelectors = [
      // Teams v2 (teams.cloud.microsoft) Live Captions & Transcript items
      '[data-tid="closed-caption-item"]',
      '[data-tid="closed-captions-renderer"] [role="listitem"]',
      '[data-tid="closed-caption-v2-window-wrapper"] [role="listitem"]',
      '[data-tid*="caption-item"]',
      '[data-tid*="transcript-item"]',
      '[data-tid="call-transcript-panel"] [role="listitem"]',
      '[aria-label*="subtítulo" i] [role="listitem"]',
      '[aria-label*="caption" i] [role="listitem"]',
      '[aria-label*="transcripción" i] [role="listitem"]',
      '[aria-label*="transcript" i] [role="listitem"]',
      '[class*="closed-caption"] [role="listitem"]',
      '[class*="CaptionItem"]',
      '[class*="TranscriptItem"]'
    ];

    const nodes = querySelectorAllDeep(candidateSelectors.join(", "));

    nodes.forEach((node) => {
      if (node.closest("#froggiteams-widget-root")) return;

      const authorEl = node.querySelector(
        '[data-tid="author"], [data-tid*="speaker"], [data-tid*="author"], [class*="author" i], [class*="speaker" i], [class*="DisplayName" i], strong, b, h4'
      );
      const textEl = node.querySelector(
        '[data-tid="closed-caption-text"], [data-tid*="caption-text"], [data-tid*="transcript-text"], [class*="captionText" i], [class*="Text" i], [dir="auto"]'
      );

      let speaker = (authorEl?.textContent || "").trim();
      let text = "";

      if (textEl && textEl !== authorEl) {
        text = (textEl.textContent || "").trim();
      } else {
        // Fallback: split lines inside the listitem
        const rawLines = (node.innerText || "")
          .split(/\r?\n/)
          .map(s => s.trim())
          .filter(Boolean);
        if (rawLines.length >= 2) {
          speaker = speaker || rawLines[0];
          text = rawLines.slice(1).join(" ");
        } else if (rawLines.length === 1) {
          text = rawLines[0];
        }
      }

      if (!speaker) {
        speaker = "Participante Teams";
      }

      // Remove speaker prefix if duplicated inside text
      if (text.startsWith(speaker)) {
        text = text.slice(speaker.length).replace(/^[:\s-]+/, "").trim();
      }

      if (text && text.length > 2) {
        pushCaptionLine(speaker, text, true);
      }
    });

    // Also check aria-live polite/assertive regions used by Teams v2 captions
    const liveRegions = document.querySelectorAll(
      '[data-tid*="caption"][aria-live], [aria-live="polite"][class*="caption" i]'
    );
    liveRegions.forEach((reg) => {
      if (reg.closest("#froggiteams-widget-root")) return;
      const raw = (reg.innerText || "").trim();
      if (!raw || raw.length < 3) return;
      const parts = raw.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      if (parts.length >= 2) {
        pushCaptionLine(parts[0], parts.slice(1).join(" "), true);
      } else {
        pushCaptionLine("Participante Teams", parts[0], true);
      }
    });
  }

  function querySelectorAllDeep(selector, root = document) {
    let results = [];
    try {
      results.push(...Array.from(root.querySelectorAll(selector)));
    } catch {}
    const allElements = root.querySelectorAll("*");
    for (const el of allElements) {
      if (el.shadowRoot) {
        results.push(...querySelectorAllDeep(selector, el.shadowRoot));
      }
    }
    return results;
  }

  /**
   * Attempts to programmatically click the "Turn on live captions" / "Activar subtítulos en vivo"
   * button in Microsoft Teams (teams.cloud.microsoft) or guides the user.
   */
  function tryEnableTeamsLiveCaptions() {
    const banner = document.getElementById("frogg-mode-banner");
    const directBtns = Array.from(
      document.querySelectorAll('button, [role="menuitem"], [role="menuitemcheckbox"]')
    );
    const captionBtn = directBtns.find((b) => {
      const label = (
        (b.getAttribute("aria-label") || "") +
        " " +
        (b.getAttribute("data-tid") || "") +
        " " +
        (b.textContent || "")
      ).toLowerCase();
      return (
        label.includes("subtítulo") ||
        label.includes("live captions") ||
        label.includes("closed-captions") ||
        label.includes("turn on live captions")
      );
    });

    if (captionBtn) {
      captionBtn.click();
      if (banner) {
        banner.innerHTML = `<span style="color:#34d399;">✅ Subtítulos en vivo de Teams activados. Habla en la reunión para ver el texto.</span>`;
      }
      return;
    }

    // Also start Microphone WebSpeech immediately so the user gets instant transcription even in solo test calls!
    if (!useWebSpeechFallback) {
      toggleWebSpeechRecognition();
    }
    if (banner) {
      banner.innerHTML = `<span style="color:#fde68a;">🎙️ Activamos tu micrófono directo + escucha de Teams. En Teams también puedes ir a <strong>Más (...) → Idioma y voz → Activar subtítulos en vivo</strong>.</span>`;
    }
  }

  /**
   * Web Speech API real-time microphone recognition (works immediately in solo test meetings on teams.cloud.microsoft)
   */
  function toggleWebSpeechRecognition() {
    const micBtn = document.getElementById("frogg-btn-mic");
    const banner = document.getElementById("frogg-mode-banner");
    const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;

    if (!SpeechRec) {
      if (banner) {
        banner.innerHTML = `<span style="color:#fca5a5;">⚠️ WebSpeech no disponible. Activa Subtítulos en vivo (Más ... → Idioma y voz) en Teams.</span>`;
      }
      return;
    }

    if (useWebSpeechFallback && speechRecognition) {
      useWebSpeechFallback = false;
      try {
        speechRecognition.stop();
      } catch {}
      micBtn?.classList.remove("active");
      if (banner) {
        banner.innerHTML = `<span>⏸ Micrófono directo desactivado. Escuchando subtítulos DOM de Teams.</span>`;
      }
      return;
    }

    useWebSpeechFallback = true;
    micBtn?.classList.add("active");
    if (banner) {
      banner.innerHTML = `<span style="color:#34d399;">🎙️ <strong>Transcribiendo micrófono en tiempo real</strong> + Subtítulos de Teams activos. ¡Habla ahora!</span>`;
    }

    speechRecognition = new SpeechRec();
    speechRecognition.lang = "es-CL";
    speechRecognition.continuous = true;
    speechRecognition.interimResults = true;

    speechRecognition.onresult = (event) => {
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const res = event.results[i];
        const phrase = res[0]?.transcript?.trim();
        if (!phrase) continue;
        const speaker =
          document.getElementById("frogg-manual-speaker")?.value || "Tú (Micrófono Teams)";
        if (res.isFinal) {
          pushCaptionLine(speaker, phrase, false);
        }
      }
    };

    speechRecognition.onerror = (event) => {
      console.warn("[FroggIteAms WebSpeech]", event.error);
      if (event.error === "not-allowed" && banner) {
        banner.innerHTML = `<span style="color:#fca5a5;">⚠️ Permiso de micrófono denegado en la pestaña. Permite el micrófono en la barra URL de Chrome o activa Subtítulos en vivo en Teams.</span>`;
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
      console.warn("[FroggIteAms] WebSpeech start error:", e);
    }
  }

  function detectTeamsMeetingTitle() {
    const raw = document.title || "";
    const cleaned = raw
      .replace(/\|?\s*Microsoft Teams/i, "")
      .replace(/^\(\d+\)\s*/, "")
      .trim();
    return cleaned || "Reunión Microsoft Teams";
  }

  function makeDraggable(element, handle) {
    if (!handle) return;
    let pos1 = 0,
      pos2 = 0,
      pos3 = 0,
      pos4 = 0;
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
      element.style.top = element.offsetTop - pos2 + "px";
      element.style.left = element.offsetLeft - pos1 + "px";
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
