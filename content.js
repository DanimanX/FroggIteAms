// content.js
// FroggIteAms In-Page Live Transcriber Widget for Microsoft Teams
// Supports:
// - https://teams.cloud.microsoft/* (New MS Teams v2 Unified Cloud App - Fluent UI v9)
// - https://teams.microsoft.com/*
// - https://teams.live.com/*
//
// Multi-layer extraction for ALL meeting participants:
// 1. MS Teams v2 Live Captions & Live Transcript DOM Scanner (Fluent UI v9 + Classic + Shadow DOM)
// 2. Active Speaker Tile / Voice Activity Tracker (tracks who is currently speaking on the stage)
// 3. Auto-Enabler for MS Teams Live Captions (clicks "More -> Language and speech -> Turn on live captions")
// 4. Web Speech API hybrid fallback (attributes speech to active remote speaker if someone else is speaking)

(function initFroggIteAmsContentScript() {
  const isTopFrame = window === window.top;
  if (window.__FROGGITEAMS_INITIALIZED_V2__) return;
  window.__FROGGITEAMS_INITIALIZED_V2__ = true;

  let isCapturing = true;
  let isMinimized = false;
  let useWebSpeechFallback = false;
  let speechRecognition = null;
  let meetingStartedMs = Date.now();
  let currentEntries = [];
  let activeStageSpeaker = ""; // Tracks who has the active speaking ring in MS Teams
  let captionNodeState = new WeakMap(); // DOM node -> last seen text
  let lastSpeakerTextMap = new Map(); // speaker -> last committed text
  let teamsCaptionsDetectedCount = 0;

  const iconUrl = chrome.runtime.getURL("assets/icon-48.png");

  if (isTopFrame) {
    mountFloatingWidget();
  }

  // Observe DOM mutations across the document
  const observer = new MutationObserver((mutations) => {
    if (!isCapturing) return;
    handleCaptionsMutations(mutations);
  });

  if (document.body) {
    observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: true
    });
  }

  // Fast polling (500ms) for Fluent UI v9 virtualized caption rows & active speaker indicator
  setInterval(() => {
    if (!isCapturing) return;
    detectActiveSpeakerOnStage();
    scanAllParticipantsCaptions();
  }, 500);

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
            <button class="frogg-icon-btn" id="frogg-btn-captions" title="Activar Subtítulos en Vivo de Teams para capturar a TODOS los asistentes">CC</button>
            <button class="frogg-icon-btn" id="frogg-btn-mic" title="Micrófono Local (WebSpeech)">🎙️</button>
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
            <span>👥 Para capturar a <strong>todos los asistentes</strong>: pulsa <strong>CC</strong> o activa <em>Subtítulos en vivo</em> en Teams (<strong>Más ... → Idioma y voz → Mostrar subtítulos en vivo</strong>).</span>
          </div>

          <div class="frogg-stream" id="frogg-stream">
            <div class="frogg-empty" id="frogg-empty">
              <p>🐸 <strong>Escuchando a todos los asistentes...</strong></p>
              <p class="frogg-hint">Asegúrate de que los <strong>Subtítulos en vivo (Live Captions)</strong> estén visibles en la parte inferior de Teams (pulsa el botón <strong>CC</strong> arriba para activarlos automáticamente).</p>
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

    chrome.storage.onChanged.addListener((changes, area) => {
      if (area === "local" && changes.froggiteams_active_meeting?.newValue) {
        const updated = changes.froggiteams_active_meeting.newValue;
        currentEntries = updated.transcript || [];
        renderStream();
      }
    });

    setInterval(() => {
      if (!isCapturing) return;
      const elapsed = Math.floor((Date.now() - meetingStartedMs) / 1000);
      const mm = String(Math.floor(elapsed / 60)).padStart(2, "0");
      const ss = String(elapsed % 60).padStart(2, "0");
      const timerEl = document.getElementById("frogg-timer");
      if (timerEl) timerEl.textContent = `${mm}:${ss}`;
    }, 1000);

    document.getElementById("frogg-btn-minimize").addEventListener("click", () => {
      isMinimized = !isMinimized;
      const body = document.getElementById("frogg-body");
      body.style.display = isMinimized ? "none" : "flex";
    });

    document.getElementById("frogg-btn-dashboard").addEventListener("click", () => {
      window.open(chrome.runtime.getURL("dashboard.html"), "_blank");
    });

    document.getElementById("frogg-btn-captions").addEventListener("click", () => {
      autoEnableTeamsLiveCaptionsSequence();
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

  /**
   * Sends a captured or updated line to background.js so that streaming updates from
   * any participant in MS Teams replace the in-progress sentence cleanly rather than duplicating.
   */
  function pushCaptionLine(speaker, text, isUpdate = false) {
    if (!isCapturing || !text) return;
    let cleanSpeaker = (speaker || "").replace(/\s+/g, " ").trim();
    let cleanText = text.replace(/\s+/g, " ").trim();
    if (cleanText.length < 2) return;

    // Filter out Teams UI system strings that aren't spoken dialogue
    if (isTeamsSystemLabel(cleanSpeaker) || isTeamsSystemLabel(cleanText)) {
      return;
    }

    if (!cleanSpeaker || cleanSpeaker === "Participante Teams") {
      cleanSpeaker = activeStageSpeaker || "Participante Teams";
    }

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

  function isTeamsSystemLabel(str) {
    const lower = str.toLowerCase();
    const ignoreExact = [
      "froggiteams",
      "subtítulos en vivo",
      "live captions",
      "mostrar subtítulos",
      "ocultar subtítulos",
      "idioma hablado",
      "spoken language",
      "configuración de subtítulos",
      "caption settings",
      "esperando a que otros se unan",
      "waiting for others to join"
    ];
    return ignoreExact.some((w) => lower === w || lower.startsWith(w + " "));
  }

  /**
   * Detects who currently has the active voice border / speaking indicator on the MS Teams call stage
   */
  function detectActiveSpeakerOnStage() {
    const activeSelectors = [
      '[data-tid*="active-speaker"]',
      '[class*="activeSpeaker" i]',
      '[class*="speaking" i][data-tid*="participant"]',
      '[aria-label*="hablando" i]',
      '[aria-label*="speaking" i]'
    ];
    const el = document.querySelector(activeSelectors.join(", "));
    if (el && !el.closest("#froggiteams-widget-root")) {
      const nameNode =
        el.querySelector('[data-tid*="displayname" i], [class*="displayName" i]') || el;
      const rawName = (nameNode.getAttribute("aria-label") || nameNode.textContent || "")
        .replace(/\b(está hablando|is speaking|hablando|speaking|con audio|muted|silenciado)\b/gi, "")
        .replace(/[(),]/g, "")
        .trim();
      if (rawName && rawName.length >= 2 && rawName.length <= 60) {
        activeStageSpeaker = rawName;
      }
    }
  }

  /**
   * Mutation handler that watches live text changes inside the MS Teams caption bar
   */
  function handleCaptionsMutations(mutations) {
    for (const m of mutations) {
      const targetEl =
        m.target.nodeType === Node.TEXT_NODE ? m.target.parentElement : m.target;
      if (!targetEl || targetEl.closest?.("#froggiteams-widget-root")) continue;

      const captionRow = targetEl.closest?.(
        '[data-tid="closed-caption-item"], [data-tid*="caption-item"], [data-tid*="transcript-item"], [class*="closed-caption"] [role="listitem"], [class*="fui-ChatMessage"], [class*="ChatMessage"], [role="listitem"]'
      );
      if (captionRow && isInsideCaptionOrTranscriptWrapper(captionRow)) {
        extractFromCaptionRow(captionRow);
      }
    }
  }

  function isInsideCaptionOrTranscriptWrapper(node) {
    let curr = node;
    while (curr && curr !== document.body) {
      const tid = (curr.getAttribute?.("data-tid") || "").toLowerCase();
      const cls = (typeof curr.className === "string" ? curr.className : "").toLowerCase();
      const aria = (curr.getAttribute?.("aria-label") || "").toLowerCase();
      const role = (curr.getAttribute?.("role") || "").toLowerCase();

      if (
        tid.includes("caption") ||
        tid.includes("transcript") ||
        cls.includes("caption") ||
        cls.includes("transcript") ||
        aria.includes("subtítulo") ||
        aria.includes("caption") ||
        aria.includes("transcripción") ||
        aria.includes("transcript") ||
        (role === "log" && (aria.includes("subt") || aria.includes("caption") || !aria))
      ) {
        return true;
      }
      curr = curr.parentElement || curr.getRootNode?.()?.host;
    }
    return false;
  }

  /**
   * Comprehensive multi-strategy scanner for ALL attendees' captions in:
   * - teams.cloud.microsoft (New Teams Fluent UI v9: uses fui-ChatMessageCompact / fui-ChatMessage__author / fui-ChatMessage__body or listitem inside caption wrapper)
   * - Classic Teams & Live Transcript side panel
   */
  function scanAllParticipantsCaptions() {
    // Strategy 1: Explicit caption/transcript containers in Teams v2 (teams.cloud.microsoft)
    const containerSelectors = [
      '[data-tid="closed-captions-renderer"]',
      '[data-tid="closed-caption-v2-window-wrapper"]',
      '[data-tid*="closed-caption"]',
      '[data-tid*="live-captions"]',
      '[data-tid*="call-transcript"]',
      '[data-tid*="transcript-panel"]',
      '[aria-label*="subtítulo" i]',
      '[aria-label*="caption" i]',
      '[aria-label*="transcripción" i]',
      '[aria-label*="transcript" i]',
      '[class*="closedCaption" i]',
      '[class*="liveCaption" i]',
      '[class*="captionWrapper" i]',
      '[class*="transcript" i][role="region"]',
      '[role="log"]'
    ];

    const containers = querySelectorAllDeep(containerSelectors.join(", "));
    const processedRows = new Set();

    containers.forEach((container) => {
      if (container.closest?.("#froggiteams-widget-root")) return;

      // Find individual speaker rows inside this container
      const rows = container.querySelectorAll(
        '[data-tid="closed-caption-item"], [data-tid*="caption-item"], [data-tid*="transcript-item"], [role="listitem"], [class*="fui-ChatMessage"], [class*="ChatMessage"], [class*="item" i]'
      );

      if (rows.length > 0) {
        rows.forEach((row) => {
          if (processedRows.has(row)) return;
          processedRows.add(row);
          extractFromCaptionRow(row);
        });
      } else {
        // Container itself might hold direct author + text blocks
        extractFromCaptionRow(container);
      }
    });

    // Strategy 2: Direct Fluent UI v9 message items that have caption author/text data-tids anywhere on page
    const directItems = querySelectorAllDeep(
      '[data-tid="closed-caption-item"], [data-tid="closed-caption-text"], [data-tid*="caption-text"], [class*="fui-ChatMessage__body"]'
    );
    directItems.forEach((el) => {
      if (el.closest?.("#froggiteams-widget-root")) return;
      const row =
        el.closest(
          '[data-tid="closed-caption-item"], [role="listitem"], [class*="fui-ChatMessage"]'
        ) || el.parentElement;
      if (row && !processedRows.has(row)) {
        processedRows.add(row);
        extractFromCaptionRow(row);
      }
    });
  }

  function extractFromCaptionRow(row) {
    if (!row || row.closest?.("#froggiteams-widget-root")) return;

    // 1. Try explicit author & text sub-elements (Teams v2 Fluent UI v9 & Classic)
    const authorEl = row.querySelector(
      '[data-tid="author"], [data-tid*="author"], [data-tid*="speaker"], [class*="fui-ChatMessage__author"], [class*="author" i], [class*="speaker" i], [class*="DisplayName" i], [id*="author" i]'
    );
    const textEl = row.querySelector(
      '[data-tid="closed-caption-text"], [data-tid*="caption-text"], [data-tid*="transcript-text"], [class*="fui-ChatMessage__body"], [class*="captionText" i], [class*="messageBody" i]'
    );

    let speaker = "";
    let text = "";

    if (authorEl && textEl && authorEl !== textEl) {
      speaker = (authorEl.innerText || authorEl.textContent || "").trim();
      text = (textEl.innerText || textEl.textContent || "").trim();
    } else {
      // 2. Fallback: parse innerText lines or leaf text nodes inside the row
      const rawLines = (row.innerText || "")
        .split(/\r?\n/)
        .map((s) => s.trim())
        .filter((s) => s.length > 0 && !/^\d{1,2}:\d{2}(\s*[ap]\.?\s*m\.?)?$/i.test(s));

      if (rawLines.length >= 2) {
        speaker = rawLines[0];
        text = rawLines.slice(1).join(" ");
      } else if (rawLines.length === 1) {
        // Check if format is "Nombre Apellido: Hola a todos"
        const colonMatch = rawLines[0].match(/^([^:]{2,45}):\s+(.+)$/);
        if (colonMatch) {
          speaker = colonMatch[1].trim();
          text = colonMatch[2].trim();
        } else if (authorEl) {
          speaker = (authorEl.textContent || "").trim();
          text = rawLines[0].replace(speaker, "").trim();
        }
      }
    }

    // Clean speaker name (remove timestamps or status badges attached to name)
    speaker = speaker
      .replace(/\b\d{1,2}:\d{2}(\s*[ap]\.?\s*m\.?)?\b/gi, "")
      .replace(/\s+/g, " ")
      .trim();

    // Remove speaker prefix if accidentally repeated at start of text
    if (speaker && text.startsWith(speaker)) {
      text = text.slice(speaker.length).replace(/^[:\s-]+/, "").trim();
    }

    if (!text || text.length < 2) return;

    const prevRowText = captionNodeState.get(row);
    if (prevRowText === `${speaker}|${text}`) return;
    captionNodeState.set(row, `${speaker}|${text}`);

    teamsCaptionsDetectedCount++;
    if (teamsCaptionsDetectedCount === 1) {
      const banner = document.getElementById("frogg-mode-banner");
      if (banner) {
        banner.innerHTML = `<span style="color:#34d399;">✅ <strong>Capturando subtítulos multi-asistente de MS Teams en tiempo real</strong>.</span>`;
      }
    }

    pushCaptionLine(speaker || activeStageSpeaker || "Asistente Teams", text, true);
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
   * Automatically navigates the MS Teams v2 menu to turn on Live Captions ("Subtítulos en vivo")
   * so that ALL remote participants' voices are transcribed by Teams into the DOM.
   */
  async function autoEnableTeamsLiveCaptionsSequence() {
    const banner = document.getElementById("frogg-mode-banner");
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

    const findClickableByKeywords = (keywords) => {
      const all = querySelectorAllDeep(
        'button, [role="menuitem"], [role="menuitemcheckbox"], [data-tid*="caption"], [data-tid*="more"]'
      );
      return all.find((el) => {
        if (el.closest?.("#froggiteams-widget-root")) return false;
        const label = (
          (el.getAttribute("aria-label") || "") +
          " " +
          (el.getAttribute("data-tid") || "") +
          " " +
          (el.textContent || "")
        ).toLowerCase();
        return keywords.some((kw) => label.includes(kw));
      });
    };

    // Step 1: Check if "Turn on live captions" / "Subtítulos en vivo" button is already visible
    let captionBtn = findClickableByKeywords([
      "activar subtítulos",
      "mostrar subtítulos en vivo",
      "turn on live captions",
      "live-captions-button",
      "closed-captions-button"
    ]);

    if (captionBtn) {
      captionBtn.click();
      if (banner) {
        banner.innerHTML = `<span style="color:#34d399;">✅ Subtítulos en vivo de Teams activados para todos los asistentes.</span>`;
      }
      return;
    }

    // Step 2: Click "More" / "Más" (...) button in the Teams call toolbar
    const moreBtn = findClickableByKeywords([
      "calling-unified-bar-more",
      "more-button",
      "más acciones",
      "more actions",
      "más opciones"
    ]);

    if (moreBtn) {
      moreBtn.click();
      await sleep(350);

      // Step 3: Check if "Language and speech" / "Idioma y voz" submenu exists in Teams v2
      const langMenu = findClickableByKeywords([
        "idioma y voz",
        "language and speech",
        "speech-and-language"
      ]);
      if (langMenu) {
        langMenu.click();
        await sleep(300);
      }

      // Step 4: Click "Turn on live captions" / "Mostrar subtítulos en vivo"
      captionBtn = findClickableByKeywords([
        "subtítulos en vivo",
        "live captions",
        "closed-captions"
      ]);

      if (captionBtn) {
        captionBtn.click();
        if (banner) {
          banner.innerHTML = `<span style="color:#34d399;">✅ ¡Subtítulos en vivo de Teams activados! Ahora se transcribirá a todos los asistentes.</span>`;
        }
        return;
      }
    }

    if (banner) {
      banner.innerHTML = `<span style="color:#fde68a;">👉 En la barra de Teams haz clic en <strong>Más (...) → Idioma y voz → Mostrar subtítulos en vivo</strong>. Así Teams decodifica el audio de todos los asistentes y FroggIteAms los captura al instante.</span>`;
    }
  }

  /**
   * Optional Web Speech API microphone fallback
   */
  function toggleWebSpeechRecognition() {
    const micBtn = document.getElementById("frogg-btn-mic");
    const banner = document.getElementById("frogg-mode-banner");
    const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;

    if (!SpeechRec) {
      if (banner) {
        banner.innerHTML = `<span style="color:#fca5a5;">⚠️ Activa Subtítulos en vivo (CC) en Teams para capturar a todos.</span>`;
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
        banner.innerHTML = `<span>👥 Modo Subtítulos de Teams activo (captura a todos los asistentes).</span>`;
      }
      return;
    }

    useWebSpeechFallback = true;
    micBtn?.classList.add("active");
    if (banner) {
      banner.innerHTML = `<span style="color:#34d399;">🎙️ Micrófono local + Subtítulos de Teams activos. (Para capturar a los demás con auriculares, pulsa <strong>CC</strong> para activar Subtítulos en vivo de Teams).</span>`;
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
          activeStageSpeaker ||
          document.getElementById("frogg-manual-speaker")?.value ||
          "Tú";
        if (res.isFinal) {
          pushCaptionLine(speaker, phrase, false);
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
    } catch {}
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
