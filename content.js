// content.js
// FroggIteAms In-Page Live Transcriber Widget for Microsoft Teams
// Supports:
// - https://teams.cloud.microsoft/* (New MS Teams v2 Unified Cloud App - Fluent UI v9)
// - https://teams.microsoft.com/*
// - https://teams.live.com/*
//
// Key Capabilities:
// 1. ZERO DATA LOSS: Continuously syncs every spoken line to permanent history + triggers FINISH_MEETING when leaving call or closing window.
// 2. RELIABLE AUTO-CC ACTIVATION: Retries enabling MS Teams Live Captions with true PointerEvent/MouseEvent sequence and sets Spoken Language to Español (México) (es-MX).
// 3. STRICT CALL DETECTION & CHAT ISOLATION: Only activates inside an active call and strictly ignores group chats/channel messages.

(function initFroggIteAmsContentScript() {
  const isTopFrame = window === window.top;
  if (window.__FROGGITEAMS_INITIALIZED_V4__) return;
  window.__FROGGITEAMS_INITIALIZED_V4__ = true;

  const SUPPORTED_CC_LANGUAGES = [
    {
      code: "es-MX",
      label: "🇲🇽 Español (México)",
      teamsKeywords: ["español (méxico)", "spanish (mexico)", "es-mx"]
    },
    {
      code: "es-ES",
      label: "🇪🇸 Español (España)",
      teamsKeywords: ["español (españa)", "spanish (spain)", "es-es"]
    },
    {
      code: "es-CL",
      label: "🇨🇱 Español (Chile / Latam)",
      teamsKeywords: ["español (méxico)", "spanish (mexico)", "español"]
    },
    {
      code: "en-US",
      label: "🇺🇸 English (US)",
      teamsKeywords: ["english (united states)", "inglés (estados unidos)", "en-us"]
    },
    {
      code: "pt-BR",
      label: "🇧🇷 Português (Brasil)",
      teamsKeywords: ["português (brasil)", "portuguese (brazil)", "pt-br"]
    }
  ];

  let selectedCcLang = "es-MX"; // Default: Español (México)
  let isCapturing = true;
  let isMinimized = false;
  let inActiveCall = false;
  let autoCcAttempts = 0;
  let lastAutoCcAttemptMs = 0;
  let autoCcEnablingInProgress = false;
  let useWebSpeechFallback = false;
  let speechRecognition = null;
  let meetingStartedMs = Date.now();
  let currentEntries = [];
  let activeStageSpeaker = "";
  let captionNodeState = new WeakMap();
  let lastSpeakerTextMap = new Map();

  const iconUrl = chrome.runtime.getURL("assets/icon-48.png");

  chrome.storage.local.get(["froggiteams_cc_lang"], (res) => {
    if (res?.froggiteams_cc_lang) {
      selectedCcLang = res.froggiteams_cc_lang;
      const sel = document.getElementById("frogg-cc-lang-select");
      if (sel) sel.value = selectedCcLang;
    }
  });

  if (isTopFrame) {
    mountFloatingWidget();
    updateWidgetVisibility(false);
  }

  // Ensure if the user closes the tab or window during a meeting, we finalize the meeting in storage
  window.addEventListener("beforeunload", () => {
    if (inActiveCall && currentEntries.length > 0) {
      try {
        chrome.runtime.sendMessage({ type: "FINISH_MEETING" });
      } catch {}
    }
  });

  const observer = new MutationObserver((mutations) => {
    if (!inActiveCall || !isCapturing) return;
    handleStrictCaptionMutations(mutations);
  });

  if (document.body) {
    observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: true
    });
  }

  // Main lifecycle & auto-CC loop (every 650ms)
  setInterval(() => {
    const nowInCall = checkIsInActiveTeamsMeeting();

    if (nowInCall && !inActiveCall) {
      // Entered an active meeting!
      inActiveCall = true;
      autoCcAttempts = 0;
      lastAutoCcAttemptMs = 0;
      meetingStartedMs = Date.now();
      lastSpeakerTextMap.clear();

      if (isTopFrame) {
        updateWidgetVisibility(true);
        const titleInput = document.getElementById("frogg-meeting-title");
        if (titleInput) titleInput.value = detectTeamsMeetingTitle();
        chrome.runtime.sendMessage({
          type: "START_MEETING",
          title: detectTeamsMeetingTitle(),
          sourceUrl: window.location.href
        });
      }
    } else if (!nowInCall && inActiveCall) {
      // Meeting just closed / ended!
      inActiveCall = false;
      autoCcAttempts = 0;
      if (isTopFrame) {
        updateWidgetVisibility(false);
        // Finalize meeting in background so AI report is generated and it is marked completed in History
        chrome.runtime.sendMessage({ type: "FINISH_MEETING" });
      }
    }

    if (!inActiveCall || !isCapturing) return;

    const someoneSpeaking = detectActiveSpeakerOnStage();
    const captionsVisible = areTeamsCaptionsCurrentlyOpen();

    // Auto-enable Live Captions (CC) if not open yet:
    // Retries up to 6 times (every 4.5 seconds) after joining call OR immediately when someone speaks
    const now = Date.now();
    if (
      !captionsVisible &&
      !autoCcEnablingInProgress &&
      autoCcAttempts < 6 &&
      (now - lastAutoCcAttemptMs > 4500 || (someoneSpeaking && autoCcAttempts < 3))
    ) {
      autoCcAttempts++;
      lastAutoCcAttemptMs = now;
      autoEnableTeamsLiveCaptionsSequence(selectedCcLang, false);
    }

    scanStrictCaptionContainers();
  }, 650);

  /**
   * Dispatches a complete sequence of Pointer & Mouse events so Fluent UI v9 / React 18
   * in https://teams.cloud.microsoft/ reliably opens menus and toggles buttons.
   */
  function simulateFluentClick(el) {
    if (!el) return false;
    try {
      el.focus?.();
      const opts = { bubbles: true, cancelable: true, view: window, buttons: 1 };
      el.dispatchEvent(new PointerEvent("pointerdown", opts));
      el.dispatchEvent(new MouseEvent("mousedown", opts));
      el.dispatchEvent(new PointerEvent("pointerup", opts));
      el.dispatchEvent(new MouseEvent("mouseup", opts));
      el.click();
      return true;
    } catch {
      try {
        el.click();
        return true;
      } catch {
        return false;
      }
    }
  }

  /**
   * Strictly checks if the current Teams view is an ACTIVE meeting/call
   */
  function checkIsInActiveTeamsMeeting() {
    const callIndicators = [
      '[data-tid="hangup-button"]',
      '[data-tid="call-hangup"]',
      '[data-tid="calling-unified-bar"]',
      '[data-tid="call-stage"]',
      '[data-tid="modern-calling-screen"]',
      '[data-tid="closed-captions-renderer"]',
      '[data-tid="closed-caption-v2-window-wrapper"]',
      'button[id*="hangup" i]',
      'button[data-tid*="hangup" i]',
      'button[aria-label*="Salir" i][id*="hangup" i]',
      'button[aria-label*="Leave" i][id*="hangup" i]'
    ];

    for (const sel of callIndicators) {
      const el = querySelectorDeepFirst(sel);
      if (el && isElementVisible(el)) {
        return true;
      }
    }
    return false;
  }

  function areTeamsCaptionsCurrentlyOpen() {
    const captionWrapperSelectors = [
      '[data-tid="closed-captions-renderer"]',
      '[data-tid="closed-caption-v2-window-wrapper"]',
      '[data-tid="closed-caption-item"]',
      '[data-tid="call-transcript-panel"]',
      '[data-tid*="closed-caption-v2"]'
    ];
    for (const sel of captionWrapperSelectors) {
      const el = querySelectorDeepFirst(sel);
      if (el && isElementVisible(el)) return true;
    }
    return false;
  }

  function isElementVisible(el) {
    if (!el) return false;
    const rect = el.getBoundingClientRect?.();
    return Boolean(rect && rect.width > 0 && rect.height > 0);
  }

  function updateWidgetVisibility(visible) {
    const root = document.getElementById("froggiteams-widget-root");
    if (!root) return;
    root.style.display = visible ? "block" : "none";
  }

  function mountFloatingWidget() {
    const existing = document.getElementById("froggiteams-widget-root");
    if (existing) existing.remove();

    const langOptionsHtml = SUPPORTED_CC_LANGUAGES.map(
      (l) =>
        `<option value="${l.code}" ${
          l.code === selectedCcLang ? "selected" : ""
        }>${l.label}</option>`
    ).join("");

    const root = document.createElement("div");
    root.id = "froggiteams-widget-root";
    root.style.display = "none";
    root.innerHTML = `
      <div class="frogg-panel" id="frogg-panel">
        <div class="frogg-header" id="frogg-drag-handle">
          <div class="frogg-brand">
            <img src="${iconUrl}" alt="FroggIteAms" class="frogg-logo" />
            <div class="frogg-title-wrap">
              <span class="frogg-title">FroggIteAms</span>
              <span class="frogg-status" id="frogg-status-badge">
                <span class="frogg-dot live"></span> Auto-Guardado Activo
              </span>
            </div>
          </div>
          <div class="frogg-header-actions">
            <button class="frogg-icon-btn" id="frogg-btn-captions" title="Forzar activación de Subtítulos en Vivo (CC) en Teams">CC</button>
            <button class="frogg-icon-btn" id="frogg-btn-mic" title="Respaldo Micrófono Local (WebSpeech)">🎙️</button>
            <button class="frogg-icon-btn" id="frogg-btn-dashboard" title="Abrir Menú Histórico y Reportes (Tactiq Hub)">📊</button>
            <button class="frogg-icon-btn" id="frogg-btn-minimize" title="Minimizar/Expandir">_</button>
          </div>
        </div>

        <div class="frogg-body" id="frogg-body">
          <div class="frogg-meeting-bar">
            <input type="text" id="frogg-meeting-title" class="frogg-meeting-input" value="${detectTeamsMeetingTitle()}" placeholder="Título de la reunión..." />
            <span class="frogg-timer" id="frogg-timer">00:00</span>
          </div>

          <div class="frogg-lang-bar">
            <label for="frogg-cc-lang-select" class="frogg-lang-label">🗣️ Idioma CC:</label>
            <select id="frogg-cc-lang-select" class="frogg-lang-select">
              ${langOptionsHtml}
            </select>
            <button id="frogg-btn-apply-lang" class="frogg-btn-apply-lang" title="Aplicar idioma de subtítulos en MS Teams">Aplicar</button>
          </div>

          <div class="frogg-mode-banner" id="frogg-mode-banner">
            <span>💾 <strong>Auto-guardado en tiempo real:</strong> No perderás datos al cerrar la reunión. Activando CC en <strong>Español (México)</strong>...</span>
          </div>

          <div class="frogg-stream" id="frogg-stream">
            <div class="frogg-empty" id="frogg-empty">
              <p>🐸 <strong>Auto-activando Subtítulos (CC)...</strong></p>
              <p class="frogg-hint">Toda frase se guarda al instante en el historial de reuniones aunque cierres la llamada.</p>
            </div>
          </div>

          <div class="frogg-quick-input">
            <input type="text" id="frogg-manual-speaker" class="frogg-speaker-input" placeholder="Hablante" value="Tú" />
            <input type="text" id="frogg-manual-text" class="frogg-text-input" placeholder="Nota rápida en la reunión..." />
            <button id="frogg-btn-add-line" class="frogg-btn-small" title="Agregar a transcripción">+</button>
          </div>

          <div class="frogg-footer">
            <button class="frogg-btn frogg-btn-secondary" id="frogg-btn-pause">⏸ Pausar</button>
            <button class="frogg-btn frogg-btn-accent" id="frogg-btn-highlight-last">⭐ Highlight</button>
            <button class="frogg-btn frogg-btn-primary" id="frogg-btn-finish">📊 Ver Reporte</button>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(root);

    chrome.storage.onChanged.addListener((changes, area) => {
      if (area === "local" && changes.froggiteams_active_meeting?.newValue) {
        const updated = changes.froggiteams_active_meeting.newValue;
        currentEntries = updated.transcript || [];
        renderStream();
      }
    });

    setInterval(() => {
      if (!inActiveCall || !isCapturing) return;
      const elapsed = Math.floor((Date.now() - meetingStartedMs) / 1000);
      const mm = String(Math.floor(elapsed / 60)).padStart(2, "0");
      const ss = String(elapsed % 60).padStart(2, "0");
      const timerEl = document.getElementById("frogg-timer");
      if (timerEl) timerEl.textContent = `${mm}:${ss}`;
    }, 1000);

    document.getElementById("frogg-cc-lang-select").addEventListener("change", (e) => {
      selectedCcLang = e.target.value;
      chrome.storage.local.set({ froggiteams_cc_lang: selectedCcLang });
      if (speechRecognition && useWebSpeechFallback) {
        speechRecognition.lang = selectedCcLang;
      }
    });

    document.getElementById("frogg-btn-apply-lang").addEventListener("click", () => {
      const sel = document.getElementById("frogg-cc-lang-select");
      selectedCcLang = sel?.value || "es-MX";
      chrome.storage.local.set({ froggiteams_cc_lang: selectedCcLang });
      autoEnableTeamsLiveCaptionsSequence(selectedCcLang, true);
    });

    document.getElementById("frogg-btn-minimize").addEventListener("click", () => {
      isMinimized = !isMinimized;
      const body = document.getElementById("frogg-body");
      body.style.display = isMinimized ? "none" : "flex";
    });

    document.getElementById("frogg-btn-dashboard").addEventListener("click", () => {
      window.open(chrome.runtime.getURL("dashboard.html"), "_blank");
    });

    document.getElementById("frogg-btn-captions").addEventListener("click", () => {
      autoEnableTeamsLiveCaptionsSequence(selectedCcLang, true);
    });

    document.getElementById("frogg-btn-pause").addEventListener("click", (e) => {
      isCapturing = !isCapturing;
      const statusBadge = document.getElementById("frogg-status-badge");
      if (isCapturing) {
        e.target.textContent = "⏸ Pausar";
        statusBadge.innerHTML = `<span class="frogg-dot live"></span> Auto-Guardado Activo`;
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
      window.open(chrome.runtime.getURL("dashboard.html"), "_blank");
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

  function isInsideChatOrChannelPane(node) {
    let curr = node;
    while (curr && curr !== document.body) {
      const tid = (curr.getAttribute?.("data-tid") || "").toLowerCase();
      const aria = (curr.getAttribute?.("aria-label") || "").toLowerCase();
      const id = (curr.id || "").toLowerCase();

      if (
        tid.includes("chat-pane") ||
        tid.includes("message-pane") ||
        tid.includes("channel-pane") ||
        tid.includes("chat-list") ||
        tid.includes("chat-item") ||
        tid.includes("reply-chain") ||
        id.includes("chat-pane") ||
        id.includes("message-pane") ||
        aria.includes("chat de la reunión") ||
        aria.includes("meeting chat") ||
        aria.includes("lista de chats") ||
        aria.includes("chat list")
      ) {
        return true;
      }
      curr = curr.parentElement || curr.getRootNode?.()?.host;
    }
    return false;
  }

  function isStrictlyInsideLiveCaptionsContainer(node) {
    if (!node || isInsideChatOrChannelPane(node)) return false;
    let curr = node;
    while (curr && curr !== document.body) {
      const tid = (curr.getAttribute?.("data-tid") || "").toLowerCase();
      const aria = (curr.getAttribute?.("aria-label") || "").toLowerCase();
      const cls = (typeof curr.className === "string" ? curr.className : "").toLowerCase();

      if (
        tid === "closed-captions-renderer" ||
        tid === "closed-caption-v2-window-wrapper" ||
        tid === "closed-caption-item" ||
        tid === "call-transcript-panel" ||
        tid.includes("closed-caption") ||
        tid.includes("live-caption") ||
        cls.includes("closed-caption") ||
        cls.includes("livecaption") ||
        aria.includes("subtítulos en vivo") ||
        aria.includes("live captions")
      ) {
        return true;
      }
      curr = curr.parentElement || curr.getRootNode?.()?.host;
    }
    return false;
  }

  function handleStrictCaptionMutations(mutations) {
    for (const m of mutations) {
      const targetEl =
        m.target.nodeType === Node.TEXT_NODE ? m.target.parentElement : m.target;
      if (!targetEl || targetEl.closest?.("#froggiteams-widget-root")) continue;

      if (!isStrictlyInsideLiveCaptionsContainer(targetEl)) continue;

      const captionRow = targetEl.closest?.(
        '[data-tid="closed-caption-item"], [data-tid*="caption-item"], [data-tid*="transcript-item"], [role="listitem"], [class*="fui-ChatMessage"]'
      );
      if (captionRow && !isInsideChatOrChannelPane(captionRow)) {
        extractFromCaptionRow(captionRow);
      }
    }
  }

  function scanStrictCaptionContainers() {
    const strictContainerSelectors = [
      '[data-tid="closed-captions-renderer"]',
      '[data-tid="closed-caption-v2-window-wrapper"]',
      '[data-tid="call-transcript-panel"]',
      '[data-tid*="closed-caption"]',
      '[aria-label*="Subtítulos en vivo" i]',
      '[aria-label*="Live captions" i]'
    ];

    const containers = querySelectorAllDeep(strictContainerSelectors.join(", "));
    const processedRows = new Set();

    containers.forEach((container) => {
      if (
        container.closest?.("#froggiteams-widget-root") ||
        isInsideChatOrChannelPane(container)
      ) {
        return;
      }

      const rows = container.querySelectorAll(
        '[data-tid="closed-caption-item"], [data-tid*="caption-item"], [data-tid*="transcript-item"], [role="listitem"], [class*="fui-ChatMessage"]'
      );

      rows.forEach((row) => {
        if (processedRows.has(row) || isInsideChatOrChannelPane(row)) return;
        processedRows.add(row);
        extractFromCaptionRow(row);
      });
    });
  }

  function extractFromCaptionRow(row) {
    if (!row || row.closest?.("#froggiteams-widget-root") || isInsideChatOrChannelPane(row)) {
      return;
    }

    const authorEl = row.querySelector(
      '[data-tid="author"], [data-tid*="author"], [data-tid*="speaker"], [class*="fui-ChatMessage__author"], [class*="author" i], [class*="speaker" i], [class*="DisplayName" i]'
    );
    const textEl = row.querySelector(
      '[data-tid="closed-caption-text"], [data-tid*="caption-text"], [data-tid*="transcript-text"], [class*="fui-ChatMessage__body"], [class*="captionText" i]'
    );

    let speaker = "";
    let text = "";

    if (authorEl && textEl && authorEl !== textEl) {
      speaker = (authorEl.innerText || authorEl.textContent || "").trim();
      text = (textEl.innerText || textEl.textContent || "").trim();
    } else {
      const rawLines = (row.innerText || "")
        .split(/\r?\n/)
        .map((s) => s.trim())
        .filter((s) => s.length > 0 && !/^\d{1,2}:\d{2}(\s*[ap]\.?\s*m\.?)?$/i.test(s));

      if (rawLines.length >= 2) {
        speaker = rawLines[0];
        text = rawLines.slice(1).join(" ");
      } else if (rawLines.length === 1) {
        const colonMatch = rawLines[0].match(/^([^:]{2,45}):\s+(.+)$/);
        if (colonMatch) {
          speaker = colonMatch[1].trim();
          text = colonMatch[2].trim();
        }
      }
    }

    speaker = speaker
      .replace(/\b\d{1,2}:\d{2}(\s*[ap]\.?\s*m\.?)?\b/gi, "")
      .replace(/\s+/g, " ")
      .trim();

    if (speaker && text.startsWith(speaker)) {
      text = text.slice(speaker.length).replace(/^[:\s-]+/, "").trim();
    }

    if (!text || text.length < 2) return;

    const prevRowText = captionNodeState.get(row);
    if (prevRowText === `${speaker}|${text}`) return;
    captionNodeState.set(row, `${speaker}|${text}`);

    pushCaptionLine(speaker || activeStageSpeaker || "Asistente Teams", text, true);
  }

  function detectActiveSpeakerOnStage() {
    const activeSelectors = [
      '[data-tid*="active-speaker"]',
      '[class*="activeSpeaker" i]',
      '[class*="speaking" i][data-tid*="participant"]',
      '[aria-label*="hablando" i]',
      '[aria-label*="speaking" i]'
    ];
    const el = querySelectorDeepFirst(activeSelectors.join(", "));
    if (el && !el.closest?.("#froggiteams-widget-root")) {
      const nameNode =
        el.querySelector('[data-tid*="displayname" i], [class*="displayName" i]') || el;
      const rawName = (nameNode.getAttribute("aria-label") || nameNode.textContent || "")
        .replace(/\b(está hablando|is speaking|hablando|speaking|con audio|muted|silenciado)\b/gi, "")
        .replace(/[(),]/g, "")
        .trim();
      if (rawName && rawName.length >= 2 && rawName.length <= 60) {
        activeStageSpeaker = rawName;
      }
      return true;
    }
    return false;
  }

  /**
   * Robust multi-step navigation for MS Teams v2 (https://teams.cloud.microsoft/)
   * Uses `simulateFluentClick` (PointerDown -> MouseDown -> PointerUp -> MouseUp -> Click)
   * so Fluent UI v9 menus open reliably, enables Live Captions, and sets Spoken Language to Español (México).
   */
  async function autoEnableTeamsLiveCaptionsSequence(
    langCode = "es-MX",
    forceLanguageChange = false
  ) {
    if (autoCcEnablingInProgress) return;
    autoCcEnablingInProgress = true;

    const banner = document.getElementById("frogg-mode-banner");
    const langObj =
      SUPPORTED_CC_LANGUAGES.find((l) => l.code === langCode) || SUPPORTED_CC_LANGUAGES[0];
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

    try {
      const findElementByAttrOrText = (keywords) => {
        const candidates = querySelectorAllDeep(
          'button, [role="menuitem"], [role="menuitemcheckbox"], [role="option"], [role="radio"], [data-tid]'
        );
        return candidates.find((el) => {
          if (el.closest?.("#froggiteams-widget-root") || !isElementVisible(el)) return false;
          const haystack = [
            el.id || "",
            el.getAttribute("data-tid") || "",
            el.getAttribute("aria-label") || "",
            el.getAttribute("title") || "",
            el.textContent || ""
          ]
            .join(" ")
            .toLowerCase();
          return keywords.some((kw) => haystack.includes(kw));
        });
      };

      // Step 1: If captions are not currently open, turn them on
      if (!areTeamsCaptionsCurrentlyOpen()) {
        // Check if a direct "LiveCaptions" button is already on the toolbar
        let directCaptionBtn = findElementByAttrOrText([
          "livecaptionsbutton",
          "closed-captions-button",
          "turn on live captions",
          "activar subtítulos en vivo",
          "mostrar subtítulos en vivo"
        ]);

        if (directCaptionBtn) {
          simulateFluentClick(directCaptionBtn);
          await sleep(700);
        } else {
          // Find the "More" (...) button in the Teams calling bar
          const moreBtn = findElementByAttrOrText([
            "callingbuttons-showmorebtn",
            "calling-unified-bar-more",
            "more-button",
            "más acciones",
            "more actions",
            "más opciones"
          ]);

          if (moreBtn) {
            simulateFluentClick(moreBtn);
            await sleep(500);

            // In Teams v2 (teams.cloud.microsoft), Live Captions is nested under "Idioma y voz" / "Language and speech"
            const langSpeechMenu = findElementByAttrOrText([
              "idioma y voz",
              "language and speech",
              "speech-and-language",
              "language-and-speech"
            ]);

            if (langSpeechMenu) {
              simulateFluentClick(langSpeechMenu);
              await sleep(450);
            }

            // Now find "Turn on live captions" / "Mostrar subtítulos en vivo" / "Activar subtítulos en vivo"
            const captionMenuItem = findElementByAttrOrText([
              "mostrar subtítulos en vivo",
              "activar subtítulos en vivo",
              "subtítulos en vivo",
              "turn on live captions",
              "show live captions",
              "live captions",
              "livecaptionsbutton",
              "closed-captions"
            ]);

            if (captionMenuItem) {
              simulateFluentClick(captionMenuItem);
              await sleep(800);
            }
          }
        }
      }

      // Step 2: Handle the "Spoken language" modal (if Teams pops it up on first caption start)
      // or explicitly change language if the user clicked "Aplicar"
      const spokenLangDialogOrSettings = async () => {
        if (forceLanguageChange && areTeamsCaptionsCurrentlyOpen()) {
          const settingsBtn = findElementByAttrOrText([
            "configuración de subtítulos",
            "caption settings",
            "cambiar idioma hablado",
            "change spoken language",
            "closed-caption-settings"
          ]);
          if (settingsBtn) {
            simulateFluentClick(settingsBtn);
            await sleep(450);
            const spokenLangItem = findElementByAttrOrText([
              "idioma hablado",
              "spoken language"
            ]);
            if (spokenLangItem) {
              simulateFluentClick(spokenLangItem);
              await sleep(450);
            }
          }
        }

        // Check if a language combobox/dropdown is open or visible inside a dialog
        const combobox = querySelectorDeepFirst(
          '[role="dialog"] [role="combobox"], [role="dialog"] button[aria-haspopup="listbox"]'
        );
        if (combobox && isElementVisible(combobox)) {
          simulateFluentClick(combobox);
          await sleep(400);
        }

        const langOption = findElementByAttrOrText(langObj.teamsKeywords);
        if (langOption) {
          simulateFluentClick(langOption);
          await sleep(350);
        }

        const confirmBtn = findElementByAttrOrText([
          "confirmar",
          "actualizar",
          "guardar",
          "confirm",
          "update",
          "save"
        ]);
        if (confirmBtn) {
          simulateFluentClick(confirmBtn);
        }
      };

      await spokenLangDialogOrSettings();

      if (areTeamsCaptionsCurrentlyOpen()) {
        if (banner) {
          banner.innerHTML = `<span style="color:#34d399;">✅ Subtítulos activos en <strong>${langObj.label}</strong> y auto-guardando cada frase en tiempo real.</span>`;
        }
      } else if (banner) {
        banner.innerHTML = `<span style="color:#fde68a;">🔄 Intentando activar CC (${autoCcAttempts}/6)... También puedes pulsar <strong>CC</strong> arriba o <strong>Más (...) → Idioma y voz → Mostrar subtítulos en vivo</strong>.</span>`;
      }
    } catch (e) {
      console.warn("[FroggIteAms Auto-CC]", e);
    } finally {
      autoCcEnablingInProgress = false;
    }
  }

  function pushCaptionLine(speaker, text, isUpdate = false) {
    if (!inActiveCall || !isCapturing || !text) return;
    let cleanSpeaker = (speaker || "").replace(/\s+/g, " ").trim();
    let cleanText = text.replace(/\s+/g, " ").trim();
    if (cleanText.length < 2) return;

    if (isTeamsSystemLabel(cleanSpeaker) || isTeamsSystemLabel(cleanText)) {
      return;
    }

    if (!cleanSpeaker || cleanSpeaker === "Participante Teams") {
      cleanSpeaker = activeStageSpeaker || "Asistente Teams";
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

  function toggleWebSpeechRecognition() {
    const micBtn = document.getElementById("frogg-btn-mic");
    const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRec) return;

    if (useWebSpeechFallback && speechRecognition) {
      useWebSpeechFallback = false;
      try {
        speechRecognition.stop();
      } catch {}
      micBtn?.classList.remove("active");
      return;
    }

    useWebSpeechFallback = true;
    micBtn?.classList.add("active");
    speechRecognition = new SpeechRec();
    speechRecognition.lang = selectedCcLang || "es-MX";
    speechRecognition.continuous = true;
    speechRecognition.interimResults = true;

    speechRecognition.onresult = (event) => {
      if (!inActiveCall) return;
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
      if (useWebSpeechFallback && inActiveCall) {
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

  function querySelectorDeepFirst(selector, root = document) {
    try {
      const found = root.querySelector(selector);
      if (found) return found;
    } catch {}
    const allElements = root.querySelectorAll("*");
    for (const el of allElements) {
      if (el.shadowRoot) {
        const foundInShadow = querySelectorDeepFirst(selector, el.shadowRoot);
        if (foundInShadow) return foundInShadow;
      }
    }
    return null;
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
      if (
        e.target.tagName === "BUTTON" ||
        e.target.tagName === "INPUT" ||
        e.target.tagName === "SELECT"
      )
        return;
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
