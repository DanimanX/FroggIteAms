// content.js
// FroggIteAms In-Page Live Transcriber Widget for Microsoft Teams
// Supports:
// - https://teams.cloud.microsoft/* (New MS Teams v2 Unified Cloud App - Fluent UI v9)
// - https://teams.microsoft.com/*
// - https://teams.live.com/*
//
// Key Capabilities:
// 1. ZERO DATA LOSS: Continuously syncs every spoken line to permanent history + triggers FINISH_MEETING when leaving call or closing window.
// 2. RELIABLE AUTO-CC ACTIVATION: Finds the call toolbar "Más" / "More" button (near Hangup/Leave), opens "Idioma y voz" / "Language and speech" via hover + pointer + keyboard events, clicks "Mostrar subtítulos en vivo", and selects "Español (México)".
// 3. STRICT CALL DETECTION & CHAT ISOLATION: Only activates inside an active call and strictly ignores group chats/channel messages.

(function initFroggIteAmsContentScript() {
  const isTopFrame = window === window.top;
  if (window.__FROGGITEAMS_INITIALIZED_V5__) return;
  window.__FROGGITEAMS_INITIALIZED_V5__ = true;

  const SUPPORTED_CC_LANGUAGES = [
    {
      code: "es-MX",
      label: "🇲🇽 Español (México)",
      teamsKeywords: ["español (méxico)", "español (mexico)", "spanish (mexico)", "es-mx"]
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
      teamsKeywords: ["english (united states)", "english (u.s.)", "inglés (estados unidos)", "en-us"]
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
      inActiveCall = false;
      autoCcAttempts = 0;
      if (isTopFrame) {
        updateWidgetVisibility(false);
        chrome.runtime.sendMessage({ type: "FINISH_MEETING" });
      }
    }

    if (!inActiveCall || !isCapturing) return;

    const someoneSpeaking = detectActiveSpeakerOnStage();
    const captionsVisible = areTeamsCaptionsCurrentlyOpen();

    const now = Date.now();
    if (
      !captionsVisible &&
      !autoCcEnablingInProgress &&
      autoCcAttempts < 5 &&
      (now - lastAutoCcAttemptMs > 5000 || (someoneSpeaking && autoCcAttempts < 2))
    ) {
      autoCcAttempts++;
      lastAutoCcAttemptMs = now;
      autoEnableTeamsLiveCaptionsSequence(selectedCcLang, false);
    }

    scanStrictCaptionContainers();
  }, 650);

  /**
   * Dispatches real coordinates Pointer, Mouse, Hover, and Keyboard events so Fluent UI v9 / React 18
   * menus and submenus in https://teams.cloud.microsoft/ reliably open.
   */
  function simulateFluentInteraction(el, includeSubmenuHover = false) {
    if (!el) return false;
    try {
      const rect = el.getBoundingClientRect();
      const clientX = rect.left + rect.width / 2;
      const clientY = rect.top + rect.height / 2;
      const evOpts = {
        bubbles: true,
        cancelable: true,
        composed: true,
        view: window,
        clientX,
        clientY,
        buttons: 1
      };

      el.focus?.();
      el.dispatchEvent(new PointerEvent("pointerover", evOpts));
      el.dispatchEvent(new MouseEvent("mouseover", evOpts));
      el.dispatchEvent(new PointerEvent("pointerenter", evOpts));
      el.dispatchEvent(new MouseEvent("mouseenter", evOpts));
      el.dispatchEvent(new PointerEvent("pointerdown", evOpts));
      el.dispatchEvent(new MouseEvent("mousedown", evOpts));
      el.dispatchEvent(new PointerEvent("pointerup", evOpts));
      el.dispatchEvent(new MouseEvent("mouseup", evOpts));
      el.click();

      if (includeSubmenuHover) {
        // Fluent UI v9 nested MenuTrigger also responds to ArrowRight or Enter when focused
        el.dispatchEvent(
          new KeyboardEvent("keydown", {
            key: "ArrowRight",
            code: "ArrowRight",
            keyCode: 39,
            bubbles: true,
            cancelable: true
          })
        );
      }
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

  function findHangupButton() {
    const selectors = [
      '[data-tid="hangup-button"]',
      '[data-tid="call-hangup"]',
      'button[id*="hangup" i]',
      'button[data-tid*="hangup" i]',
      'button[aria-label*="Salir" i]',
      'button[aria-label*="Colgar" i]',
      'button[aria-label*="Leave" i]'
    ];
    for (const sel of selectors) {
      const candidates = querySelectorAllDeep(sel);
      for (const c of candidates) {
        if (c.closest?.("#froggiteams-widget-root")) continue;
        if (isElementVisible(c)) return c;
      }
    }
    return null;
  }

  function checkIsInActiveTeamsMeeting() {
    if (findHangupButton()) return true;

    const otherCallIndicators = [
      '[data-tid="calling-unified-bar"]',
      '[data-tid="call-stage"]',
      '[data-tid="modern-calling-screen"]',
      '[data-tid="closed-captions-renderer"]',
      '[data-tid="closed-caption-v2-window-wrapper"]'
    ];
    for (const sel of otherCallIndicators) {
      const el = querySelectorDeepFirst(sel);
      if (el && isElementVisible(el)) return true;
    }
    return false;
  }

  function areTeamsCaptionsCurrentlyOpen() {
    const captionWrapperSelectors = [
      '[data-tid="closed-captions-renderer"]',
      '[data-tid="closed-caption-v2-window-wrapper"]',
      '[data-tid="closed-caption-item"]',
      '[data-tid="call-transcript-panel"]',
      '[data-tid*="closed-caption-v2"]',
      '[aria-label*="Subtítulos en vivo" i][role="region"]',
      '[aria-label*="Live captions" i][role="region"]'
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
            <button class="frogg-icon-btn" id="frogg-btn-captions" title="Activar Subtítulos en Vivo (CC) en Teams">CC</button>
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
            <button id="frogg-btn-apply-lang" class="frogg-btn-apply-lang" title="Activar CC y aplicar idioma en MS Teams">Aplicar</button>
          </div>

          <div class="frogg-mode-banner" id="frogg-mode-banner">
            <span>💾 <strong>Auto-guardado activo:</strong> Tus datos se guardan frase a frase. Idioma CC: <strong>Español (México)</strong>.</span>
          </div>

          <div class="frogg-stream" id="frogg-stream">
            <div class="frogg-empty" id="frogg-empty">
              <p>🐸 <strong>Reunión detectada en MS Teams</strong></p>
              <p class="frogg-hint">Activando Subtítulos en vivo (CC) automáticamente...</p>
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
      autoEnableTeamsLiveCaptionsSequence(selectedCcLang, true, true);
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
      autoEnableTeamsLiveCaptionsSequence(selectedCcLang, true, true);
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

  /**
   * Finds the exact "More (...)" button inside the active MS Teams call control bar
   * by locating the toolbar that contains the Hangup/Leave button (avoiding sidebar/app-bar "..." buttons).
   */
  function findCallToolbarMoreButton() {
    const hangup = findHangupButton();

    // 1. Search inside the same toolbar container as the Hangup button first
    if (hangup) {
      let toolbar =
        hangup.closest('[role="toolbar"], [data-tid*="calling" i], [data-tid*="unified-bar" i]') ||
        hangup.parentElement?.parentElement?.parentElement;

      if (toolbar) {
        const toolbarButtons = Array.from(toolbar.querySelectorAll("button, [role='button']"));
        const moreInToolbar = toolbarButtons.find((btn) => {
          if (btn === hangup || !isElementVisible(btn)) return false;
          const desc = [
            btn.id || "",
            btn.getAttribute("data-tid") || "",
            btn.getAttribute("aria-label") || "",
            btn.getAttribute("title") || "",
            btn.textContent || ""
          ]
            .join(" ")
            .toLowerCase();
          return (
            desc.includes("more") ||
            desc.includes("más") ||
            desc.includes("mas ") ||
            desc.includes("showmore") ||
            desc.includes("button-more")
          );
        });
        if (moreInToolbar) return moreInToolbar;
      }
    }

    // 2. Fallback: search globally for explicit calling More button IDs/data-tids
    const explicitSelectors = [
      '#callingButtons-showMoreBtn',
      '[data-tid="callingButtons-showMoreBtn"]',
      '[data-tid="button-more-menu"]',
      '[data-tid="calling-unified-bar-more"]',
      'button[aria-label="Más" i]',
      'button[aria-label="Más acciones" i]',
      'button[aria-label="More" i]',
      'button[aria-label="More actions" i]'
    ];
    for (const sel of explicitSelectors) {
      const candidates = querySelectorAllDeep(sel);
      const visible = candidates.find(
        (el) => !el.closest?.("#froggiteams-widget-root") && isElementVisible(el)
      );
      if (visible) return visible;
    }

    return null;
  }

  /**
   * Searches currently visible popup menus (`[role="menu"]`, `[role="menuitem"]`, etc.)
   * for any item matching the given keywords.
   */
  function findMenuItemByKeywords(keywords) {
    const items = querySelectorAllDeep(
      '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"], [role="option"], [role="menu"] button, [data-tid*="caption" i], [data-tid*="speech" i]'
    );
    return items.find((el) => {
      if (el.closest?.("#froggiteams-widget-root") || !isElementVisible(el)) return false;
      const text = [
        el.id || "",
        el.getAttribute("data-tid") || "",
        el.getAttribute("aria-label") || "",
        el.textContent || ""
      ]
        .join(" ")
        .toLowerCase();
      return keywords.some((kw) => text.includes(kw));
    });
  }

  /**
   * Robust multi-step navigation for MS Teams v2 (https://teams.cloud.microsoft/)
   * 1. Wakes up the call toolbar if auto-hidden
   * 2. Clicks the call toolbar's "Más (...)" button
   * 3. Opens "Idioma y voz" / "Language and speech" (if nested)
   * 4. Clicks "Mostrar subtítulos en vivo" / "Turn on live captions"
   * 5. Sets Spoken Language to Español (México) (or user-selected language)
   */
  async function autoEnableTeamsLiveCaptionsSequence(
    langCode = "es-MX",
    forceLanguageChange = false,
    manualTrigger = false
  ) {
    if (autoCcEnablingInProgress) return;
    autoCcEnablingInProgress = true;

    const banner = document.getElementById("frogg-mode-banner");
    const langObj =
      SUPPORTED_CC_LANGUAGES.find((l) => l.code === langCode) || SUPPORTED_CC_LANGUAGES[0];
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

    try {
      if (banner && manualTrigger) {
        banner.innerHTML = `<span style="color:#a7f3d0;">⏳ Abriendo menú de MS Teams para activar <strong>Subtítulos en vivo (${langObj.label})</strong>...</span>`;
      }

      // Wake up call control bar by dispatching mousemove on the call stage
      const stage =
        querySelectorDeepFirst('[data-tid="call-stage"], [data-tid="modern-calling-screen"]') ||
        document.body;
      stage.dispatchEvent(
        new MouseEvent("mousemove", { bubbles: true, clientX: 400, clientY: 200 })
      );
      await sleep(200);

      // Step 1: Turn on Live Captions if not already open
      if (!areTeamsCaptionsCurrentlyOpen()) {
        // Check if "Subtítulos en vivo" / "Live captions" button is directly on the bar
        let directCaptionBtn = findMenuItemByKeywords([
          "mostrar subtítulos en vivo",
          "activar subtítulos en vivo",
          "turn on live captions",
          "show live captions",
          "livecaptionsbutton",
          "closed-captions-button"
        ]);

        if (directCaptionBtn) {
          simulateFluentInteraction(directCaptionBtn);
          await sleep(700);
        } else {
          // Find the Call Toolbar "Más (...)" button
          const moreBtn = findCallToolbarMoreButton();
          if (moreBtn) {
            // Only click More if a menu isn't already open
            const existingMenu = findMenuItemByKeywords([
              "idioma y voz",
              "language and speech",
              "subtítulos",
              "captions"
            ]);
            if (!existingMenu) {
              simulateFluentInteraction(moreBtn);
              await sleep(550);
            }

            // Check if "Mostrar subtítulos en vivo" is directly in the More menu
            let captionItem = findMenuItemByKeywords([
              "mostrar subtítulos en vivo",
              "activar subtítulos en vivo",
              "subtítulos en vivo",
              "turn on live captions",
              "show live captions",
              "live captions"
            ]);

            // If not directly visible, open the "Idioma y voz" / "Language and speech" submenu
            if (!captionItem) {
              const langAndSpeechItem = findMenuItemByKeywords([
                "idioma y voz",
                "language and speech",
                "speech-and-language",
                "language-and-speech",
                "voz e idioma"
              ]);

              if (langAndSpeechItem) {
                simulateFluentInteraction(langAndSpeechItem, true);
                await sleep(550);

                captionItem = findMenuItemByKeywords([
                  "mostrar subtítulos en vivo",
                  "activar subtítulos en vivo",
                  "subtítulos en vivo",
                  "turn on live captions",
                  "show live captions",
                  "live captions",
                  "subtítulos",
                  "captions"
                ]);
              }
            }

            if (captionItem) {
              simulateFluentInteraction(captionItem);
              await sleep(900);
            }
          }
        }
      }

      // Step 2: Handle Spoken Language modal (if Teams shows "What language is everyone speaking?" or user clicked "Aplicar")
      if (forceLanguageChange || areTeamsCaptionsCurrentlyOpen()) {
        if (forceLanguageChange && areTeamsCaptionsCurrentlyOpen()) {
          const allBtns = querySelectorAllDeep("button, [role='button'], [role='menuitem']");
          const settingsBtn = allBtns.find((el) => {
            if (!isElementVisible(el) || el.closest?.("#froggiteams-widget-root")) return false;
            const txt = [
              el.getAttribute("data-tid") || "",
              el.getAttribute("aria-label") || "",
              el.getAttribute("title") || ""
            ]
              .join(" ")
              .toLowerCase();
            return (
              txt.includes("configuración de subtítulos") ||
              txt.includes("caption settings") ||
              txt.includes("cambiar idioma hablado") ||
              txt.includes("change spoken language") ||
              txt.includes("closed-caption-settings")
            );
          });

          if (settingsBtn) {
            simulateFluentInteraction(settingsBtn);
            await sleep(450);

            const spokenLangItem = findMenuItemByKeywords([
              "idioma hablado",
              "spoken language"
            ]);
            if (spokenLangItem) {
              simulateFluentInteraction(spokenLangItem);
              await sleep(450);
            }
          }
        }

        // If a dialog with a language combobox is open, select Español (México) (or chosen language)
        const dialogCombo = querySelectorDeepFirst(
          '[role="dialog"] [role="combobox"], [role="dialog"] button[aria-haspopup="listbox"], [data-tid*="spoken-language" i]'
        );
        if (dialogCombo && isElementVisible(dialogCombo)) {
          simulateFluentInteraction(dialogCombo);
          await sleep(450);
        }

        const langOpt = findMenuItemByKeywords(langObj.teamsKeywords);
        if (langOpt) {
          simulateFluentInteraction(langOpt);
          await sleep(350);
        }

        const allDialogBtns = querySelectorAllDeep('[role="dialog"] button, button[data-tid*="confirm" i]');
        const confirmBtn = allDialogBtns.find((b) => {
          if (!isElementVisible(b)) return false;
          const t = (b.textContent || b.getAttribute("aria-label") || "").toLowerCase();
          return (
            t.includes("confirmar") ||
            t.includes("actualizar") ||
            t.includes("guardar") ||
            t.includes("confirm") ||
            t.includes("update") ||
            t.includes("save")
          );
        });
        if (confirmBtn) {
          simulateFluentInteraction(confirmBtn);
          await sleep(300);
        }
      }

      if (areTeamsCaptionsCurrentlyOpen()) {
        if (banner) {
          banner.innerHTML = `<span style="color:#34d399;">✅ <strong>Subtítulos en vivo activos (${langObj.label})</strong>. Capturando a todos los asistentes y auto-guardando.</span>`;
        }
      } else if (banner) {
        banner.innerHTML = `<span style="color:#fde68a;">⚠️ Si tu organización oculta el botón automático en Teams, haz clic en la barra de Teams en <strong>Más (...) → Idioma y voz → Mostrar subtítulos en vivo</strong> (o pulsa <strong>🎙️</strong> arriba para transcribir por micrófono).</span>`;
      }
    } catch (e) {
      console.warn("[FroggIteAms Auto-CC]", e);
    } finally {
      autoCcEnablingInProgress = false;
    }
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
