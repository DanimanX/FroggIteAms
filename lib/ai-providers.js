// lib/ai-providers.js
// Multi-provider AI engine for FroggIteAms:
// Supports:
// 1. Google Gemini via OAuth 2.0 Bearer Token or API Key (gemini-2.0-flash / gemini-1.5-flash)
// 2. DeepSeek via API Key (deepseek-chat)
// 3. OpenAI / Groq / Custom OpenAI-compatible endpoint
// 4. Local Heuristic Fallback (so the user can always generate reports even offline or without tokens)

export const DEFAULT_AI_CONFIG = {
  activeProvider: "gemini", // 'gemini' | 'deepseek' | 'openai' | 'custom' | 'local'
  gemini: {
    authMode: "apikey", // 'apikey' | 'oauth'
    apiKey: "",
    oauthToken: "",
    oauthClientId: "",
    model: "gemini-2.0-flash"
  },
  deepseek: {
    apiKey: "",
    model: "deepseek-chat",
    endpoint: "https://api.deepseek.com/chat/completions"
  },
  openai: {
    apiKey: "",
    model: "gpt-4o-mini",
    endpoint: "https://api.openai.com/v1/chat/completions"
  },
  custom: {
    apiKey: "",
    model: "llama-3.3-70b-versatile",
    endpoint: "https://api.groq.com/openai/v1/chat/completions"
  }
};

const SYSTEM_PROMPT_REQUESTS_REPORT = `Eres FroggIteAms AI, un asistente experto en análisis de reuniones corporativas de Microsoft Teams.
A partir de la transcripción de la reunión, debes identificar con precisión TODAS las solicitudes realizadas en la reunión (compromisos, tareas pedidas, aprobaciones requeridas, envíos de links/accesos, coordinaciones), los highlights más importantes y los acuerdos clave.

Devuelve ÚNICAMENTE un objeto JSON válido (sin bloques markdown \`\`\`json) con la siguiente estructura exacta:
{
  "summary": "Resumen ejecutivo claro y conciso de la reunión en español (2 a 4 oraciones).",
  "requests": [
    {
      "id": "REQ-01",
      "requester": "Nombre de quien solicita",
      "assignee": "Nombre del responsable o equipo a quien se le pide",
      "description": "Descripción clara y accionable de la solicitud realizada en la reunión",
      "priority": "Crítica | Alta | Media | Baja",
      "deadline": "Plazo mencionado (ej. Hoy 15:00 hrs, Mañana mediodía, Por definir)",
      "timestamp": "MM:SS",
      "status": "Pendiente"
    }
  ],
  "keyAgreements": [
    "Acuerdo clave 1",
    "Acuerdo clave 2"
  ],
  "suggestedHighlights": [
    {
      "time": "MM:SS",
      "speaker": "Nombre del participante",
      "text": "Cita textual relevante"
    }
  ]
}`;

export async function generateRequestsReportWithAI(meeting, aiConfig = DEFAULT_AI_CONFIG) {
  const provider = aiConfig?.activeProvider || "gemini";
  const transcriptText = (meeting.transcript || [])
    .map(t => `[${t.time}] ${t.speaker}: ${t.text}`)
    .join("\n");

  if (!transcriptText.trim()) {
    throw new Error("La reunión no tiene líneas de transcripción para analizar.");
  }

  const userPrompt = `Título de la reunión: ${meeting.title}
Fecha: ${meeting.startedDisplay || meeting.startedAt}
Participantes: ${(meeting.participants || []).join(", ")}

Transcripción completa:
${transcriptText.slice(0, 48000)}`;

  try {
    if (provider === "local") {
      return generateHeuristicRequestsReport(meeting, "Motor Heurístico Local (FroggIteAms)");
    }

    if (provider === "gemini") {
      const gem = aiConfig.gemini || {};
      const tokenOrKey = gem.authMode === "oauth" ? gem.oauthToken : gem.apiKey;
      if (!tokenOrKey) {
        return generateHeuristicRequestsReport(
          meeting,
          "FroggIteAms Smart Analyzer (Configura tu token de Gemini/DeepSeek en Ajustes IA para análisis neuronal)"
        );
      }
      return await callGeminiProvider(userPrompt, gem);
    }

    if (provider === "deepseek") {
      const ds = aiConfig.deepseek || {};
      if (!ds.apiKey) {
        return generateHeuristicRequestsReport(
          meeting,
          "FroggIteAms Smart Analyzer (Sin API Key de DeepSeek configurada, usando motor heurístico)"
        );
      }
      return await callOpenAICompatibleProvider(userPrompt, {
        endpoint: ds.endpoint || "https://api.deepseek.com/chat/completions",
        apiKey: ds.apiKey,
        model: ds.model || "deepseek-chat",
        providerLabel: `DeepSeek (${ds.model || "deepseek-chat"})`
      });
    }

    if (provider === "openai" || provider === "custom") {
      const cfg = aiConfig[provider] || {};
      if (!cfg.apiKey) {
        return generateHeuristicRequestsReport(
          meeting,
          `FroggIteAms Smart Analyzer (Sin API Key en ${provider}, usando motor heurístico)`
        );
      }
      return await callOpenAICompatibleProvider(userPrompt, {
        endpoint: cfg.endpoint || "https://api.openai.com/v1/chat/completions",
        apiKey: cfg.apiKey,
        model: cfg.model || "gpt-4o-mini",
        providerLabel: `${provider.toUpperCase()} (${cfg.model || "default"})`
      });
    }

    return generateHeuristicRequestsReport(meeting, "FroggIteAms Local Engine");
  } catch (err) {
    console.warn("[FroggIteAms AI] Fallback to heuristic due to error:", err);
    const fallback = generateHeuristicRequestsReport(
      meeting,
      `FroggIteAms Respaldo Local (Aviso IA: ${err.message})`
    );
    fallback.warning = err.message;
    return fallback;
  }
}

async function callGeminiProvider(userPrompt, geminiConfig) {
  const model = geminiConfig.model || "gemini-2.0-flash";
  const isOAuth = geminiConfig.authMode === "oauth";
  const baseUrl = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
  const url = isOAuth
    ? baseUrl
    : `${baseUrl}?key=${encodeURIComponent(geminiConfig.apiKey.trim())}`;

  const headers = {
    "Content-Type": "application/json"
  };
  if (isOAuth) {
    headers["Authorization"] = `Bearer ${geminiConfig.oauthToken.trim()}`;
  }

  const response = await fetch(url, {
    method: "POST",
    headers,
    body: JSON.stringify({
      systemInstruction: {
        parts: [{ text: SYSTEM_PROMPT_REQUESTS_REPORT }]
      },
      contents: [
        {
          role: "user",
          parts: [{ text: userPrompt }]
        }
      ],
      generationConfig: {
        temperature: 0.2,
        responseMimeType: "application/json"
      }
    })
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Error Gemini (${response.status}): ${errText.slice(0, 180)}`);
  }

  const data = await response.json();
  const rawText = data?.candidates?.[0]?.content?.parts?.[0]?.text || "{}";
  const parsed = safeParseJson(rawText);

  return normalizeReportObject(
    parsed,
    `Google Gemini (${model} vía ${isOAuth ? "OAuth 2.0" : "API Key"})`
  );
}

async function callOpenAICompatibleProvider(userPrompt, { endpoint, apiKey, model, providerLabel }) {
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey.trim()}`
    },
    body: JSON.stringify({
      model,
      temperature: 0.2,
      messages: [
        { role: "system", content: SYSTEM_PROMPT_REQUESTS_REPORT },
        { role: "user", content: userPrompt }
      ]
    })
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Error ${providerLabel} (${response.status}): ${errText.slice(0, 180)}`);
  }

  const data = await response.json();
  const rawText = data?.choices?.[0]?.message?.content || "{}";
  const parsed = safeParseJson(rawText);

  return normalizeReportObject(parsed, providerLabel);
}

/**
 * Cleanly cleans or refines live speech / caption text using the user's configured AI provider
 */
export async function refineLiveCaptionWithAI(speaker, text, aiConfig) {
  if (!text || text.length < 15) return text;
  const provider = aiConfig?.activeProvider || "local";
  if (provider === "local") return text;

  try {
    const prompt = `Corrige únicamente puntuación y errores evidentes de transcripción automática de voz en español manteniendo intacto el sentido y palabras del hablante (${speaker}). Devuelve SOLO el texto corregido sin comillas:\n\n${text}`;
    if (provider === "gemini") {
      const gem = aiConfig.gemini || {};
      const tokenOrKey = gem.authMode === "oauth" ? gem.oauthToken : gem.apiKey;
      if (!tokenOrKey) return text;
      const model = gem.model || "gemini-2.0-flash";
      const baseUrl = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
      const url = gem.authMode === "oauth" ? baseUrl : `${baseUrl}?key=${encodeURIComponent(gem.apiKey.trim())}`;
      const headers = { "Content-Type": "application/json" };
      if (gem.authMode === "oauth") headers["Authorization"] = `Bearer ${gem.oauthToken.trim()}`;

      const res = await fetch(url, {
        method: "POST",
        headers,
        body: JSON.stringify({
          contents: [{ role: "user", parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.1, maxOutputTokens: 250 }
        })
      });
      if (!res.ok) return text;
      const data = await res.json();
      return data?.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || text;
    }

    if (provider === "deepseek" || provider === "openai" || provider === "custom") {
      const cfg = aiConfig[provider] || {};
      if (!cfg.apiKey) return text;
      const res = await fetch(cfg.endpoint || "https://api.deepseek.com/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${cfg.apiKey.trim()}`
        },
        body: JSON.stringify({
          model: cfg.model || "deepseek-chat",
          temperature: 0.1,
          max_tokens: 250,
          messages: [{ role: "user", content: prompt }]
        })
      });
      if (!res.ok) return text;
      const data = await res.json();
      return data?.choices?.[0]?.message?.content?.trim() || text;
    }
  } catch {
    return text;
  }
  return text;
}

function safeParseJson(rawText) {
  const cleaned = String(rawText || "")
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/```\s*$/i, "")
    .trim();
  try {
    return JSON.parse(cleaned);
  } catch {
    const firstBrace = cleaned.indexOf("{");
    const lastBrace = cleaned.lastIndexOf("}");
    if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
      return JSON.parse(cleaned.slice(firstBrace, lastBrace + 1));
    }
    throw new Error("La respuesta del modelo IA no contenía un JSON válido.");
  }
}

function normalizeReportObject(parsed, providerLabel) {
  const requests = Array.isArray(parsed.requests)
    ? parsed.requests.map((r, i) => ({
        id: r.id || `REQ-${String(i + 1).padStart(2, "0")}`,
        requester: r.requester || "Participante",
        assignee: r.assignee || "Por asignar",
        description: r.description || "Solicitud identificada en la reunión",
        priority: r.priority || "Media",
        deadline: r.deadline || "Por definir",
        timestamp: r.timestamp || "--:--",
        status: r.status || "Pendiente"
      }))
    : [];

  return {
    generatedAt: new Date().toISOString(),
    provider: providerLabel,
    summary: parsed.summary || "Análisis completado sobre las intervenciones de la reunión.",
    requests,
    keyAgreements: Array.isArray(parsed.keyAgreements) ? parsed.keyAgreements : [],
    suggestedHighlights: Array.isArray(parsed.suggestedHighlights) ? parsed.suggestedHighlights : []
  };
}

/**
 * Smart deterministic Spanish heuristic analyzer when no external token is configured
 * or when offline. Detects action verbs, requests, deadlines, and participants.
 */
export function generateHeuristicRequestsReport(meeting, providerLabel = "FroggIteAms Smart Analyzer") {
  const transcript = meeting.transcript || [];
  const participants = meeting.participants || Array.from(new Set(transcript.map(t => t.speaker)));

  const requestPatterns = [
    /\b(necesitamos que|por favor|podrías|podríamos dejar|hay que|tiene que ir la aprobación|dejemos|fijemos|asegúrate que|enviar|mandar|aceptar|aprobar|agenda|agendar)\b/i
  ];

  const detectedRequests = [];
  const keyAgreements = [];

  for (const item of transcript) {
    const text = item.text || "";
    if (text.length < 22) continue;

    if (requestPatterns.some(rx => rx.test(text))) {
      // Find mentioned participant in the line if any
      let assignee = "Equipo / Por confirmar";
      for (const p of participants) {
        const firstName = p.split(/[\s,]+/)[0];
        if (firstName && firstName.length > 2 && new RegExp(`\\b${firstName}\\b`, "i").test(text) && p !== item.speaker) {
          assignee = p;
          break;
        }
      }

      let deadline = "Por definir";
      if (/\b(hoy|15:00|tarde)\b/i.test(text)) deadline = "Hoy";
      else if (/\b(mañana|mediodía|12:00|1230)\b/i.test(text)) deadline = "Mañana";

      let priority = "Media";
      if (/\b(importante|urgente|vencen|terminan|aprobación|sin intermitencia)\b/i.test(text)) {
        priority = "Alta";
      }

      detectedRequests.push({
        id: `REQ-${String(detectedRequests.length + 1).padStart(2, "0")}`,
        requester: item.speaker,
        assignee,
        description: text,
        priority,
        deadline,
        timestamp: item.time,
        status: "Pendiente"
      });
    }

    if (/\b(entonces mañana|quedamos|de acuerdo|perfecto|cerrado)\b/i.test(text) && text.length > 35 && keyAgreements.length < 5) {
      keyAgreements.push(`[${item.time}] ${item.speaker}: ${text}`);
    }
  }

  // If none matched strict patterns, take top highlighted items
  if (detectedRequests.length === 0) {
    const source = (meeting.highlights && meeting.highlights.length > 0)
      ? meeting.highlights.slice(0, 5)
      : transcript.slice(0, 5);
    source.forEach((h, i) => {
      detectedRequests.push({
        id: `REQ-${String(i + 1).padStart(2, "0")}`,
        requester: h.speaker,
        assignee: "Participantes de la reunión",
        description: h.text,
        priority: "Media",
        deadline: "Próxima sesión",
        timestamp: h.time,
        status: "Pendiente"
      });
    });
  }

  return {
    generatedAt: new Date().toISOString(),
    provider: providerLabel,
    summary: `Reunión "${meeting.title}" con ${participants.length} participantes y ${transcript.length} intervenciones registradas. Se detectaron ${detectedRequests.length} solicitudes y compromisos accionables.`,
    requests: detectedRequests.slice(0, 12),
    keyAgreements: keyAgreements.length > 0
      ? keyAgreements
      : ["Se registraron los puntos principales y se asignaron seguimientos en la transcripción."]
  };
}
