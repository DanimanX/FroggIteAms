// lib/exporter.js
// Formats meeting transcripts to match the exact Tactiq format ("Coordinación Traspaso Cuenta Ms Azure.txt")
// as well as CSV and Requests Report formats.

export function formatMeetingDateTactiq(isoDateString, fallbackDisplay) {
  if (fallbackDisplay) return fallbackDisplay;
  try {
    const d = new Date(isoDateString);
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
  } catch {
    return isoDateString || "Fecha desconocida";
  }
}

/**
 * Generates exact Tactiq-compatible TXT format matching "Coordinación Traspaso Cuenta Ms Azure.txt"
 */
export function exportMeetingToTxt(meeting, includeReport = false) {
  const title = meeting.title || "Reunión Microsoft Teams";
  const startedStr = formatMeetingDateTactiq(meeting.startedAt, meeting.startedDisplay);
  const durationMin = meeting.durationMinutes || Math.max(1, Math.round((meeting.transcript?.length || 1) * 0.5));
  const durationStr = meeting.durationDisplay || `${durationMin} minutes`;
  const participantsList = (meeting.participants && meeting.participants.length > 0)
    ? meeting.participants.join(", ")
    : Array.from(new Set((meeting.transcript || []).map(t => t.speaker))).join(", ");
  const sourceUrl = meeting.sourceUrl || `https://teams.microsoft.com/l/meetup-join/${meeting.id}`;

  const highlights = (meeting.highlights && meeting.highlights.length > 0)
    ? meeting.highlights
    : (meeting.transcript || []).filter(t => t.highlighted);

  const highlightLines = highlights.map((h, idx) => {
    const prefix = idx === 0 ? "    > " : "> ";
    return `${prefix}${h.time} ${h.speaker}: ${h.text}`;
  }).join("\n");

  const transcriptLines = (meeting.transcript || []).map((t, idx) => {
    const prefix = idx === 0 ? "  " : "";
    return `${prefix}${t.time} ${t.speaker}: ${t.text}`;
  }).join("\n");

  let output = `# ${title}

  Meeting started: ${startedStr}
  Duration: ${durationStr}
  Participants: ${participantsList}

  [View original transcript](${sourceUrl})

    

  

  
    ## Highlights

${highlightLines || "    > Sin highlights marcados."}

  ## Transcript

${transcriptLines}
  `;

  if (includeReport && meeting.requestsReport) {
    output += `\n\n  ## Reporte de Solicitudes y Compromisos (FroggIteAms AI)\n\n`;
    output += `  Resumen Ejecutivo: ${meeting.requestsReport.summary || "Sin resumen"}\n\n`;
    if (meeting.requestsReport.requests && meeting.requestsReport.requests.length > 0) {
      meeting.requestsReport.requests.forEach((req, idx) => {
        output += `  ${idx + 1}. [${req.priority || "Media"}] (${req.timestamp || "--:--"}) ${req.description}\n`;
        output += `     - Solicitante: ${req.requester || "N/A"} -> Responsable: ${req.assignee || "Por asignar"}\n`;
        output += `     - Plazo: ${req.deadline || "No especificado"} | Estado: ${req.status || "Pendiente"}\n`;
      });
    }
    if (meeting.requestsReport.keyAgreements && meeting.requestsReport.keyAgreements.length > 0) {
      output += `\n  Acuerdos Clave:\n`;
      meeting.requestsReport.keyAgreements.forEach(a => {
        output += `  * ${a}\n`;
      });
    }
  }

  return output;
}

/**
 * Generates UTF-8 CSV (with BOM for Excel compatibility) of the meeting transcript
 */
export function exportMeetingToCsv(meeting) {
  const escapeCsv = (val) => {
    const str = String(val ?? "").replace(/"/g, '""');
    return `"${str}"`;
  };

  const headers = [
    "Meeting Title",
    "Meeting Date",
    "Timestamp",
    "Speaker",
    "Message",
    "Is Highlight",
    "Is Action Request"
  ];

  const highlightSet = new Set(
    ((meeting.highlights || []).map(h => `${h.time}|${h.speaker}|${h.text}`))
  );

  const rows = [headers.map(escapeCsv).join(",")];
  const dateStr = formatMeetingDateTactiq(meeting.startedAt, meeting.startedDisplay);

  for (const entry of (meeting.transcript || [])) {
    const isHighlight = Boolean(entry.highlighted || highlightSet.has(`${entry.time}|${entry.speaker}|${entry.text}`));
    const isRequest = Boolean(entry.isRequest || /\b(necesitamos|por favor|podrías|hay que|dejar listo|aprobar|enviar|mandar|agendar|coordinar)\b/i.test(entry.text));
    rows.push([
      meeting.title || "Reunión MS Teams",
      dateStr,
      entry.time,
      entry.speaker,
      entry.text,
      isHighlight ? "YES" : "NO",
      isRequest ? "YES" : "NO"
    ].map(escapeCsv).join(","));
  }

  return "\uFEFF" + rows.join("\r\n");
}

/**
 * Generates a dedicated CSV or Markdown/TXT report of the requests & commitments made in the meeting
 */
export function exportRequestsReportToTxt(meeting) {
  const report = meeting.requestsReport;
  const dateStr = formatMeetingDateTactiq(meeting.startedAt, meeting.startedDisplay);
  if (!report) {
    return `# Reporte de Solicitudes - ${meeting.title}\nFecha: ${dateStr}\n\nNo se ha generado un reporte aún.`;
  }

  let lines = [
    `# Reporte de Solicitudes y Compromisos - ${meeting.title}`,
    `Fecha de Reunión: ${dateStr}`,
    `Duración: ${meeting.durationDisplay || `${meeting.durationMinutes || 0} minutes`}`,
    `Participantes (${(meeting.participants || []).length}): ${(meeting.participants || []).join(", ")}`,
    `Generado por: ${report.provider || "FroggIteAms AI"}`,
    ``,
    `## 1. Resumen Ejecutivo`,
    `${report.summary || "Sin resumen disponible."}`,
    ``,
    `## 2. Solicitudes Realizadas en la Reunión (${(report.requests || []).length})`
  ];

  (report.requests || []).forEach((req, idx) => {
    lines.push(
      ``,
      `### [${req.id || `SOL-${String(idx + 1).padStart(2, "0")}`}] ${req.description}`,
      `- **Solicitado por:** ${req.requester || "Participante"}`,
      `- **Asignado a / Responsable:** ${req.assignee || "Por confirmar"}`,
      `- **Prioridad:** ${req.priority || "Media"}`,
      `- **Fecha Límite / Plazo:** ${req.deadline || "Por definir"}`,
      `- **Minuto en reunión:** ${req.timestamp || "--:--"}`,
      `- **Estado:** ${req.status || "Pendiente"}`
    );
  });

  if (report.keyAgreements && report.keyAgreements.length > 0) {
    lines.push(``, `## 3. Acuerdos Clave y Próximos Pasos`);
    report.keyAgreements.forEach(item => {
      lines.push(`- ${item}`);
    });
  }

  return lines.join("\n");
}

export function exportRequestsReportToCsv(meeting) {
  const escapeCsv = (val) => `"${String(val ?? "").replace(/"/g, '""')}"`;
  const headers = [
    "ID",
    "Reunión",
    "Minuto",
    "Solicitante",
    "Responsable",
    "Solicitud / Compromiso",
    "Prioridad",
    "Plazo / Fecha Límite",
    "Estado"
  ];
  const rows = [headers.map(escapeCsv).join(",")];
  const reqs = meeting.requestsReport?.requests || [];
  for (const r of reqs) {
    rows.push([
      r.id || "REQ",
      meeting.title || "Reunión MS Teams",
      r.timestamp || "--:--",
      r.requester || "",
      r.assignee || "",
      r.description || "",
      r.priority || "Media",
      r.deadline || "",
      r.status || "Pendiente"
    ].map(escapeCsv).join(","));
  }
  return "\uFEFF" + rows.join("\r\n");
}

export function sanitizeFilename(name) {
  return (name || "Reunion_Teams").replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").trim();
}
