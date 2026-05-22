/**
 * Google Apps Script para guardar/actualizar tickets en Google Sheets.
 *
 * Pasos:
 * 1. Crea un Google Sheet.
 * 2. Extensiones -> Apps Script.
 * 3. Pega este archivo completo.
 * 4. Deploy -> New deployment -> Web app.
 * 5. Execute as: Me.
 * 6. Who has access: Anyone with the link.
 * 7. Copia la URL /exec y guárdala en Render como:
 *    GOOGLE_SHEETS_WEBHOOK_URL=https://script.google.com/macros/s/XXXXX/exec
 */

const SHEET_NAME = 'Tickets';

const HEADERS = [
  'ID Interno',
  'Ticket Slack',
  'Última acción',
  'Última sincronización',
  'Fecha Apertura',
  'Fecha Asignación',
  'Fecha Llegada',
  'Fecha Cierre',
  'Estado',
  'Fecha Cancelación',
  'Motivo Cancelación',
  'Cancelado Por',
  'Site',
  'Torre',
  'Zona',
  'ACC Nombre',
  'ACC IP',
  'Edge Nombre',
  'Edge IP',
  'Modelo',
  'Tipo',
  'Afectados',
  'Técnico',
  'Semáforo',
  'Reincidente',
  'Criterio Reincidencia',
  'Motivo Reincidencia',
  'Evento #',
  'Observación',
  'Topología',
  'Ubicación',
  'Causa Raíz',
  'Clasificación',
  'Acciones',
  'Estado ACC',
  'Estado Edge',
  'Estado Switch',
  'Potencia / Velocidad',
  'Solución',
  'Escalamiento',
  'Riesgo',
  'Descripción Riesgo',
  'Recomendación',
  'MTTR Minutos'
];

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet_(ss);

    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const ticket = body.ticket || body;
    const action = body.action || 'upsert';
    const syncedAt = body.synced_at || new Date().toISOString();

    if (!ticket || !ticket.id) {
      return json_({ ok: false, error: 'Payload inválido: falta ticket.id' });
    }

    const row = buildRow_(ticket, action, syncedAt);
    const rowIndex = findTicketRow_(sheet, ticket.id);

    if (rowIndex) {
      sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return json_({ ok: true, action, ticket_id: ticket.id, updated: Boolean(rowIndex) });
  } catch (err) {
    return json_({ ok: false, error: err.message, stack: err.stack });
  }
}

function getOrCreateSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    const width = Math.max(sheet.getLastColumn(), HEADERS.length);
    const firstRow = sheet.getRange(1, 1, 1, width).getValues()[0];
    const mustRefreshHeaders = HEADERS.some((h, i) => String(firstRow[i] || '') !== h);
    if (mustRefreshHeaders) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

function findTicketRow_(sheet, ticketId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const id = String(ticketId);
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === id) return i + 2;
  }
  return null;
}


function formatDateOnly_(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d.getTime())) return value;
  return Utilities.formatDate(d, 'America/Bogota', 'dd/MM/yyyy');
}

function formatTimeOnly_(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d.getTime())) return value;
  return Utilities.formatDate(d, 'America/Bogota', 'HH:mm:ss');
}

function buildRow_(t, action, syncedAt) {
  const cierre = t.cierre || {};
  return [
    t.id || '',
    t.slack_num || '',
    action || '',
    formatTimeOnly_(syncedAt),
    formatDateOnly_(t.fecha_apertura),
    formatDateOnly_(t.fecha_asignacion),
    formatDateOnly_(t.fecha_llegada),
    formatDateOnly_(t.fecha_cierre),
    t.estado || '',
    formatDateOnly_(t.fecha_cancelacion),
    t.motivo_cancelacion || '',
    t.cancelado_por || '',
    t.site || '',
    t.torre || '',
    t.zona || '',
    t.acc_nombre || '',
    t.acc_ip || '',
    t.edge_nombre || '',
    t.edge_ip || '',
    t.modelo || '',
    t.tipo || '',
    t.afectados || 0,
    t.tecnico_nombre || '',
    t.semaforo || '',
    t.es_reincidente ? 'Sí' : 'No',
    t.criterio_reincidencia || '',
    t.motivo_reincidencia || '',
    t.evento_num || '',
    t.observacion || '',
    t.topologia_url || '',
    t.ubicacion_url || '',
    cierre.causa_raiz || '',
    cierre.clasificacion || '',
    Array.isArray(cierre.acciones) ? cierre.acciones.join(', ') : (cierre.acciones || ''),
    cierre.estado_acc || '',
    cierre.estado_edge || '',
    cierre.estado_switch || '',
    cierre.potencia_dbm || '',
    cierre.solucion || '',
    cierre.escalamiento || '',
    cierre.riesgo ? 'Sí' : 'No',
    cierre.desc_riesgo || '',
    cierre.recomendacion || '',
    t.mttr_minutos || ''
  ];
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
