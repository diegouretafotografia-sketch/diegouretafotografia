/**
 * Diego Ureta Fotografía — receptor de cotizaciones  (v2)
 * - Guarda cada solicitud en la hoja "Cotizaciones" con ID correlativo y estado.
 * - Te envía un correo con botón directo a WhatsApp del cliente.
 * - Envía al cliente un correo de confirmación.
 *
 * INSTALACIÓN
 * 1) En la Google Sheet: Extensiones > Apps Script. Pega este código y guarda (Ctrl+S).
 * 2) Selector de funciones > "setup" > Ejecutar. Acepta los permisos.
 * 3) Selector de funciones > "test" > Ejecutar. Revisa tu correo y la pestaña "Cotizaciones".
 * 4) Implementar > Nueva implementación > App web > Ejecutar como: Yo / Acceso: Cualquier usuario.
 * 5) Copia la URL que termina en /exec.  (Si la abres en el navegador verás "OK" — es normal.)
 *
 * Cada vez que cambies este código: Implementar > Administrar implementaciones > ✏️ > Versión: Nueva > Implementar.
 */

const CONFIG = {
  NOTIFY_TO: 'diegouretafotografia@gmail.com',   // dónde te llegan las cotizaciones
  FROM_NAME: 'Diego Ureta Fotografía',
  SHEET: 'Cotizaciones',
  TZ: 'America/Santiago',
  REPLY_TO_CLIENT: true,                          // correo automático de confirmación al cliente
  INSTAGRAM: 'https://instagram.com/diegouretafotografia'
};

const HEAD = ['ID', 'Fecha solicitud', 'Estado', 'Servicio', 'Nombre', 'WhatsApp', 'Link WhatsApp', 'Email',
  'Fecha evento', 'Lugar', 'Formato', 'Origen', 'Idioma', 'Detalle', 'Respuestas completas', 'Notas'];
const ESTADOS = ['Nuevo', 'Contactado', 'Cotizado', 'Reservado', 'Perdido'];

/* ---------- Puntos de entrada ---------- */

function doGet() {
  return ContentService.createTextOutput('OK — Receptor de cotizaciones de Diego Ureta Fotografía activo.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const d = parse_(e);
    if (!d || d.website) return out_({ ok: true });                     // sin datos o bot
    if (!d.nombre || !d.email || !d.whatsapp) return out_({ ok: false, error: 'Faltan datos' });

    const sh = sheet_();
    const id = nextId_(sh);
    const wa = phone_(d.pais, d.whatsapp);
    const waLink = waLink_(wa, d);
    const now = new Date();

    const skip = ['lang', 'labels', 'servicio', 'servicio_nombre', 'pais', 'website', 'nombre', 'whatsapp', 'email'];
    const answers = Object.keys(d)
      .filter(k => skip.indexOf(k) < 0 && String(d[k]).trim() !== '')
      .map(k => [label_(d, k), String(d[k])]);
    const pick = suf => { const k = Object.keys(d).find(k => /^[a-z]_/.test(k) && k.endsWith(suf)); return k ? d[k] : ''; };

    sh.appendRow([id, now, 'Nuevo', d.servicio_nombre || d.servicio || '', d.nombre, '+' + wa, waLink, d.email,
      pick('_fecha'), pick('_lugar'), pick('_formato'), d.origen || '', d.lang || 'ES', d.detalle || '',
      answers.map(a => a[0] + ': ' + a[1]).join('\n'), '']);

    notifyMe_(id, d, wa, waLink, answers);
    if (CONFIG.REPLY_TO_CLIENT) notifyClient_(d);
    return out_({ ok: true, id: id });
  } catch (err) {
    console.error(err);
    try { MailApp.sendEmail(CONFIG.NOTIFY_TO, '⚠️ Error en formulario web', String(err) + '\n\n' + (e && e.postData ? e.postData.contents : '')); } catch (_) {}
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ---------- Correos ---------- */

function notifyMe_(id, d, wa, waLink, answers) {
  const rows = answers.map(a =>
    '<tr><td style="padding:8px 12px;color:#9a9286;vertical-align:top;width:42%;border-bottom:1px solid #222">' + esc_(a[0]) + '</td>' +
    '<td style="padding:8px 12px;color:#ece6da;border-bottom:1px solid #222">' + esc_(a[1]).replace(/\n/g, '<br>') + '</td></tr>').join('');
  const html =
    '<div style="background:#0c0c0c;padding:30px;font-family:Helvetica,Arial,sans-serif">' +
    '<div style="max-width:600px;margin:auto;border:1px solid #2a2722;padding:28px">' +
    '<div style="color:#c9a86a;letter-spacing:4px;font-size:11px">NUEVA COTIZACIÓN · ' + id + '</div>' +
    '<h2 style="color:#ece6da;font-weight:400;margin:10px 0 4px">' + esc_(d.nombre) + '</h2>' +
    '<div style="color:#e2c891;margin-bottom:6px">' + esc_(d.servicio_nombre || '') + '</div>' +
    '<div style="color:#9a9286;font-size:13px;margin-bottom:22px">+' + wa + ' · ' + esc_(d.email) + ' · ' + esc_(d.origen || '—') + '</div>' +
    '<a href="' + waLink + '" style="display:inline-block;background:#25D366;color:#fff;padding:13px 22px;text-decoration:none;font-weight:bold;border-radius:4px">💬 Responder por WhatsApp</a> ' +
    '<a href="mailto:' + esc_(d.email) + '" style="display:inline-block;border:1px solid #c9a86a;color:#e2c891;padding:12px 20px;text-decoration:none;border-radius:4px;margin-left:6px">✉️ Responder por correo</a>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:24px;font-size:14px">' + rows + '</table>' +
    '<p style="margin-top:22px"><a href="' + SpreadsheetApp.getActive().getUrl() + '" style="color:#c9a86a">Abrir planilla de cotizaciones →</a></p>' +
    '</div></div>';
  MailApp.sendEmail({
    to: CONFIG.NOTIFY_TO, replyTo: d.email, name: 'Web ' + CONFIG.FROM_NAME,
    subject: '📷 ' + id + ' · ' + (d.servicio_nombre || '') + ' · ' + d.nombre,
    htmlBody: html, body: answers.map(a => a[0] + ': ' + a[1]).join('\n')
  });
}

function notifyClient_(d) {
  const en = d.lang === 'EN', n = first_(d.nombre);
  const subject = en ? 'I received your request — ' + CONFIG.FROM_NAME : 'Recibí tu solicitud — ' + CONFIG.FROM_NAME;
  const body = en
    ? 'Hi ' + n + ',\n\nThanks for writing! I received your request for "' + d.servicio_nombre + '". I\'ll get back to you within 24 hours by WhatsApp or email with availability and a proposal.\n\nMeanwhile you can see my latest work here: ' + CONFIG.INSTAGRAM + '\n\nDiego Ureta\n' + CONFIG.FROM_NAME
    : 'Hola ' + n + ',\n\n¡Gracias por escribirme! Recibí tu solicitud de "' + d.servicio_nombre + '". Te responderé en menos de 24 horas por WhatsApp o correo con disponibilidad y una propuesta.\n\nMientras tanto puedes ver mi trabajo reciente aquí: ' + CONFIG.INSTAGRAM + '\n\nDiego Ureta\n' + CONFIG.FROM_NAME;
  try { MailApp.sendEmail({ to: d.email, replyTo: CONFIG.NOTIFY_TO, name: CONFIG.FROM_NAME, subject: subject, body: body }); } catch (_) {}
}

/* ---------- Utilidades ---------- */

function parse_(e) {
  if (!e || !e.postData || !e.postData.contents) return null;
  try { return JSON.parse(e.postData.contents); } catch (_) { return e.parameter || null; }
}
function label_(d, k) { return (d.labels && d.labels[k]) ? d.labels[k] : k; }
function phone_(pais, num) {
  let n = String(num || '').replace(/\D/g, '');
  pais = String(pais || '').replace(/\D/g, '');
  if (pais) {
    n = n.replace(/^0+/, '');
    if (pais === '56' && n.length === 8) n = '9' + n;                 // 8 dígitos sin el 9
    if (n.indexOf(pais) !== 0 || (pais === '56' && n.length === 9)) n = pais + n;
  }
  return n;
}
function waLink_(wa, d) {
  const en = d.lang === 'EN';
  const msg = en
    ? 'Hi ' + first_(d.nombre) + ', this is Diego Ureta from ' + CONFIG.FROM_NAME + ' 📷. I got your quote request for ' + (d.servicio_nombre || '').toLowerCase() + '. Do you have a few minutes to chat?'
    : 'Hola ' + first_(d.nombre) + ', soy Diego Ureta de ' + CONFIG.FROM_NAME + ' 📷. Recibí tu solicitud de cotización para ' + (d.servicio_nombre || '').toLowerCase() + '. ¿Tienes unos minutos para conversar?';
  return 'https://wa.me/' + wa + '?text=' + encodeURIComponent(msg);
}
function nextId_(sh) {
  const last = sh.getLastRow();
  let n = 0;
  if (last > 1) { const m = String(sh.getRange(last, 1).getValue()).match(/(\d+)$/); if (m) n = +m[1]; }
  return 'COT-' + String(n + 1).padStart(4, '0');
}
function first_(s) { return String(s || '').trim().split(/\s+/)[0]; }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET, 0);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#c9a86a');
    sh.getRange('B2:B').setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build());
    sh.getRange('O2:O').setWrap(true);
    sh.setColumnWidth(15, 380);
  }
  return sh;
}

/* ---------- Ejecutar a mano desde el editor ---------- */

/** Paso 2: crea la hoja y pide los permisos. */
function setup() { sheet_(); Logger.log('Hoja lista: ' + CONFIG.SHEET); }

/** Paso 3: simula una solicitud real. */
function test() {
  const r = doPost({ postData: { contents: JSON.stringify({
    nombre: 'Prueba Test', pais: '56', whatsapp: '9 2034 3751', email: CONFIG.NOTIFY_TO,
    servicio: 'B', servicio_nombre: 'Pedida de mano', b_fecha: 'Febrero 2027',
    b_sorpresa: 'Sí, necesito que sea encubierto y coordinar la logística contigo', b_lugar: 'Mirador al atardecer',
    detalle: 'Esto es una prueba.', origen: 'Instagram', lang: 'ES',
    labels: { b_fecha: 'Fecha estimada o exacta de la pedida', b_sorpresa: '¿Será una propuesta 100% sorpresa?', b_lugar: '¿Ya tienes pensado el lugar y la hora?', detalle: 'Detalle extra', origen: '¿Cómo llegaste a mi página?' }
  }) } });
  Logger.log(r.getContent());
}
