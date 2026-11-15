// ============================================================
// Chat.gs — Chat del equipo (una sola sala, solo texto)
// ============================================================
//
// Los mensajes se guardan en la hoja CHAT_EQUIPO (una fila por mensaje). La app los
// consulta cada pocos segundos mientras el chat está abierto. Nunca se borra nada.
//
// Columnas: ID | Fecha | Autor | Id_autor | Texto
//   Id_autor: 'admin' para Álvaro (con cualquiera de sus enlaces), o el ID del empleado.
// ============================================================

var CHAT_ = {
  SHEET:     'CHAT_EQUIPO',
  HEAD:      ['ID', 'Fecha', 'Autor', 'Id_autor', 'Texto'],
  MAX_TEXTO: 1000,   // letras por mensaje
  INICIAL:   60,     // mensajes que se ven al abrir el chat
  MAX_LEER:  100,    // tope de mensajes nuevos por consulta
  CACHE_ULT: 'chat_ultimo_id'
};

function _chat_hoja_() {
  var sh = getOrCreateSheet(CHAT_.SHEET);
  if (sh.getLastRow() === 0) { sh.appendRow(CHAT_.HEAD); applyHeaderStyle_(sh, CHAT_.HEAD.length); }
  return sh;
}

/** ID del último mensaje leído directamente de la hoja (0 si no hay ninguno). */
function _chat_idEnHoja_(sh) {
  var last = sh.getLastRow();
  return last < 2 ? 0 : (Number(sh.getRange(last, 1).getValue()) || 0);
}

/** ID del último mensaje. Se guarda en caché para que consultar "¿hay algo nuevo?" no lea la hoja. */
function _chat_ultimoId_() {
  var c = CacheService.getScriptCache(), v = c.get(CHAT_.CACHE_ULT);
  if (v !== null) return Number(v) || 0;
  var id = _chat_idEnHoja_(_chat_hoja_());
  try { c.put(CHAT_.CACHE_ULT, String(id), 21600); } catch (_) {}
  return id;
}

/** Quién escribe: Álvaro es siempre 'admin' (también si entra con su ficha de empleado). */
function _chat_idAutor_(user) {
  return user.rol === 'admin' ? 'admin' : String(user.id || '');
}

/**
 * Mensajes del chat. p.desde = ID del último mensaje que ya tiene la pantalla (0 = primera carga:
 * devuelve los últimos). Devuelve { ultimo, mensajes: [{ id, fecha, autor, mio, texto }] }.
 */
function wa_chatLeer_(user, p) {
  var desde = Math.floor(Number(p.desde)) || 0;
  var ult = _chat_ultimoId_();
  if (desde > ult) desde = 0;                                   // la pantalla va por delante de la hoja: recarga
  if (desde > 0 && desde >= ult) return { ultimo: ult, mensajes: [] };   // nada nuevo: sin leer la hoja

  var sh = _chat_hoja_(), last = sh.getLastRow();
  if (last < 2) return { ultimo: 0, mensajes: [] };
  var n = Math.min(last - 1, desde > 0 ? CHAT_.MAX_LEER : CHAT_.INICIAL);
  var vals = sh.getRange(last - n + 1, 1, n, CHAT_.HEAD.length).getValues();
  var yo = _chat_idAutor_(user), tz = Session.getScriptTimeZone(), out = [];
  vals.forEach(function(r) {
    var id = Number(r[0]);
    if (!id || id <= desde) return;
    out.push({
      id: id,
      fecha: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd HH:mm') : '',
      autor: String(r[2] || ''),
      mio: String(r[3] || '') === yo,
      texto: String(r[4] || '')
    });
  });
  return { ultimo: ult, mensajes: out };
}

/** Añade un mensaje al chat. p.texto = el mensaje. */
function wa_chatEnviar_(user, p) {
  var texto = String(p.texto == null ? '' : p.texto).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim();
  if (!texto) _wa_fail_('Escribe un mensaje.');
  if (texto.length > CHAT_.MAX_TEXTO) _wa_fail_('El mensaje es demasiado largo (máximo ' + CHAT_.MAX_TEXTO + ' letras).');
  return _wa_withLock_(function() {
    var sh = _chat_hoja_();
    var id = _chat_idEnHoja_(sh) + 1, row = sh.getLastRow() + 1;
    sh.getRange(row, 2).setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange(row, 3, 1, 3).setNumberFormat('@');             // formato texto: nunca se interpreta como fórmula
    sh.getRange(row, 1, 1, CHAT_.HEAD.length).setValues([[id, new Date(), String(user.nombre || ''), _chat_idAutor_(user), texto]]);
    try { CacheService.getScriptCache().put(CHAT_.CACHE_ULT, String(id), 21600); } catch (_) {}
    return { id: id };
  });
}
