// =============================================================
// Notas.gs — Notas diarias de Álvaro: notas y listas con casillas
// =============================================================
//
// - Cada nota tiene fecha, título, texto y una lista opcional de puntos con casilla.
// - Se pueden compartir con empleados: el empleado la ve en su pantalla y puede marcar
//   los puntos (no cambiar el texto ni borrarla).
// - "Pendiente" = tiene algún punto sin marcar, o (si no tiene lista) no está hecha.
// - Un recordatorio opcional crea un evento con aviso en el Google Calendar de Álvaro
//   (mismo mecanismo que las tareas de Agenda.gs).
// =============================================================

var NOT_ = {
  SHEET: 'NOTAS',
  HEAD: ['ID_Nota', 'Fecha', 'Titulo', 'Texto', 'Items', 'Compartida_Con', 'Hecha', 'Aviso_Hora', 'Aviso_Min', 'ID_Evento', 'Creada', 'Actualizada', 'Autor'],
  COL: { ID: 1, FECHA: 2, TITULO: 3, TEXTO: 4, ITEMS: 5, COMP: 6, HECHA: 7, HORA: 8, AVISO: 9, EVENTO: 10, CREADA: 11, ACT: 12, AUTOR: 13 },
  MAX_ITEMS: 40,
  ID_RE: /^NOT-\d{3,6}$/,
  FILTROS: ['pendientes', 'todas', 'compartidas', 'hechas']
};

function _not_hoja_() {
  var sh = getOrCreateSheet(NOT_.SHEET);
  if (sh.getLastRow() === 0) { sh.appendRow(NOT_.HEAD); applyHeaderStyle_(sh, NOT_.HEAD.length); return sh; }
  _ag_aseguraColumna_(sh, NOT_.COL.AUTOR, 'Autor');   // hojas anteriores a que cada persona tuviera sus notas
  return sh;
}

function _not_items_(raw) {
  try {
    var a = JSON.parse(String(raw || '[]'));
    if (!Array.isArray(a)) return [];
    return a.slice(0, NOT_.MAX_ITEMS).map(function(x) { return { t: String(x && x.t || ''), h: x && x.h === true }; }).filter(function(x) { return x.t; });
  } catch (_) { return []; }
}

function _not_comp_(raw) {
  return String(raw || '').split('|').map(function(s) { return s.trim(); }).filter(Boolean);
}

function _not_todas_() {
  var sh = _not_hoja_(), last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, NOT_.HEAD.length).getValues(), C = NOT_.COL, out = [], ahora = _ag_ahora_();
  vals.forEach(function(r, i) {
    var id = String(r[C.ID - 1] || '').trim();
    if (!id) return;
    var aviso = r[C.AVISO - 1], items = _not_items_(r[C.ITEMS - 1]), hecha = _wa_bool_(r[C.HECHA - 1]);
    var n = {
      row: i + 2, id: id, fecha: _ag_iso_(r[C.FECHA - 1]), titulo: String(r[C.TITULO - 1] || ''), texto: String(r[C.TEXTO - 1] || ''),
      items: items, compartida: _not_comp_(r[C.COMP - 1]), hecha: hecha,
      hora: _wa_horaTxt_(r[C.HORA - 1]), aviso: (aviso === '' || aviso == null || isNaN(Number(aviso))) ? '' : Number(aviso),
      evento: String(r[C.EVENTO - 1] || '').trim(),
      autor: String(r[C.AUTOR - 1] || '').trim() || 'Álvaro'
    };
    n.pendiente = items.length ? items.some(function(x) { return !x.h; }) : !hecha;
    n.atrasado = n.pendiente && _ag_vencido_(n.fecha, n.hora, ahora);
    out.push(n);
  });
  return out;
}

function _not_pub_(n, user) {
  var mia = user ? _not_esMia_(n, user) : true;
  return {
    id: n.id, fecha: n.fecha, titulo: n.titulo, texto: n.texto, items: n.items, compartida: n.compartida, hecha: n.hecha,
    hora: mia ? n.hora : '', aviso: mia ? n.aviso : '', calendario: mia ? !!n.evento : false, pendiente: n.pendiente, atrasado: n.atrasado,
    nItems: n.items.length, nHechos: n.items.filter(function(x) { return x.h; }).length,
    autor: n.autor, mia: mia
  };
}

/** ¿La escribió esta persona? (lo anterior a esta versión es de Álvaro) */
function _not_esMia_(n, user) { return _wa_norm_(n.autor) === _wa_norm_(user.nombre); }

/**
 * ¿La puede ver? La suya; o una compartida con ella ("Álvaro" o su nombre); o una compartida con
 * "*" (todo el equipo de empleados). Álvaro solo ve las suyas y las que un empleado le comparte.
 */
function _not_visible_(n, user) {
  if (_not_esMia_(n, user)) return true;
  var yo = _wa_norm_(user.nombre);
  return n.compartida.some(function(c) { return c === '*' ? user.rol !== 'admin' : _wa_norm_(c) === yo; });
}

/** Busca una nota que esa persona pueda ver; con soloMia=true, solo si la escribió ella. */
function _not_buscar_(id, user, soloMia) {
  id = String(id || '');
  if (!NOT_.ID_RE.test(id)) _wa_fail_('Esa nota no existe.');
  var n = _not_todas_().filter(function(x) { return x.id === id; })[0];
  if (!n || !_not_visible_(n, user) || (soloMia && !_not_esMia_(n, user))) _wa_fail_('Esa nota no existe.');
  return n;
}

function _not_ordena_(lista, filtro) {
  lista.sort(function(a, b) {
    if (a.pendiente !== b.pendiente) return a.pendiente ? -1 : 1;
    if (a.fecha !== b.fecha) return filtro === 'pendientes' ? (a.fecha < b.fecha ? -1 : 1) : (a.fecha < b.fecha ? 1 : -1);
    return a.id < b.id ? 1 : -1;
  });
}

// ---- Acciones ----

/** Notas que puede ver esta persona, con filtro (pendientes / todas / compartidas / hechas). */
function wa_notaLista_(user, p) {
  var filtro = NOT_.FILTROS.indexOf(String(p.filtro)) !== -1 ? String(p.filtro) : 'pendientes';
  var total = _not_todas_().filter(function(n) { return _not_visible_(n, user); });
  var lista = total.filter(function(n) {
    if (filtro === 'pendientes') return n.pendiente;
    if (filtro === 'hechas') return !n.pendiente;
    if (filtro === 'compartidas') return n.compartida.length > 0;
    return true;
  });
  _not_ordena_(lista, filtro);
  return {
    filtro: filtro, notas: lista.slice(0, 200).map(function(n) { return _not_pub_(n, user); }),
    cuenta: { pendientes: total.filter(function(n) { return n.pendiente; }).length, todas: total.length,
              compartidas: total.filter(function(n) { return n.compartida.length > 0; }).length,
              hechas: total.filter(function(n) { return !n.pendiente; }).length }
  };
}

/** Empleado: notas que Álvaro le ha compartido (las pendientes primero). */
function wa_miNotas_(user, p) {
  var lista = _not_todas_().filter(function(n) { return _not_visible_(n, user); });
  _not_ordena_(lista, 'pendientes');
  return { notas: lista.slice(0, 60).map(function(n) { return _not_pub_(n, user); }) };
}

function wa_notaGuardar_(user, p) {
  var fecha = _ag_fechaIso_(p.fecha);
  var titulo = _wa_s_(p.titulo, 100);
  if (titulo.trim().length < 2) _wa_fail_('Ponle un título a la nota.');
  var texto = _wa_s_(p.texto, 2000);
  var items = [];
  (Array.isArray(p.items) ? p.items : []).forEach(function(x) {
    var t = _wa_s_(x && x.t, 120).trim();
    if (t) items.push({ t: t, h: !!(x && x.h === true) });
  });
  if (items.length > NOT_.MAX_ITEMS) _wa_fail_('La lista puede tener como mucho ' + NOT_.MAX_ITEMS + ' puntos.');
  var activos = _wa_listaOperariosActivos_(), comp = [];
  (Array.isArray(p.compartida) ? p.compartida : []).forEach(function(c) {
    var ok = String(c) === '*' ? '*' : _wa_enActivos_(c, activos);
    if (!ok && user.rol !== 'admin' && _wa_norm_(c) === _wa_norm_('Álvaro')) ok = 'Álvaro';   // un empleado puede compartir con Álvaro
    if (!ok) _wa_fail_('Solo se puede compartir con el equipo, con Álvaro o con empleados de la lista.');
    if (_wa_norm_(ok) === _wa_norm_(user.nombre)) return;   // con uno mismo no hace falta
    if (comp.indexOf(ok) === -1) comp.push(ok);
  });
  var hecha = p.hecha === true;
  var hora = _wa_hora_(p.hora), aviso = '';
  if (p.aviso !== '' && p.aviso != null) {
    if (user.rol !== 'admin') _wa_fail_('Los avisos en el móvil solo los puede pedir Álvaro (se envían a su Google Calendar).');
    aviso = Number(p.aviso);
    if (AG_.AVISOS.indexOf(aviso) === -1) _wa_fail_('El aviso no es válido.');
    if (!hora) _wa_fail_('Para pedir un aviso indica la hora.');
    if (_ag_inicio_(fecha, hora).getTime() - aviso * 60000 <= Date.now()) _wa_fail_('Ese aviso ya ha pasado: elige otra hora o menos antelación.');
  }
  var esNueva = !p.id;
  return _wa_withLock_(function() {
    var sh = _not_hoja_(), C = NOT_.COL, antes = null;
    if (!esNueva) antes = _not_buscar_(p.id, user, true);   // solo quien la escribió la edita
    var id = esNueva ? nextId(NOT_.SHEET, 'NOT') : antes.id;
    var pendiente = items.length ? items.some(function(x) { return !x.h; }) : !hecha;
    var evento = '';
    if (antes && antes.evento) { _ag_borraEvento_(antes.evento); }
    if (aviso !== '' && pendiente) evento = _ag_crearEvento_(titulo, texto, fecha, hora, aviso, '📝');
    try {
      var d = _parseReportDate_(_wa_fechaDMY_(fecha)), ahora = new Date();
      var fila = [id, d, titulo, texto, JSON.stringify(items), comp.length ? '|' + comp.join('|') + '|' : '', hecha, '', aviso, evento,
                  antes ? '' : ahora, ahora, antes ? antes.autor : user.nombre];
      var row = esNueva ? sh.getLastRow() + 1 : antes.row;
      if (esNueva) sh.appendRow(fila);
      else {
        for (var c = 0; c < fila.length; c++) { if (c === C.CREADA - 1) continue; sh.getRange(row, c + 1).setValue(fila[c]); }
      }
      sh.getRange(row, C.FECHA).setNumberFormat('dd/MM/yyyy');
      var cel = sh.getRange(row, C.HORA); cel.setNumberFormat('@'); cel.setValue(hora);
    } catch (e) { _ag_borraEvento_(evento); throw e; }
    logEvent('INFO', 'wa_notaGuardar', id + (esNueva ? ' creada' : ' editada') + ' ' + fecha, titulo);
    return { id: id, calendario: !!evento };
  });
}

/** Marca/desmarca un punto de la lista (Álvaro o el empleado con quien está compartida). */
function wa_notaItem_(user, p) {
  var idx = Number(p.idx);
  return _wa_withLock_(function() {
    var n = _not_buscar_(p.id, user);
    if (!(idx >= 0 && idx < n.items.length && idx === Math.floor(idx))) _wa_fail_('Ese punto no existe.');
    n.items[idx].h = p.hecho === true;
    _not_hoja_().getRange(n.row, NOT_.COL.ITEMS).setValue(JSON.stringify(n.items));
    _not_hoja_().getRange(n.row, NOT_.COL.ACT).setValue(new Date());
    var pendiente = n.items.some(function(x) { return !x.h; });
    if (!pendiente && n.evento) { _ag_borraEvento_(n.evento); _not_hoja_().getRange(n.row, NOT_.COL.EVENTO).setValue(''); }   // todo hecho: ya no tiene que sonar
    return { ok: true, pendiente: pendiente, nHechos: n.items.filter(function(x) { return x.h; }).length };
  });
}

/** Marca/desmarca una nota sin lista como hecha. */
function wa_notaHecha_(user, p) {
  return _wa_withLock_(function() {
    var n = _not_buscar_(p.id, user);
    if (n.items.length) _wa_fail_('Esta nota tiene lista: marca sus puntos.');
    var hecha = p.hecho === true, sh = _not_hoja_();
    sh.getRange(n.row, NOT_.COL.HECHA).setValue(hecha);
    sh.getRange(n.row, NOT_.COL.ACT).setValue(new Date());
    if (hecha && n.evento) { _ag_borraEvento_(n.evento); sh.getRange(n.row, NOT_.COL.EVENTO).setValue(''); }
    return { ok: true };
  });
}

/** Mueve una nota pendiente (y su aviso de Calendar) a otra fecha, y opcionalmente a otra hora. */
function wa_notaMover_(user, p) {
  var m = _ag_nuevoMomento_(p);
  return _wa_withLock_(function() {
    var n = _not_buscar_(p.id, user, true);
    if (!n.pendiente) _wa_fail_('Esa nota ya está hecha.');
    var hora = m.hora === null ? n.hora : m.hora, sh = _not_hoja_(), C = NOT_.COL;
    var ev = n.evento ? _ag_reharAviso_(n.evento, n.aviso, n.titulo, n.texto, m.fecha, hora, '📝') : '';
    var cf = sh.getRange(n.row, C.FECHA); cf.setValue(_parseReportDate_(_wa_fechaDMY_(m.fecha))); cf.setNumberFormat('dd/MM/yyyy');
    var ch = sh.getRange(n.row, C.HORA); ch.setNumberFormat('@'); ch.setValue(hora);
    sh.getRange(n.row, C.EVENTO).setValue(ev);
    sh.getRange(n.row, C.ACT).setValue(new Date());
    logEvent('INFO', 'wa_notaMover', n.id + ' movida ' + n.fecha + ' → ' + m.fecha, n.titulo);
    return { ok: true, fecha: m.fecha, avisoPerdido: !!n.evento && !ev };
  });
}

function wa_notaBorrar_(user, p) {
  return _wa_withLock_(function() {
    var n = _not_buscar_(p.id, user, true);
    _ag_borraEvento_(n.evento);
    _not_hoja_().deleteRow(n.row);
    logEvent('INFO', 'wa_notaBorrar', n.id + ' borrada', n.titulo);
    return { ok: true };
  });
}

// =============================================================
// Libreta del día: un texto libre por día, que se guarda solo mientras se escribe
// y aparece en el calendario de ese día. Una fila por día en la hoja LIBRETA.
// =============================================================

var LIB_ = { SHEET: 'LIBRETA', HEAD: ['Fecha', 'Texto', 'Actualizado', 'Autor'], MAX: 20000 };

function _lib_hoja_() {
  var sh = getOrCreateSheet(LIB_.SHEET);
  if (sh.getLastRow() === 0) { sh.appendRow(LIB_.HEAD); applyHeaderStyle_(sh, LIB_.HEAD.length); return sh; }
  _ag_aseguraColumna_(sh, 4, 'Autor');
  return sh;
}

function _lib_filas_() {
  var sh = _lib_hoja_(), last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, LIB_.HEAD.length).getValues(), out = [];
  vals.forEach(function(r, i) {
    var f = _ag_iso_(r[0]);
    if (f) out.push({ row: i + 2, fecha: f, texto: String(r[1] || ''), act: r[2], autor: String(r[3] || '').trim() || 'Álvaro' });
  });
  return out;
}

function _lib_dia_(fecha, user) {
  var f = _lib_filas_().filter(function(x) { return x.fecha === fecha && _ag_esMio_(x.autor, user); });
  return f.length ? f[f.length - 1] : null;
}

function wa_libretaGet_(user, p) {
  var fecha = _ag_fechaIso_(p.fecha), x = _lib_dia_(fecha, user);
  return { fecha: fecha, texto: x ? x.texto : '', actualizado: x ? _maq_fecha_(x.act) : '' };
}

/** Libretas de días anteriores de esta persona (las más recientes primero), con un trozo del texto. */
function wa_libretaLista_(user, p) {
  var filas = _lib_filas_().filter(function(x) { return _ag_esMio_(x.autor, user) && x.texto.trim(); });
  filas.sort(function(a, b) { return a.fecha < b.fecha ? 1 : (a.fecha > b.fecha ? -1 : 0); });
  return {
    dias: filas.slice(0, 30).map(function(x) {
      var t = x.texto.trim();
      return { fecha: x.fecha, resumen: t.length > 160 ? t.slice(0, 160) + '…' : t, largo: t.length };
    })
  };
}

/** Guarda el texto completo del día (lo llama la app cada pocos segundos mientras se escribe). Texto vacío = borra el día. */
function wa_libretaGuardar_(user, p) {
  var fecha = _ag_fechaIso_(p.fecha);
  var texto = String(p.texto == null ? '' : p.texto).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  if (texto.length > LIB_.MAX) _wa_fail_('La libreta de un día admite como mucho ' + LIB_.MAX + ' letras. Pasa lo que sobre a otra nota.');
  var vacio = !texto.trim();
  return _wa_withLock_(function() {
    var sh = _lib_hoja_(), x = _lib_dia_(fecha, user);
    if (vacio) { if (x) sh.deleteRow(x.row); return { ok: true, guardado: '' }; }
    var row = x ? x.row : sh.getLastRow() + 1, ahora = new Date();
    var cT = sh.getRange(row, 2); cT.setNumberFormat('@'); cT.setValue(texto);   // formato texto: nunca se interpreta como fórmula
    var cF = sh.getRange(row, 1); cF.setNumberFormat('dd/MM/yyyy'); cF.setValue(_parseReportDate_(_wa_fechaDMY_(fecha)));
    sh.getRange(row, 3).setValue(ahora);
    sh.getRange(row, 4).setValue(user.nombre);
    return { ok: true, guardado: _maq_fecha_(ahora) };
  });
}
