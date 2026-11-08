// =============================================================
// Agenda.gs — Calendario de Álvaro: trabajos del día, tareas y recordatorios
// =============================================================
//
// - Los TRABAJOS se leen de la hoja TRABAJOS (no se duplican aquí).
// - Las TAREAS y RECORDATORIOS se guardan en la hoja AGENDA (se crea sola).
// - Si Álvaro pide un aviso, se crea un evento en SU Google Calendar con un
//   recordatorio emergente: así la notificación le llega al móvil aunque la
//   app esté cerrada (la web app no puede avisar por sí sola).
//
// Permiso necesario (una sola vez): scope calendar. Se concede ejecutando
// "webappAutorizarCalendario" desde el editor con la cuenta de Álvaro.
// =============================================================

var AG_ = {
  SHEET: 'AGENDA',
  HEAD: ['ID_Item', 'Fecha', 'Hora', 'Titulo', 'Detalle', 'Hecho', 'Aviso_Min', 'ID_Evento', 'Creado'],
  COL: { ID: 1, FECHA: 2, HORA: 3, TITULO: 4, DETALLE: 5, HECHO: 6, AVISO: 7, EVENTO: 8, CREADO: 9 },
  AVISOS: [0, 5, 10, 15, 30, 60, 120, 1440],   // minutos antes
  DURACION_MIN: 30,
  ID_RE: /^AGE-\d{3,6}$/
};

function _ag_hoja_() {
  var sh = getOrCreateSheet(AG_.SHEET);
  if (sh.getLastRow() === 0) { sh.appendRow(AG_.HEAD); applyHeaderStyle_(sh, AG_.HEAD.length); }
  return sh;
}

function _ag_iso_(d) {
  return (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d.getTime()))
    ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
}

/** Fecha y hora de ahora (hora local del script), para saber qué está atrasado. */
function _ag_ahora_() {
  var d = new Date(), tz = Session.getScriptTimeZone();
  return { hoy: Utilities.formatDate(d, tz, 'yyyy-MM-dd'), hora: Utilities.formatDate(d, tz, 'HH:mm') };
}

/** Atrasado = su día ya pasó, o es hoy y su hora prevista ya pasó. Sin hora prevista, hasta que acabe el día. */
function _ag_vencido_(fecha, hora, ahora) {
  if (!fecha) return false;
  if (fecha < ahora.hoy) return true;
  return fecha === ahora.hoy && !!hora && hora < ahora.hora;
}

/** Un trabajo sin completar está atrasado a su hora de fin (o a la de inicio si aún no ha empezado y no tiene fin). */
function _ag_trabajoAtrasado_(t, ahora) {
  if (t.estado !== 'Pendiente' && t.estado !== 'En curso') return false;
  var hFin = _wa_horaTxt_(t.hFin), hIni = _wa_horaTxt_(t.hIni);
  return _ag_vencido_(_ag_iso_(t.fecha), hFin || (t.estado === 'Pendiente' ? hIni : ''), ahora);
}

function _ag_todos_() {
  var sh = _ag_hoja_(), last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, AG_.HEAD.length).getValues(), C = AG_.COL, out = [];
  vals.forEach(function(r, i) {
    var id = String(r[C.ID - 1] || '').trim();
    if (!id) return;
    var aviso = r[C.AVISO - 1];
    out.push({
      row: i + 2, id: id, fecha: _ag_iso_(r[C.FECHA - 1]), hora: _wa_horaTxt_(r[C.HORA - 1]),
      titulo: String(r[C.TITULO - 1] || ''), detalle: String(r[C.DETALLE - 1] || ''),
      hecho: _wa_bool_(r[C.HECHO - 1]),
      aviso: (aviso === '' || aviso == null || isNaN(Number(aviso))) ? '' : Number(aviso),
      evento: String(r[C.EVENTO - 1] || '').trim()
    });
  });
  return out;
}

function _ag_pub_(x) {
  return { id: x.id, fecha: x.fecha, hora: x.hora, titulo: x.titulo, detalle: x.detalle, hecho: x.hecho, aviso: x.aviso, calendario: !!x.evento,
           atrasado: !x.hecho && _ag_vencido_(x.fecha, x.hora, _ag_ahora_()) };
}

function _ag_fechaIso_(v) {
  var s = String(v || '').trim();
  _wa_fechaDMY_(s);   // valida que exista (31/02 no vale)
  return s;
}

// ---- Google Calendar ----

function _ag_inicio_(fechaISO, hora) {
  var p = fechaISO.split('-'), hm = hora.split(':');
  return new Date(+p[0], +p[1] - 1, +p[2], +hm[0], +hm[1], 0);
}

function _ag_crearEvento_(titulo, detalle, fechaISO, hora, aviso, icono) {
  var ini = _ag_inicio_(fechaISO, hora), fin = new Date(ini.getTime() + AG_.DURACION_MIN * 60000);
  var ev = CalendarApp.getDefaultCalendar().createEvent((icono || '⏰') + ' ' + titulo, ini, fin,
    { description: (detalle ? detalle + '\n\n' : '') + 'Creado desde Álvaro Gestión.' });
  ev.removeAllReminders();
  ev.addPopupReminder(aviso);
  return ev.getId();
}

function _ag_borraEvento_(idEvento) {
  if (!idEvento) return;
  try {
    var ev = CalendarApp.getDefaultCalendar().getEventById(idEvento);
    if (ev) ev.deleteEvent();
  } catch (e) { logEvent('WARN', '_ag_borraEvento_', e.message, idEvento); }
}

// ---- Acciones (solo Álvaro) ----

/** Resumen de un mes: por cada día, nº de trabajos y de tareas/recordatorios pendientes. */
function wa_agMes_(user, p) {
  var mes = String(p.mes || '');
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(mes)) _wa_fail_('El mes no es válido.');
  var dias = {};
  var ahora = _ag_ahora_();
  function dia(k) { return dias[k] || (dias[k] = { t: 0, a: 0, h: 0, n: 0, o: 0 }); }
  _wa_trabajos_().forEach(function(t) {
    var k = _ag_iso_(t.fecha);
    if (k && k.indexOf(mes) === 0) { dia(k).t++; if (_ag_trabajoAtrasado_(t, ahora)) dia(k).o++; }
  });
  _ag_todos_().forEach(function(x) {
    if (x.fecha && x.fecha.indexOf(mes) === 0) {
      if (x.hecho) dia(x.fecha).h++; else { dia(x.fecha).a++; if (_ag_vencido_(x.fecha, x.hora, ahora)) dia(x.fecha).o++; }
    }
  });
  _not_todas_().forEach(function(n) {
    if (n.fecha && n.fecha.indexOf(mes) === 0 && n.pendiente) { dia(n.fecha).n++; if (n.atrasado) dia(n.fecha).o++; }
  });
  return { mes: mes, dias: dias };
}

/** Todo lo de un día: trabajos y tareas/recordatorios. */
function wa_agDia_(user, p) {
  var fecha = _ag_fechaIso_(p.fecha);
  var trabajos = [], ahora = _ag_ahora_();
  _wa_trabajos_().forEach(function(t) {
    if (_ag_iso_(t.fecha) !== fecha) return;
    trabajos.push({
      id: t.id, cliente: t.cliente, operario: t.operario, op2: t.op2,
      hIni: _wa_horaTxt_(t.hIni), hFin: _wa_horaTxt_(t.hFin), trabajo: t.trabajo, estado: t.estado,
      atrasado: _ag_trabajoAtrasado_(t, ahora)
    });
  });
  trabajos.sort(function(a, b) { return (a.hIni || '99:99') < (b.hIni || '99:99') ? -1 : 1; });
  var items = _ag_todos_().filter(function(x) { return x.fecha === fecha; }).map(_ag_pub_);
  items.sort(function(a, b) {
    if (a.hecho !== b.hecho) return a.hecho ? 1 : -1;
    return (a.hora || '99:99') < (b.hora || '99:99') ? -1 : ((a.hora || '99:99') > (b.hora || '99:99') ? 1 : 0);
  });
  var notas = _not_todas_().filter(function(n) { return n.fecha === fecha; });
  notas.sort(function(a, b) { return a.pendiente !== b.pendiente ? (a.pendiente ? -1 : 1) : (a.id < b.id ? -1 : 1); });
  return { fecha: fecha, trabajos: trabajos, items: items, notas: notas.map(_not_pub_) };
}

/** Todo lo que está sin hacer (trabajos, tareas y notas), con su marca de atrasado. */
function _ag_pendientesTodos_() {
  var ahora = _ag_ahora_(), out = [];
  _wa_trabajos_().forEach(function(t) {
    if (t.estado !== 'Pendiente' && t.estado !== 'En curso') return;
    var f = _ag_iso_(t.fecha); if (!f) return;
    var hIni = _wa_horaTxt_(t.hIni), hFin = _wa_horaTxt_(t.hFin);
    out.push({ tipo: 'trabajo', id: t.id, titulo: t.cliente, detalle: [t.trabajo, [t.operario, t.op2].filter(Boolean).join(' y ')].filter(Boolean).join(' · '),
               fecha: f, hora: hIni ? hIni + (hFin ? '–' + hFin : '') : '', orden: hIni || hFin || '', atrasado: _ag_trabajoAtrasado_(t, ahora) });
  });
  _ag_todos_().forEach(function(x) {
    if (x.hecho || !x.fecha) return;
    out.push({ tipo: 'tarea', id: x.id, titulo: x.titulo, detalle: x.detalle, fecha: x.fecha, hora: x.hora, orden: x.hora, atrasado: _ag_vencido_(x.fecha, x.hora, ahora) });
  });
  _not_todas_().forEach(function(n) {
    if (!n.pendiente || !n.fecha) return;
    var pend = n.items.filter(function(i) { return !i.h; }).length;
    out.push({ tipo: 'nota', id: n.id, titulo: n.titulo, detalle: n.items.length ? pend + ' de ' + n.items.length + ' puntos por hacer' : n.texto, fecha: n.fecha, hora: n.hora, orden: n.hora, atrasado: n.atrasado });
  });
  return { ahora: ahora, items: out };
}

function _ag_cuentaAtrasados_() {
  return _ag_pendientesTodos_().items.filter(function(x) { return x.atrasado; }).length;
}

/** Pendientes de Álvaro: atrasado (rojo), para hoy y próximos 7 días. */
function wa_agPendientes_(user, p) {
  var r = _ag_pendientesTodos_(), limite = _ag_iso_(new Date(Date.now() + 7 * 86400000));
  function porFecha(a, b) { return a.fecha !== b.fecha ? (a.fecha < b.fecha ? -1 : 1) : ((a.orden || '99:99') < (b.orden || '99:99') ? -1 : ((a.orden || '99:99') > (b.orden || '99:99') ? 1 : 0)); }
  var atrasados = [], hoy = [], proximos = [];
  r.items.forEach(function(x) {
    if (x.atrasado) atrasados.push(x);
    else if (x.fecha === r.ahora.hoy) hoy.push(x);
    else if (x.fecha > r.ahora.hoy && x.fecha <= limite) proximos.push(x);
  });
  atrasados.sort(porFecha); hoy.sort(porFecha); proximos.sort(porFecha);
  return { hoy: r.ahora.hoy, atrasados: atrasados, paraHoy: hoy, proximos: proximos };
}

function wa_agCrear_(user, p) {
  var fecha = _ag_fechaIso_(p.fecha);
  var titulo = _wa_s_(p.titulo, 120);
  if (titulo.trim().length < 2) _wa_fail_('Escribe qué hay que hacer o recordar.');
  var detalle = _wa_s_(p.detalle, 500);
  var hora = _wa_hora_(p.hora);
  var aviso = '';
  if (p.aviso !== '' && p.aviso != null) {
    aviso = Number(p.aviso);
    if (AG_.AVISOS.indexOf(aviso) === -1) _wa_fail_('El aviso no es válido.');
    if (!hora) _wa_fail_('Para pedir un aviso indica la hora.');
    if (_ag_inicio_(fecha, hora).getTime() - aviso * 60000 <= Date.now()) _wa_fail_('Ese aviso ya ha pasado: elige otra hora o menos antelación.');
  }
  return _wa_withLock_(function() {
    var sh = _ag_hoja_(), C = AG_.COL, id = nextId(AG_.SHEET, 'AGE');
    var evento = aviso !== '' ? _ag_crearEvento_(titulo, detalle, fecha, hora, aviso) : '';
    try {
      var f = _wa_fechaDMY_(fecha), d = _parseReportDate_(f);
      var fila = new Array(AG_.HEAD.length).fill('');
      fila[C.ID - 1] = id; fila[C.FECHA - 1] = d; fila[C.HORA - 1] = ''; fila[C.TITULO - 1] = titulo; fila[C.DETALLE - 1] = detalle;
      fila[C.HECHO - 1] = false; fila[C.AVISO - 1] = aviso; fila[C.EVENTO - 1] = evento; fila[C.CREADO - 1] = new Date();
      sh.appendRow(fila);
      var row = sh.getLastRow();
      sh.getRange(row, C.FECHA).setNumberFormat('dd/MM/yyyy');
      if (hora) { var c = sh.getRange(row, C.HORA); c.setNumberFormat('@'); c.setValue(hora); }
    } catch (e) { _ag_borraEvento_(evento); throw e; }   // sin fila no debe quedar un aviso suelto
    logEvent('INFO', 'wa_agCrear', id + ' ' + fecha + (hora ? ' ' + hora : ''), titulo);
    return { id: id, calendario: !!evento };
  });
}

function wa_agHecho_(user, p) {
  var id = String(p.id || '');
  if (!AG_.ID_RE.test(id)) _wa_fail_('Esa tarea no existe.');
  var hecho = p.hecho === true;
  return _wa_withLock_(function() {
    var x = _ag_todos_().filter(function(y) { return y.id === id; })[0];
    if (!x) _wa_fail_('Esa tarea no existe.');
    var sh = _ag_hoja_();
    sh.getRange(x.row, AG_.COL.HECHO).setValue(hecho);
    if (hecho && x.evento) { _ag_borraEvento_(x.evento); sh.getRange(x.row, AG_.COL.EVENTO).setValue(''); }   // hecha: ya no tiene que sonar
    return { ok: true };
  });
}

function wa_agBorrar_(user, p) {
  var id = String(p.id || '');
  if (!AG_.ID_RE.test(id)) _wa_fail_('Esa tarea no existe.');
  return _wa_withLock_(function() {
    var x = _ag_todos_().filter(function(y) { return y.id === id; })[0];
    if (!x) _wa_fail_('Esa tarea no existe.');
    _ag_borraEvento_(x.evento);
    _ag_hoja_().deleteRow(x.row);
    logEvent('INFO', 'wa_agBorrar', id + ' borrada', x.titulo);
    return { ok: true };
  });
}

// ---- Permiso de Google Calendar ----

/**
 * Concede UNA VEZ el permiso de Google Calendar a la web app. Ejecutar desde el
 * editor de Apps Script con la cuenta de Álvaro y aceptar la ventana de permisos.
 * Hace lo mismo que la app (crear un evento con aviso) y lo borra al instante:
 * así, si funciona aquí, funcionará al guardar un recordatorio.
 */
function webappAutorizarCalendario() {
  var cal = CalendarApp.getDefaultCalendar();
  var ini = new Date(2000, 0, 1, 10, 0), ev = cal.createEvent('Prueba de permiso (se borra sola)', ini, new Date(ini.getTime() + 60000));
  ev.removeAllReminders(); ev.addPopupReminder(10);
  ev.deleteEvent();
  Logger.log('Permiso de Calendar CORRECTO (calendario "' + cal.getName() + '"). Ya se pueden crear recordatorios con aviso.');
}
