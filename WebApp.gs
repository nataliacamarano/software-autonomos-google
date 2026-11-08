// ============================================================
// WebApp.gs — Web app móvil: vista de Álvaro (admin) y vista de empleado
// ============================================================
//
// Una sola web app con dos interfaces. El acceso se decide por un enlace
// personal (?t=TOKEN):
//   - Token de admin   → CONFIG, clave "webapp_admin_token"
//   - Token de empleado → hoja OPERARIOS, columna 10 "Token_Acceso"
//
// Estados de empleado:
//   Activo     → Activo = TRUE  y Bloqueado vacío/FALSE
//   Bloqueado  → Activo = TRUE  y Bloqueado = TRUE (no puede entrar a la app)
//   De baja    → Activo = FALSE (no aparece en listas ni puede entrar)
// Los clientes solo tienen Activo / De baja.
//
// Nunca se borra nada. Todas las altas reutilizan las funciones existentes
// (createCliente, createOperario, createTrabajo) para escribir igual que los
// formularios de Google.
//
// Solo wa_api() y doGet() son llamables desde el navegador; el resto de
// funciones terminan en "_" y son privadas.
// ============================================================

var WA_ = {
  CFG_ADMIN_TOKEN: 'webapp_admin_token',
  COL_BLOQ:        9,
  COL_TOKEN:       10,
  MAX_FOTOS:       5,
  MAX_FOTO_BYTES:  4 * 1024 * 1024,
  TOKEN_RE:        /^[A-Za-z0-9]{16,128}$/,
  URL_RE:          /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/,
  URL_BASE:        'https://script.google.com/macros/s/AKfycbzOgZdq7pV6OIXRRdo5kysBSlOoNETk1HRFEvzsxD7jVkszMCOo_XsVrZiXGoJMEX17Cg/exec',
  ESTADOS_PARTE:   ['Pendiente', 'En curso', 'Completado']
};

// ---- Entrada web ----

function doGet(e) {
  var t = (e && e.parameter && e.parameter.t) ? String(e.parameter.t) : '';
  if (!WA_.TOKEN_RE.test(t)) t = '';
  var m = (e && e.parameter && e.parameter.m) ? String(e.parameter.m) : '';   // código de máquina (QR)
  if (!MAQ_.ID_RE.test(m)) m = '';
  var tpl = HtmlService.createTemplateFromFile('WebAppUI');
  tpl.tokenJson = JSON.stringify(t);
  tpl.maqJson = JSON.stringify(m);
  return tpl.evaluate()
    .setTitle('Álvaro Gestión')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---- API única (llamada desde la interfaz con google.script.run) ----

function wa_api(token, action, payload) {
  try {
    var user = _wa_auth_(token);
    if (!user) return { ok: false, code: 'AUTH', error: 'Este enlace no es válido.' };
    if (user.bloqueado) return { ok: false, code: 'BLOQ', error: 'Tu acceso está desactivado. Habla con Álvaro.' };

    var def = _wa_actions_()[String(action)];
    if (!def) return { ok: false, error: 'Acción desconocida.' };
    if (def.roles.indexOf(user.rol) === -1) {
      logEvent('WARN', 'wa_api', 'Acción sin permiso: ' + action, user.rol + ' ' + (user.nombre || ''));
      return { ok: false, code: 'PERM', error: 'No tienes permiso para esto.' };
    }

    payload = payload || {};

    // Evita duplicados si el móvil reintenta el envío (misma petición = mismo resultado)
    var cache = null, cacheKey = '';
    if (def.write && payload.reqId && /^[A-Za-z0-9_-]{8,64}$/.test(String(payload.reqId))) {
      cache = CacheService.getScriptCache();
      cacheKey = 'wa_req_' + payload.reqId;
      var prev = cache.get(cacheKey);
      if (prev) return JSON.parse(prev);
    }

    var data = def.fn(user, payload);
    var out = { ok: true, data: data };
    if (cache) { try { cache.put(cacheKey, JSON.stringify(out), 21600); } catch (_) {} }
    return out;
  } catch (e) {
    var msg = String(e && e.message || e);
    if (msg.indexOf('USR:') === 0) return { ok: false, error: msg.substring(4) };
    logEvent('ERROR', 'wa_api', String(action) + ': ' + msg, e && e.stack);
    if (msg.indexOf('No tienes permiso para llamar a') !== -1) {
      return { ok: false, error: 'Falta conceder un permiso de Google a la aplicación. Avisa a quien la administra (no es problema de lo que has escrito).' };
    }
    return { ok: false, error: 'No se pudo completar. Inténtalo otra vez.' };
  }
}

function _wa_actions_() {
  var A = ['admin'], E = ['empleado'], AE = ['admin', 'empleado'];
  return {
    init:             { roles: AE, fn: _wa_init_ },
    resumen:          { roles: A,  fn: _wa_resumenAdmin_ },
    crearTrabajo:     { roles: A,  fn: _wa_crearTrabajo_,  write: true },
    crearCliente:     { roles: A,  fn: _wa_crearCliente_,  write: true },
    crearEmpleado:    { roles: A,  fn: _wa_crearEmpleado_, write: true },
    crearGasto:       { roles: A,  fn: _wa_crearGasto_,    write: true },
    pendientes:       { roles: A,  fn: _wa_pendientes_ },
    horasEquipo:      { roles: A,  fn: _wa_horasEquipo_ },
    listaPersonas:    { roles: A,  fn: _wa_listaPersonas_ },
    accionPersona:    { roles: A,  fn: _wa_accionPersona_, write: true },
    editarPersona:    { roles: A,  fn: _wa_editarPersona_, write: true },
    enlaceEmpleado:   { roles: A,  fn: _wa_enlaceEmpleado_ },
    regenerarEnlace:  { roles: A,  fn: _wa_regenerarEnlace_, write: true },
    maqLista:         { roles: AE, fn: wa_maqLista_ },
    maqFicha:         { roles: AE, fn: wa_maqFicha_ },
    maqUso:           { roles: AE, fn: wa_maqUso_,    write: true },
    maqLeerEnlace:    { roles: A,  fn: wa_maqLeerEnlace_ },
    maqCrear:         { roles: A,  fn: wa_maqCrear_,  write: true },
    maqEditar:        { roles: A,  fn: wa_maqEditar_, write: true },
    maqEstado:        { roles: A,  fn: wa_maqEstado_, write: true },
    maqReparacion:    { roles: A,  fn: wa_maqReparacion_, write: true },
    maqReparacionBorrar: { roles: A, fn: wa_maqReparacionBorrar_, write: true },
    maqBaja:          { roles: A,  fn: wa_maqBaja_,   write: true },
    agMes:            { roles: A,  fn: wa_agMes_ },
    agDia:            { roles: A,  fn: wa_agDia_ },
    agCrear:          { roles: A,  fn: wa_agCrear_, write: true },
    agHecho:          { roles: A,  fn: wa_agHecho_, write: true },
    agBorrar:         { roles: A,  fn: wa_agBorrar_, write: true },
    notaLista:        { roles: A,  fn: wa_notaLista_ },
    notaGuardar:      { roles: A,  fn: wa_notaGuardar_, write: true },
    notaItem:         { roles: AE, fn: wa_notaItem_, write: true },
    notaHecha:        { roles: AE, fn: wa_notaHecha_, write: true },
    notaBorrar:       { roles: A,  fn: wa_notaBorrar_, write: true },
    miNotas:          { roles: E,  fn: wa_miNotas_ },
    miInicio:         { roles: E,  fn: _wa_miInicio_ },
    enviarParte:      { roles: E,  fn: _wa_enviarParte_,   write: true }
  };
}

// ---- Autenticación ----

function _wa_newToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function _wa_bool_(v) {
  return v === true || String(v).toLowerCase() === 'true';
}

/**
 * Devuelve { rol:'admin' } | { rol:'empleado', ... } | { bloqueado:true } | null.
 */
function _wa_auth_(token) {
  token = String(token || '');
  if (!WA_.TOKEN_RE.test(token)) return null;

  var adminTok = getConfig(WA_.CFG_ADMIN_TOKEN);
  if (adminTok && token === adminTok) return { rol: 'admin', nombre: 'Álvaro' };

  var ope = _wa_operarios_();
  for (var i = 0; i < ope.length; i++) {
    if (ope[i].token && ope[i].token === token) {
      if (!ope[i].activo || ope[i].bloq) return { bloqueado: true };
      return { rol: 'empleado', id: ope[i].id, nombre: ope[i].label };
    }
  }
  return null;
}

function _wa_withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { try { lock.releaseLock(); } catch (_) {} }
}

// ---- Validación y utilidades ----

function _wa_fail_(msg) { throw new Error('USR:' + msg); }

/** Limpia un texto: sin control chars, longitud máxima, y a prueba de fórmulas de Sheets. */
function _wa_s_(v, max) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim();
  if (max && s.length > max) s = s.substring(0, max);
  if (/^[=@]/.test(s) || /^[+\-]\s*[A-Za-z(]/.test(s)) s = "'" + s;
  return s;
}

function _wa_num_(v, max) {
  var n = parseFloat(String(v == null ? '' : v).replace(',', '.'));
  if (isNaN(n) || n < 0) return 0;
  return Math.min(n, max || 1000000);
}

/**
 * Valida y normaliza un teléfono de cualquier país.
 * - Admite "+prefijo" o "00prefijo" (00 se convierte en +); entre 6 y 15 cifras (E.164).
 * - Sin prefijo se guarda tal cual (la app asume España solo al enviar por WhatsApp).
 * - Si empieza por "+", "0", "-" o son solo cifras, se guarda como TEXTO (apóstrofo inicial,
 *   que Sheets no muestra) para que no lo convierta en fórmula ni pierda ceros.
 */
function _wa_tel_(v) {
  var bruto = String(v == null ? '' : v).trim();
  if (!bruto) _wa_fail_('Escribe un teléfono.');
  var s = bruto.replace(/[^\d+()\-.\s]/g, '').replace(/\s+/g, ' ').trim();
  if (s.replace(/^\(\s*/, '').indexOf('+', 1) !== -1) _wa_fail_('El teléfono no es válido: el "+" solo puede ir al principio.');
  if (s.slice(0, 2) === '00') s = '+' + s.slice(2).trim();
  var digitos = s.replace(/\D/g, '');
  if (digitos.length < 6 || digitos.length > 15) {
    _wa_fail_('El teléfono no es válido. Si es de otro país, escribe el prefijo (por ejemplo +44 7700 900123).');
  }
  if (/^[+0\-(]/.test(s) || /^\d+$/.test(s)) s = "'" + s;
  return s;
}

function _wa_hora_(v, etiqueta) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var m = s.match(/^([01]?\d|2[0-3]):([0-5]\d)$/);
  if (!m) _wa_fail_('La hora "' + s + '" no es válida (usa formato 09:30).');
  return ('0' + m[1]).slice(-2) + ':' + m[2];
}

function _wa_minutos_(hhmm) {
  var p = String(hhmm).split(':');
  return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
}

/** 'YYYY-MM-DD' → 'DD/MM/YYYY' (formato que espera createTrabajo). */
function _wa_fechaDMY_(iso) {
  var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) _wa_fail_('La fecha no es válida.');
  var y = +m[1], mo = +m[2], d = +m[3];
  var dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) _wa_fail_('La fecha no es válida.');
  return m[3] + '/' + m[2] + '/' + m[1];
}

function _wa_hoyDMY_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function _wa_hoyISO_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _wa_label_(nombre, apellidos, empresa) {
  var l = (String(nombre || '').trim() + ' ' + String(apellidos || '').trim()).trim();
  if (empresa) l += ' (' + String(empresa).trim() + ')';
  return l;
}

function _wa_norm_(s) { return _normalizeClientLabel_(s); }

function _wa_emailOk_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(s); }

function _wa_redondea_(h) { return Math.round(h * 10) / 10; }

function _wa_horaTxt_(raw) {
  var m = _parseHHMM_(raw);
  if (m === null) return '';
  return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2);
}

function _wa_fechaTxt_(d) {
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM') : '';
}

// ---- Lectura de hojas ----

function _wa_clientes_() {
  var sheet = getSheet(SHEET_NAMES.CLIENTES);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var ncol = Math.min(COL_CLI.ACTIVO, sheet.getMaxColumns());
  var vals = sheet.getRange(2, 1, last - 1, ncol).getValues();
  var out = [];
  vals.forEach(function(r, i) {
    var id = String(r[COL_CLI.ID - 1] || '').trim();
    if (!id) return;
    out.push({
      row: i + 2, id: id,
      nombre: String(r[COL_CLI.NOMBRE - 1] || '').trim(),
      apellidos: String(r[COL_CLI.APELLIDOS - 1] || '').trim(),
      empresa: String(r[COL_CLI.EMPRESA - 1] || '').trim(),
      email: String(r[COL_CLI.EMAIL - 1] || '').trim(),
      tel: String(r[COL_CLI.TELEFONO - 1] || '').trim(),
      dir: String(r[COL_CLI.DIRECCION - 1] || '').trim(),
      cif: String(r[COL_CLI.CIF_NIF - 1] || '').trim(),
      activo: _wa_bool_(r[COL_CLI.ACTIVO - 1]),
      label: _wa_label_(r[COL_CLI.NOMBRE - 1], r[COL_CLI.APELLIDOS - 1], r[COL_CLI.EMPRESA - 1])
    });
  });
  return out;
}

function _wa_operarios_() {
  var sheet = getSheet(SHEET_NAMES.OPERARIOS);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var ncol = Math.min(WA_.COL_TOKEN, sheet.getMaxColumns());
  var vals = sheet.getRange(2, 1, last - 1, ncol).getValues();
  var out = [];
  vals.forEach(function(r, i) {
    var id = String(r[COL_OPE.ID - 1] || '').trim();
    if (!id) return;
    out.push({
      row: i + 2, id: id,
      nombre: String(r[COL_OPE.NOMBRE - 1] || '').trim(),
      apellidos: String(r[COL_OPE.APELLIDOS - 1] || '').trim(),
      email: String(r[COL_OPE.EMAIL - 1] || '').trim(),
      tel: String(r[COL_OPE.TELEFONO - 1] || '').trim(),
      activo: _wa_bool_(r[COL_OPE.ACTIVO - 1]),
      bloq: ncol >= WA_.COL_BLOQ ? _wa_bool_(r[WA_.COL_BLOQ - 1]) : false,
      token: ncol >= WA_.COL_TOKEN ? String(r[WA_.COL_TOKEN - 1] || '').trim() : '',
      label: _wa_label_(r[COL_OPE.NOMBRE - 1], r[COL_OPE.APELLIDOS - 1])
    });
  });
  return out;
}

/** Añade (si faltan) las columnas Bloqueado y Token_Acceso a OPERARIOS. No toca datos. */
function _wa_ensureOperarioCols_() {
  var sheet = getSheet(SHEET_NAMES.OPERARIOS);
  if (sheet.getMaxColumns() < WA_.COL_TOKEN) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), WA_.COL_TOKEN - sheet.getMaxColumns());
  }
  var defs = [[WA_.COL_BLOQ, 'Bloqueado'], [WA_.COL_TOKEN, 'Token_Acceso']];
  defs.forEach(function(d) {
    var cell = sheet.getRange(1, d[0]);
    var cur = String(cell.getValue() || '').trim();
    if (!cur) cell.setValue(d[1]);
    else if (cur !== d[1]) throw new Error('OPERARIOS: la columna ' + d[0] + ' ya se usa para "' + cur + '".');
  });
  return sheet;
}

function _wa_trabajos_() {
  var sheet = getSheet('TRABAJOS');
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var ncol = Math.min(MAQ_.COL_TRA, sheet.getMaxColumns());   // incluye la columna "Maquinaria" si existe
  var vals = sheet.getRange(2, 1, last - 1, ncol).getValues();
  var out = [];
  vals.forEach(function(r, i) {
    var id = String(r[COL_TRA.ID - 1] || '').trim();
    if (!id) return;
    out.push({
      row: i + 2, id: id,
      maq: ncol >= MAQ_.COL_TRA ? String(r[MAQ_.COL_TRA - 1] || '').trim() : '',
      cliente: String(r[COL_TRA.CLIENTE - 1] || '').trim(),
      operario: String(r[COL_TRA.OPERARIO - 1] || '').trim(),
      fecha: _parseTrabajosDate_(r[COL_TRA.FECHA_TRABAJO - 1]),
      hIni: r[COL_TRA.HORA_INICIO - 1], hFin: r[COL_TRA.HORA_FIN - 1],
      trabajo: String(r[COL_TRA.TRABAJO_REALIZADO - 1] || '').trim(),
      incid: String(r[COL_TRA.INCIDENCIAS - 1] || '').trim(),
      estado: String(r[COL_TRA.ESTADO - 1] || '').trim(),
      op2: String(r[COL_TRA.OPERARIO_2 - 1] || '').trim(),
      hIni2: r[COL_TRA.HORA_INICIO_2 - 1], hFin2: r[COL_TRA.HORA_FIN_2 - 1]
    });
  });
  return out;
}

/** Horas que un operario (por etiqueta) aporta a un trabajo: como 1º o como 2º. */
function _wa_horasDe_(t, labelNorm) {
  var h = 0;
  if (_wa_norm_(t.operario) === labelNorm) h += _calcHoras_(t.hIni, t.hFin);
  if (t.op2 && _wa_norm_(t.op2) === labelNorm) h += _calcHoras_(t.hIni2, t.hFin2);
  return h;
}

function _wa_esDelMes_(fecha, y, m) {
  return fecha && fecha.getFullYear() === y && fecha.getMonth() === m;
}

function _wa_listaClientesActivos_() { return _getTrabajosDropdownClientes_(); }
function _wa_listaOperariosActivos_() { return _getTrabajosDropdownOperarios_(); }

function _wa_enActivos_(label, lista) {
  var n = _wa_norm_(label);
  for (var i = 0; i < lista.length; i++) if (_wa_norm_(lista[i]) === n) return lista[i];
  return null;
}

// ---- Facturas pendientes de cobro (solo lectura) ----

function _wa_facturasPendientes_() {
  var res = { total: 0, n: 0, lista: [] };
  try {
    var sheet = getSheet('FACTURAS');
    var last = sheet.getLastRow();
    if (last < 2) return res;
    var ncol = sheet.getLastColumn();
    var vals = sheet.getRange(1, 1, last, ncol).getValues();
    var hd = vals[0].map(function(h) { return String(h).toLowerCase().replace(/[\s_]/g, ''); });
    var ix = function(n) { return hd.indexOf(n); };
    var cId = ix('idfactura'), cCli = ix('nombrecliente'), cTot = ix('total'), cEst = ix('estado'),
        cPago = ix('fechapago'), cCobro = ix('estadocobro'), cFec = ix('fechaemision');
    if (cId < 0 || cTot < 0) return res;
    for (var i = 1; i < vals.length; i++) {
      var r = vals[i];
      if (!String(r[cId] || '').trim()) continue;
      var est = cEst >= 0 ? String(r[cEst] || '').toLowerCase() : '';
      var cob = cCobro >= 0 ? String(r[cCobro] || '').toLowerCase() : '';
      var pagoOk = cPago >= 0 && String(r[cPago] || '').trim() !== '';
      var cerrada = /pagad|cobrad|anulad/.test(est + ' ' + cob);
      if (pagoOk || cerrada) continue;
      var tot = Number(r[cTot]) || 0;
      res.total += tot; res.n++;
      if (res.lista.length < 50) {
        var f = cFec >= 0 && r[cFec] instanceof Date ? _wa_fechaTxt_(r[cFec]) : '';
        res.lista.push({ id: String(r[cId]), cliente: cCli >= 0 ? String(r[cCli] || '') : '', total: Math.round(tot * 100) / 100, fecha: f });
      }
    }
    res.total = Math.round(res.total * 100) / 100;
  } catch (e) {
    logEvent('WARN', '_wa_facturasPendientes_', e.message, '');
  }
  return res;
}

// ---- Acciones comunes ----

function _wa_init_(user, p) {
  var base = { rol: user.rol, nombre: user.nombre, hoy: _wa_hoyISO_() };
  if (user.rol === 'admin') {
    base.clientes = _wa_listaClientesActivos_();
    base.operarios = _wa_listaOperariosActivos_();
    base.tiposGasto = TIPOS_GASTO;
    base.estadosGasto = ESTADOS_PAGO_GASTO;
    base.estadosTrabajo = ESTADOS_TRABAJO;
    base.resumen = _wa_resumenAdmin_(user, {});
    base.maquinas = _maq_activasCorto_();
    base.estadosMaq = MAQ_.ESTADOS;
    base.combustibles = MAQ_.COMBUSTIBLES;
    base.tiposMaq = ['Desbrozadora', 'Cortacésped', 'Motosierra', 'Soplador', 'Cortasetos', 'Hidrolimpiadora', 'Taladro', 'Otra'];
    return base;
  }
  base.clientes = _wa_listaClientesActivos_();
  base.estadosParte = WA_.ESTADOS_PARTE;
  base.maquinas = _maq_activasCorto_();
  var mi = _wa_miInicio_(user, {});
  base.trabajos = mi.trabajos; base.recientes = mi.recientes; base.horasMes = mi.horasMes;
  return base;
}

// ---- Admin: resumen, pendientes, horas ----

function _wa_resumenAdmin_(user, p) {
  var trabajos = _wa_trabajos_();
  var hoy = new Date();
  var dow = (hoy.getDay() + 6) % 7;
  var lunes = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - dow);
  var domingo = new Date(lunes.getFullYear(), lunes.getMonth(), lunes.getDate() + 7);
  var semana = 0, pend = 0;
  trabajos.forEach(function(t) {
    if (t.fecha && t.fecha >= lunes && t.fecha < domingo) semana++;
    if (t.estado === 'Pendiente' || t.estado === 'En curso') pend++;
  });
  var fac = _wa_facturasPendientes_();
  return { trabajosSemana: semana, trabajosPendientes: pend, pendienteCobro: fac.total, facturasPendientes: fac.n, maq: _maq_resumen_() };
}

function _wa_pendientes_(user, p) {
  var trabajos = _wa_trabajos_().filter(function(t) { return t.estado === 'Pendiente' || t.estado === 'En curso'; });
  trabajos.sort(function(a, b) { return (a.fecha ? a.fecha.getTime() : 0) - (b.fecha ? b.fecha.getTime() : 0); });
  var fac = _wa_facturasPendientes_();
  return {
    trabajos: trabajos.slice(0, 60).map(function(t) {
      return { id: t.id, cliente: t.cliente, operario: t.operario, op2: t.op2, fecha: _wa_fechaTxt_(t.fecha), trabajo: t.trabajo, estado: t.estado, maq: t.maq };
    }),
    facturas: fac.lista, totalFacturas: fac.total
  };
}

function _wa_horasEquipo_(user, p) {
  var hoy = new Date();
  var y = hoy.getFullYear(), m = hoy.getMonth();
  var mm = String(p.mes || '').match(/^(\d{4})-(\d{2})$/);
  if (mm) { y = +mm[1]; m = +mm[2] - 1; }
  var map = {};
  _wa_trabajos_().forEach(function(t) {
    if (!_wa_esDelMes_(t.fecha, y, m)) return;
    [[t.operario, _calcHoras_(t.hIni, t.hFin)], [t.op2, t.op2 ? _calcHoras_(t.hIni2, t.hFin2) : 0]].forEach(function(par) {
      var nom = String(par[0] || '').trim();
      if (!nom) return;
      var k = _wa_norm_(nom);
      if (!map[k]) map[k] = { nombre: nom, horas: 0, trabajos: 0 };
      map[k].horas += par[1];
      map[k].trabajos++;
    });
  });
  var lista = Object.keys(map).map(function(k) {
    return { nombre: map[k].nombre, horas: _wa_redondea_(map[k].horas), trabajos: map[k].trabajos };
  }).sort(function(a, b) { return b.horas - a.horas; });
  return { mes: y + '-' + ('0' + (m + 1)).slice(-2), equipo: lista };
}

// ---- Admin: altas ----

function _wa_crearCliente_(user, p) {
  var d = {
    nombre: _wa_s_(p.nombre, 80), apellidos: _wa_s_(p.apellidos, 120), empresa: _wa_s_(p.empresa, 120),
    email: _wa_s_(p.email, 120).toLowerCase(), telefono: p.telefono,
    direccion: _wa_s_(p.direccion, 250), cifNif: _wa_s_(p.cifNif, 20)
  };
  if (!d.nombre || !d.apellidos) _wa_fail_('Escribe nombre y apellidos.');
  if (!_wa_emailOk_(d.email)) _wa_fail_('El email no es válido.');
  d.telefono = _wa_tel_(d.telefono);
  if (!d.direccion) _wa_fail_('Escribe la dirección.');

  return _wa_withLock_(function() {
    var nuevo = _wa_norm_(_wa_label_(d.nombre, d.apellidos, d.empresa));
    _wa_clientes_().forEach(function(c) {
      if (_wa_norm_(c.label) === nuevo) _wa_fail_('Ya existe un cliente con ese nombre' + (c.activo ? '.' : ' (está de baja: puedes reactivarlo en Gestión).'));
    });
    var r = createCliente(d);
    if (!r.success) _wa_fail_(r.error || 'No se pudo guardar el cliente.');
    return { id: r.id };
  });
}

function _wa_crearEmpleado_(user, p) {
  var d = {
    nombre: _wa_s_(p.nombre, 80), apellidos: _wa_s_(p.apellidos, 120),
    email: _wa_s_(p.email, 120).toLowerCase(), telefono: p.telefono
  };
  if (!d.nombre || !d.apellidos) _wa_fail_('Escribe nombre y apellidos.');
  if (!_wa_emailOk_(d.email)) _wa_fail_('El email no es válido.');
  d.telefono = _wa_tel_(d.telefono);

  return _wa_withLock_(function() {
    var nuevo = _wa_norm_(_wa_label_(d.nombre, d.apellidos));
    _wa_operarios_().forEach(function(o) {
      if (_wa_norm_(o.label) === nuevo) _wa_fail_('Ya existe un empleado con ese nombre' + (o.activo ? '.' : ' (está de baja: puedes reactivarlo en Gestión).'));
    });
    var r = createOperario(d);
    if (!r.success) _wa_fail_(r.error || 'No se pudo guardar el empleado.');
    var sheet = _wa_ensureOperarioCols_();
    var row = findRowById_(sheet, r.id);
    var tok = _wa_newToken_();
    sheet.getRange(row, WA_.COL_BLOQ).setValue(false);
    sheet.getRange(row, WA_.COL_TOKEN).setValue(tok);
    return { id: r.id, enlace: _wa_enlace_(tok) };
  });
}

function _wa_crearTrabajo_(user, p) {
  var clientes = _wa_listaClientesActivos_();
  var operarios = _wa_listaOperariosActivos_();
  var cliente = _wa_enActivos_(p.cliente, clientes);
  if (!cliente) _wa_fail_('Elige un cliente de la lista.');
  var operario = _wa_enActivos_(p.operario, operarios);
  if (!operario) _wa_fail_('Elige un empleado de la lista.');
  var operario2 = '';
  if (p.operario2) {
    operario2 = _wa_enActivos_(p.operario2, operarios);
    if (!operario2) _wa_fail_('El empleado 2 no es válido.');
    if (_wa_norm_(operario2) === _wa_norm_(operario)) _wa_fail_('El empleado 2 tiene que ser distinto del primero.');
  }
  var fecha = _wa_fechaDMY_(p.fecha);
  var estado = String(p.estado || 'Pendiente');
  if (ESTADOS_TRABAJO.indexOf(estado) === -1) _wa_fail_('El estado no es válido.');
  // Un trabajo futuro puede llevar solo la hora de inicio (no se sabe cuánto durará): el empleado
  // pondrá las horas reales al completarlo. Un trabajo ya completado sí necesita las dos.
  var h1 = _wa_hora_(p.horaInicio), h2 = _wa_hora_(p.horaFin);
  if (h2 && !h1) _wa_fail_('Pon también la hora de inicio, o deja las dos vacías.');
  if (h1 && !h2 && estado === 'Completado') _wa_fail_('Un trabajo completado necesita hora de inicio y de fin.');
  if (h1 && h2 && _wa_minutos_(h2) <= _wa_minutos_(h1)) _wa_fail_('La hora de fin debe ser posterior a la de inicio.');
  var i2 = '', f2 = '';
  if (operario2) {
    i2 = _wa_hora_(p.horaInicio2); f2 = _wa_hora_(p.horaFin2);
    if (f2 && !i2) _wa_fail_('Pon también la hora de inicio del empleado 2, o deja las dos vacías.');
    if (i2 && !f2 && estado === 'Completado') _wa_fail_('Un trabajo completado necesita hora de inicio y de fin del empleado 2.');
    if (i2 && f2 && _wa_minutos_(f2) <= _wa_minutos_(i2)) _wa_fail_('La hora de fin del empleado 2 debe ser posterior a la de inicio.');
  }
  var fotos = _wa_validaFotos_(p.fotos);
  var maqs = _maq_resolver_(p.maquinas, true);   // las rotas o en reparación no se pueden asignar

  return _wa_withLock_(function() {
    var id = nextId('TRABAJOS', 'TRA');
    var f = _wa_subirFotosTrabajo_(fotos, id);
    var r = createTrabajo({
      id: id, cliente: cliente, operario: operario, operario2: operario2,
      fechaTrabajo: fecha, horaInicio: h1, horaFin: h2, horaInicio2: i2, horaFin2: f2,
      trabajoRealizado: _wa_s_(p.trabajo, 2000), incidencias: _wa_s_(p.incidencias, 2000),
      estado: estado, ubicacion: _wa_s_(p.ubicacion, 250), materiales: _wa_num_(p.materiales, 100000),
      fotos: f.urls.join(', '), foto1: f.urls[0] || '', foto2: f.urls[1] || '', foto3: f.urls[2] || '',
      foto4: f.urls[3] || '', foto5: f.urls[4] || '', carpetaFotos: f.carpeta
    });
    if (!r.success) _wa_fail_(r.error || 'No se pudo guardar el trabajo.');
    _maq_ponerEnTrabajo_(id, maqs);
    return { id: id };
  });
}

// ---- Gastos ----

function _wa_crearGasto_(user, p) {
  var tipo = String(p.tipo || '');
  if (TIPOS_GASTO.indexOf(tipo) === -1) _wa_fail_('Elige un tipo de gasto.');
  var estadoPago = String(p.estadoPago || 'Pendiente');
  if (ESTADOS_PAGO_GASTO.indexOf(estadoPago) === -1) _wa_fail_('El estado del gasto no es válido.');
  var importe = _wa_num_(p.importe, 1000000);
  var obs = _wa_s_(p.observaciones, 1000);
  var fotos = _wa_validaFotos_(p.fotos ? [p.fotos] : []);

  var res = _wa_withLock_(function() {
    var sheet = getSheet('GASTOS');
    var ahora = new Date();
    var id = _nextGastoId_();
    var fotoUrl = '', archivoId = '', estadoProc = 'Pendiente OCR';
    if (fotos.length) {
      var f = _wa_subirFotoGasto_(fotos[0], id);
      fotoUrl = f.url; archivoId = f.id;
    } else {
      estadoProc = 'Revisar';
      obs = obs ? (obs + ' | Falta foto del ticket/factura') : 'Falta foto del ticket/factura';
    }
    var row = new Array(HEADERS_GASTOS.length).fill('');
    row[COL_GAS.ID_GASTO - 1] = id;
    row[COL_GAS.FECHA_ALTA - 1] = ahora;
    row[COL_GAS.TIPO_GASTO - 1] = tipo;
    row[COL_GAS.FECHA_GASTO - 1] = ahora;   // el OCR la sustituye si lee otra fecha en el ticket
    row[COL_GAS.ESTADO_PAGO - 1] = estadoPago;
    row[COL_GAS.FOTO_URL - 1] = fotoUrl;
    row[COL_GAS.ARCHIVO_ID - 1] = archivoId;
    row[COL_GAS.REVISADO - 1] = false;
    row[COL_GAS.ENVIAR_GESTOR - 1] = false;
    row[COL_GAS.OBSERVACIONES - 1] = obs;
    row[COL_GAS.ESTADO_PROCESAMIENTO - 1] = estadoProc;
    var target = _nextEmptyGastosRow_(sheet);
    sheet.getRange(target, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();
    logEvent('INFO', 'wa_crearGasto', 'Gasto registrado: ' + id, tipo + ' | ' + estadoProc);
    return { id: id, row: target, archivoId: archivoId };
  });

  // OCR fuera del bloqueo (puede tardar) — igual que el formulario de gastos
  if (res.archivoId) {
    try { procesarOCRGasto(res.id); }
    catch (e) {
      logEvent('ERROR', 'wa_crearGasto', 'OCR fallido para ' + res.id + ': ' + e.message, '');
      try { getSheet('GASTOS').getRange(res.row, COL_GAS.ESTADO_PROCESAMIENTO).setValue('OCR fallido'); } catch (_) {}
    }
  }
  // Importe indicado a mano: solo si el OCR no encontró total
  if (importe > 0) {
    try {
      var sh = getSheet('GASTOS');
      var celda = sh.getRange(res.row, COL_GAS.TOTAL);
      if (String(celda.getValue() || '').trim() === '') celda.setValue(importe);
    } catch (e2) { logEvent('WARN', 'wa_crearGasto', 'No se pudo guardar el importe: ' + e2.message, res.id); }
  }
  try {
    if (getTargetSS_().getSheetByName('PANEL_GASTOS')) _actualizarPanelGastos_();
  } catch (e3) { logEvent('WARN', 'wa_crearGasto', 'Panel no actualizado: ' + e3.message, ''); }
  return { id: res.id };
}

// ---- Fotos ----

var WA_MIME_EXT_ = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp' };

function _wa_validaFotos_(fotos) {
  if (!fotos) return [];
  if (!Array.isArray(fotos)) _wa_fail_('Las fotos no son válidas.');
  if (fotos.length > WA_.MAX_FOTOS) _wa_fail_('Máximo ' + WA_.MAX_FOTOS + ' fotos.');
  return fotos.map(function(f) {
    if (!f || !WA_MIME_EXT_[f.mime] || typeof f.b64 !== 'string' || !f.b64) _wa_fail_('Una de las fotos no es válida.');
    if (f.b64.length * 0.75 > WA_.MAX_FOTO_BYTES) _wa_fail_('Una foto es demasiado grande.');
    return f;
  });
}

function _wa_guardaFoto_(foto, folder, nombre) {
  var bytes = Utilities.base64Decode(foto.b64);
  var blob = Utilities.newBlob(bytes, foto.mime, nombre + '.' + WA_MIME_EXT_[foto.mime]);
  return folder.createFile(blob);
}

/** Guarda las fotos de un trabajo en carpeta_trabajos_id/año/ID (como el flujo del formulario). */
function _wa_subirFotosTrabajo_(fotos, id, carpetaExistente) {
  var res = { urls: [], carpeta: carpetaExistente || '' };
  if (!fotos.length) return res;
  var raiz = getConfig('carpeta_trabajos_id');
  var folder = _getOrCreateTrabajoFolder_(id, new Date().getFullYear(), raiz) || DriveApp.getRootFolder();
  if (folder.getId() !== DriveApp.getRootFolder().getId()) res.carpeta = folder.getUrl();
  fotos.forEach(function(f, i) {
    res.urls.push(_wa_guardaFoto_(f, folder, id + '_web_' + Date.now() + '_' + (i + 1)).getUrl());
  });
  return res;
}

function _wa_subirFotoGasto_(foto, id) {
  var folderId = getConfig(GAS_CFG.FOLDER_ID);
  var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var file = _wa_guardaFoto_(foto, folder, id + '_web');
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (_) {}
  return { id: file.getId(), url: file.getUrl() };
}

// ---- Gestión de personas ----

function _wa_listaPersonas_(user, p) {
  var tipo = String(p.tipo || '');
  var trabajos = _wa_trabajos_();
  var hoy = new Date(), y = hoy.getFullYear(), m = hoy.getMonth();

  if (tipo === 'e') {
    return _wa_operarios_().map(function(o) {
      var n = _wa_norm_(o.label), h = 0;
      trabajos.forEach(function(t) { if (_wa_esDelMes_(t.fecha, y, m)) h += _wa_horasDe_(t, n); });
      return {
        id: o.id, nombre: o.nombre, apellidos: o.apellidos, label: o.label, tel: o.tel, email: o.email,
        estado: !o.activo ? 'baja' : (o.bloq ? 'bloqueado' : 'activo'), horasMes: _wa_redondea_(h)
      };
    });
  }
  if (tipo === 'c') {
    var cuenta = {};
    trabajos.forEach(function(t) { var k = _wa_norm_(t.cliente); cuenta[k] = (cuenta[k] || 0) + 1; });
    return _wa_clientes_().map(function(c) {
      return {
        id: c.id, nombre: c.nombre, apellidos: c.apellidos, empresa: c.empresa, label: c.label,
        tel: c.tel, email: c.email, dir: c.dir, cif: c.cif,
        estado: c.activo ? 'activo' : 'baja', nTrabajos: cuenta[_wa_norm_(c.label)] || 0
      };
    });
  }
  _wa_fail_('Tipo no válido.');
}

function _wa_accionPersona_(user, p) {
  var tipo = String(p.tipo || ''), accion = String(p.accion || ''), id = String(p.id || '');
  return _wa_withLock_(function() {
    if (tipo === 'e') {
      if (['bloquear', 'desbloquear', 'baja', 'reactivar'].indexOf(accion) === -1) _wa_fail_('Acción no válida.');
      var sheet = _wa_ensureOperarioCols_();
      var row = findRowById_(sheet, id);
      if (!row) _wa_fail_('No se encuentra al empleado.');
      if (accion === 'bloquear')    { sheet.getRange(row, WA_.COL_BLOQ).setValue(true); }
      if (accion === 'desbloquear') { sheet.getRange(row, WA_.COL_BLOQ).setValue(false); }
      if (accion === 'baja')        { sheet.getRange(row, COL_OPE.ACTIVO).setValue(false); }
      if (accion === 'reactivar')   { sheet.getRange(row, COL_OPE.ACTIVO).setValue(true); sheet.getRange(row, WA_.COL_BLOQ).setValue(false); }
    } else if (tipo === 'c') {
      if (['baja', 'reactivar'].indexOf(accion) === -1) _wa_fail_('Acción no válida.');
      var sc = getSheet(SHEET_NAMES.CLIENTES);
      var rc = findRowById_(sc, id);
      if (!rc) _wa_fail_('No se encuentra al cliente.');
      sc.getRange(rc, COL_CLI.ACTIVO).setValue(accion === 'reactivar');
    } else {
      _wa_fail_('Tipo no válido.');
    }
    logEvent('INFO', 'wa_accionPersona', tipo + ' ' + id + ' → ' + accion, '');
    if (accion === 'baja' || accion === 'reactivar') {
      try { refreshTrabajosDropdowns(); } catch (e) { logEvent('WARN', 'wa_accionPersona', 'Desplegables no actualizados: ' + e.message, ''); }
    }
    return { ok: true };
  });
}

function _wa_editarPersona_(user, p) {
  var tipo = String(p.tipo || ''), id = String(p.id || ''), d = p.datos || {};
  return _wa_withLock_(function() {
    if (tipo === 'e') {
      var dat = { nombre: _wa_s_(d.nombre, 80), apellidos: _wa_s_(d.apellidos, 120), email: _wa_s_(d.email, 120).toLowerCase(), telefono: d.telefono };
      if (!dat.nombre || !dat.apellidos) _wa_fail_('Escribe nombre y apellidos.');
      if (!_wa_emailOk_(dat.email)) _wa_fail_('El email no es válido.');
      dat.telefono = _wa_tel_(dat.telefono);
      var r = updateOperario(id, dat);
      if (!r.success) _wa_fail_(r.error);
    } else if (tipo === 'c') {
      var dc = {
        nombre: _wa_s_(d.nombre, 80), apellidos: _wa_s_(d.apellidos, 120), empresa: _wa_s_(d.empresa, 120),
        email: _wa_s_(d.email, 120).toLowerCase(), telefono: d.telefono,
        direccion: _wa_s_(d.direccion, 250), cifNif: _wa_s_(d.cifNif, 20).toUpperCase()
      };
      if (!dc.nombre || !dc.apellidos) _wa_fail_('Escribe nombre y apellidos.');
      if (!_wa_emailOk_(dc.email)) _wa_fail_('El email no es válido.');
      if (!dc.direccion) _wa_fail_('Escribe la dirección.');
      dc.telefono = _wa_tel_(dc.telefono);
      var rc = updateCliente(id, dc);
      if (!rc.success) _wa_fail_(rc.error);
    } else {
      _wa_fail_('Tipo no válido.');
    }
    return { ok: true };
  });
}

// ---- Enlaces personales ----

/**
 * Dirección pública de la web app. NO se usa ScriptApp.getService().getUrl():
 * desde la hoja devuelve una dirección que no es la de la implementación publicada (da 404).
 * Se toma de CONFIG ("webapp_base_url") si existe y es válida; si no, la de la implementación actual.
 */
function _wa_baseUrl_() {
  var cfg = '';
  try { cfg = String(getConfig('webapp_base_url') || '').trim(); } catch (_) {}
  if (WA_.URL_RE.test(cfg)) return cfg;
  return WA_.URL_BASE;
}

function _wa_enlace_(token) {
  return _wa_baseUrl_() + '?t=' + token;
}

function _wa_tokenEmpleado_(id, regenerar) {
  return _wa_withLock_(function() {
    var sheet = _wa_ensureOperarioCols_();
    var row = findRowById_(sheet, id);
    if (!row) _wa_fail_('No se encuentra al empleado.');
    var celda = sheet.getRange(row, WA_.COL_TOKEN);
    var tok = String(celda.getValue() || '').trim();
    if (!tok || regenerar) { tok = _wa_newToken_(); celda.setValue(tok); }
    return tok;
  });
}

function _wa_enlaceEmpleado_(user, p) {
  return { enlace: _wa_enlace_(_wa_tokenEmpleado_(String(p.id || ''), false)) };
}

function _wa_regenerarEnlace_(user, p) {
  var r = { enlace: _wa_enlace_(_wa_tokenEmpleado_(String(p.id || ''), true)) };
  logEvent('INFO', 'wa_regenerarEnlace', 'Enlace regenerado: ' + p.id, '');
  return r;
}

// ---- Vista de empleado ----

function _wa_miInicio_(user, p) {
  var n = _wa_norm_(user.nombre);
  var hoy = new Date(), y = hoy.getFullYear(), m = hoy.getMonth();
  var mios = _wa_trabajos_().filter(function(t) {
    return _wa_norm_(t.operario) === n || (t.op2 && _wa_norm_(t.op2) === n);
  });
  var horas = 0;
  mios.forEach(function(t) { if (_wa_esDelMes_(t.fecha, y, m)) horas += _wa_horasDe_(t, n); });

  var asignados = mios.filter(function(t) { return t.estado === 'Pendiente' || t.estado === 'En curso'; })
    .sort(function(a, b) { return (a.fecha ? a.fecha.getTime() : 0) - (b.fecha ? b.fecha.getTime() : 0); })
    .slice(0, 30)
    .map(function(t) { return { id: t.id, cliente: t.cliente, fecha: _wa_fechaTxt_(t.fecha), trabajo: t.trabajo, estado: t.estado }; });

  var recientes = mios.filter(function(t) { return t.estado === 'Completado'; })
    .sort(function(a, b) { return (b.fecha ? b.fecha.getTime() : 0) - (a.fecha ? a.fecha.getTime() : 0); })
    .slice(0, 5)
    .map(function(t) {
      return { id: t.id, cliente: t.cliente, fecha: _wa_fechaTxt_(t.fecha), horas: _wa_redondea_(_wa_horasDe_(t, n)) };
    });
  return { trabajos: asignados, recientes: recientes, horasMes: _wa_redondea_(horas) };
}

function _wa_enviarParte_(user, p) {
  var h1 = _wa_hora_(p.horaInicio), h2 = _wa_hora_(p.horaFin);
  if (!h1 || !h2) _wa_fail_('Escribe la hora de inicio y la de fin.');
  if (_wa_minutos_(h2) <= _wa_minutos_(h1)) _wa_fail_('La hora de fin debe ser posterior a la de inicio.');
  var estado = String(p.estado || 'Completado');
  if (WA_.ESTADOS_PARTE.indexOf(estado) === -1) _wa_fail_('El estado no es válido.');
  var trabajoTxt = _wa_s_(p.trabajo, 2000);
  if (!trabajoTxt) _wa_fail_('Cuenta qué has hecho.');
  var incid = _wa_s_(p.incidencias, 2000);
  var materiales = _wa_num_(p.materiales, 100000);
  var fotos = _wa_validaFotos_(p.fotos);
  var maqs = _maq_resolver_(p.maquinas, false);   // en un parte se admite cualquiera: pudo romperse mientras se usaba
  var miNorm = _wa_norm_(user.nombre);

  // Rellenar un trabajo ya asignado
  if (p.idTrabajo) {
    return _wa_withLock_(function() {
      var sheet = getSheet('TRABAJOS');
      var row = findRowById_(sheet, String(p.idTrabajo));
      if (!row) _wa_fail_('No se encuentra ese trabajo.');
      var r = sheet.getRange(row, 1, 1, Math.min(COL_TRA.HORA_FIN_2, sheet.getMaxColumns())).getValues()[0];
      var esOp1 = _wa_norm_(r[COL_TRA.OPERARIO - 1]) === miNorm;
      var esOp2 = !esOp1 && _wa_norm_(r[COL_TRA.OPERARIO_2 - 1]) === miNorm;
      if (!esOp1 && !esOp2) _wa_fail_('Ese trabajo no es tuyo.');
      var id = String(p.idTrabajo);

      if (esOp1) { sheet.getRange(row, COL_TRA.HORA_INICIO).setValue(h1); sheet.getRange(row, COL_TRA.HORA_FIN).setValue(h2); }
      else       { sheet.getRange(row, COL_TRA.HORA_INICIO_2).setValue(h1); sheet.getRange(row, COL_TRA.HORA_FIN_2).setValue(h2); }
      sheet.getRange(row, COL_TRA.TRABAJO_REALIZADO).setValue(trabajoTxt);
      if (incid) sheet.getRange(row, COL_TRA.INCIDENCIAS).setValue(incid);
      sheet.getRange(row, COL_TRA.ESTADO).setValue(estado);
      if (materiales > 0) sheet.getRange(row, COL_TRA.MATERIALES).setValue(materiales);
      if (!r[COL_TRA.FECHA_TRABAJO - 1]) {
        sheet.getRange(row, COL_TRA.FECHA_TRABAJO).setValue(_parseReportDate_(_wa_hoyDMY_()));
      }

      if (fotos.length) {
        var carpeta = String(r[COL_TRA.CARPETA_FOTOS - 1] || '');
        var f = _wa_subirFotosTrabajo_(fotos, id, carpeta);
        var libres = [];
        for (var k = 0; k < 5; k++) if (!String(r[COL_TRA.FOTO_1 - 1 + k] || '').trim()) libres.push(COL_TRA.FOTO_1 + k);
        f.urls.forEach(function(u, i) { if (libres[i]) sheet.getRange(row, libres[i]).setValue(u); });
        var previas = String(r[COL_TRA.FOTOS - 1] || '').trim();
        sheet.getRange(row, COL_TRA.FOTOS).setValue((previas ? previas + ', ' : '') + f.urls.join(', '));
        if (f.carpeta) sheet.getRange(row, COL_TRA.CARPETA_FOTOS).setValue(f.carpeta);
      }
      _maq_ponerEnTrabajo_(id, maqs);
      logEvent('INFO', 'wa_enviarParte', 'Parte completado: ' + id, user.nombre);
      return { id: id };
    });
  }

  // Parte nuevo
  var cliente = _wa_enActivos_(p.cliente, _wa_listaClientesActivos_());
  if (!cliente) _wa_fail_('Elige un cliente de la lista.');
  return _wa_withLock_(function() {
    var id = nextId('TRABAJOS', 'TRA');
    var f = _wa_subirFotosTrabajo_(fotos, id);
    var r = createTrabajo({
      id: id, cliente: cliente, operario: user.nombre, fechaTrabajo: _wa_hoyDMY_(),
      horaInicio: h1, horaFin: h2, trabajoRealizado: trabajoTxt, incidencias: incid, estado: estado,
      materiales: materiales, fotos: f.urls.join(', '), foto1: f.urls[0] || '', foto2: f.urls[1] || '',
      foto3: f.urls[2] || '', foto4: f.urls[3] || '', foto5: f.urls[4] || '', carpetaFotos: f.carpeta
    });
    if (!r.success) _wa_fail_(r.error || 'No se pudo guardar el parte.');
    _maq_ponerEnTrabajo_(id, maqs);
    logEvent('INFO', 'wa_enviarParte', 'Parte nuevo: ' + id, user.nombre);
    return { id: id };
  });
}

// ---- Enlaces (menú de la hoja) ----

/**
 * Muestra el enlace de Álvaro (admin). Crea el token la primera vez.
 * Los enlaces de los empleados se piden desde la propia app (Gestión → ficha).
 */
function webappMostrarEnlaces() {
  var ui = SpreadsheetApp.getUi();
  var tok = getConfig(WA_.CFG_ADMIN_TOKEN);
  if (!tok) { tok = _wa_newToken_(); setConfig(WA_.CFG_ADMIN_TOKEN, tok); }
  _wa_ensureOperarioCols_();
  var url = _wa_enlace_(tok);
  Logger.log('Enlace de Álvaro: ' + url);
  ui.showModalDialog(_wa_dialogoEnlace_(url, 'Enlace personal de Álvaro',
    'Guárdalo en favoritos o en la pantalla de inicio del móvil. No lo compartas: da acceso a todo. ' +
    'Los enlaces de los empleados se generan desde la app: Gestión → empleado → "Enlace personal".'),
    'Enlace personal de Álvaro');
}

function _wa_htmlEsc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Ventana con el enlace en un cuadro seleccionable, botón "Copiar" y enlace para abrirlo. */
function _wa_dialogoEnlace_(url, titulo, nota) {
  var u = _wa_htmlEsc_(url);
  var html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#202124">' +
    '<p style="margin:0 0 8px">' + _wa_htmlEsc_(nota) + '</p>' +
    '<textarea id="u" readonly onfocus="this.select()" style="width:100%;height:96px;box-sizing:border-box;font-size:13px;padding:8px">' + u + '</textarea>' +
    '<p style="margin:10px 0 0">' +
    '<button id="c" style="padding:8px 14px;font-size:14px;cursor:pointer">Copiar enlace</button> ' +
    '<a href="' + u + '" target="_blank" style="margin-left:12px">Abrir en una pestaña nueva</a></p>' +
    '<p id="m" style="margin:8px 0 0;color:#188038;min-height:18px"></p>' +
    '<script>' +
    'var t=document.getElementById("u"),m=document.getElementById("m");' +
    't.select();' +
    'document.getElementById("c").onclick=function(){t.focus();t.select();var ok=false;' +
    'try{ok=document.execCommand("copy");}catch(e){}' +
    'if(!ok&&navigator.clipboard){navigator.clipboard.writeText(t.value).then(function(){m.textContent="Enlace copiado";},function(){m.textContent="Pulsa Ctrl+C para copiarlo";});}' +
    'else{m.textContent=ok?"Enlace copiado":"Pulsa Ctrl+C para copiarlo";}};' +
    '</script></div>';
  return HtmlService.createHtmlOutput(html).setWidth(520).setHeight(300).setTitle(titulo);
}

/** Alternativa sin menú: escribe el enlace de Álvaro en el registro del editor (Ver → Registros). */
function webappLogEnlace() {
  var tok = getConfig(WA_.CFG_ADMIN_TOKEN);
  if (!tok) { tok = _wa_newToken_(); setConfig(WA_.CFG_ADMIN_TOKEN, tok); }
  _wa_ensureOperarioCols_();
  Logger.log('Enlace de Álvaro: ' + _wa_enlace_(tok));
}

/** Genera un enlace de Álvaro nuevo (el anterior deja de funcionar). Ejecutar solo si se filtra. */
function webappRegenerarEnlaceAdmin() {
  setConfig(WA_.CFG_ADMIN_TOKEN, _wa_newToken_());
  Logger.log('Enlace de admin regenerado. Ábrelo desde el menú Sistema → Web app — Ver enlace.');
}
