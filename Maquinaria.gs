// ============================================================
// Maquinaria.gs — Inventario de maquinaria con QR (web app)
// ============================================================
//
// Hojas (se crean solas la primera vez):
//   MAQUINARIA  una fila por máquina (estado, ubicación, foto, último uso...)
//   MAQ_USOS    historial: usos escaneados por QR, incidencias y cambios de estado
// TRABAJOS gana una columna "Maquinaria" (col. 23) con los nombres de las máquinas usadas.
//
// Estados:  Disponible · En uso · En reparación · Rota
//   - Se deja en "En el depósito" o "En la camioneta"  → Disponible
//   - Se deja "En casa del cliente" o "Sigo usándola"   → En uso
//   - Incidencia con "En reparación" o "Rota"           → ese estado (manda sobre lo anterior)
//   - Una máquina Rota / En reparación solo la saca de ese estado Álvaro.
//
// El QR de cada máquina contiene  <dirección de la web app>?m=MAQ-001  (sin ningún código secreto).
// ============================================================

var MAQ_ = {
  SHEET:      'MAQUINARIA',
  SHEET_USOS: 'MAQ_USOS',
  HEAD: ['ID_Maquina', 'Fecha_Alta', 'Nombre', 'Tipo', 'Marca_Modelo', 'Num_Serie', 'Estado', 'Ubicacion',
         'Cliente_Actual', 'Ultimo_Operario', 'Ultimo_Uso', 'Foto_URL', 'Foto_ID', 'Notas', 'Activo'],
  COL: { ID: 1, ALTA: 2, NOMBRE: 3, TIPO: 4, MARCA: 5, SERIE: 6, ESTADO: 7, UBIC: 8, CLIENTE: 9,
         OPERARIO: 10, ULTUSO: 11, FOTO_URL: 12, FOTO_ID: 13, NOTAS: 14, ACTIVO: 15 },
  HEAD_USOS: ['ID_Uso', 'Fecha_Hora', 'ID_Maquina', 'Maquina', 'ID_Trabajo', 'Cliente', 'Operario',
              'Dejada_En', 'Incidencia', 'Descripcion', 'Estado_Resultante', 'Origen'],
  ESTADOS: ['Disponible', 'En uso', 'En reparación', 'Rota'],
  DEJADA:  ['Sigo usándola yo', 'En casa del cliente', 'En la camioneta', 'En el depósito'],
  RESULTADOS: ['Sigue en servicio', 'En reparación', 'Rota'],
  UBICACIONES_ADMIN: ['Depósito', 'Camioneta'],
  COL_TRA: 23,
  CFG_CARPETA: 'maquinaria_folder_id',
  ID_RE: /^MAQ-\d{3,6}$/
};

// ---- Hojas ----

function _maq_hoja_() {
  var sh = getOrCreateSheet(MAQ_.SHEET);
  if (sh.getLastRow() === 0) { sh.appendRow(MAQ_.HEAD); applyHeaderStyle_(sh, MAQ_.HEAD.length); }
  return sh;
}

function _maq_hojaUsos_() {
  var sh = getOrCreateSheet(MAQ_.SHEET_USOS);
  if (sh.getLastRow() === 0) { sh.appendRow(MAQ_.HEAD_USOS); applyHeaderStyle_(sh, MAQ_.HEAD_USOS.length); }
  return sh;
}

function _maq_fecha_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM HH:mm');
  }
  return String(v || '');
}

function _maq_todas_() {
  var sh = _maq_hoja_(), last = sh.getLastRow();
  if (last < 2) return [];
  var C = MAQ_.COL, vals = sh.getRange(2, 1, last - 1, MAQ_.HEAD.length).getValues(), out = [];
  vals.forEach(function(r, i) {
    var id = String(r[C.ID - 1] || '').trim();
    if (!id) return;
    out.push({
      row: i + 2, id: id,
      nombre: String(r[C.NOMBRE - 1] || '').trim(), tipo: String(r[C.TIPO - 1] || '').trim(),
      marca: String(r[C.MARCA - 1] || '').trim(), serie: String(r[C.SERIE - 1] || '').trim(),
      estado: MAQ_.ESTADOS.indexOf(String(r[C.ESTADO - 1])) !== -1 ? String(r[C.ESTADO - 1]) : 'Disponible',
      ubic: String(r[C.UBIC - 1] || '').trim(), cliente: String(r[C.CLIENTE - 1] || '').trim(),
      operario: String(r[C.OPERARIO - 1] || '').trim(), ultUso: _maq_fecha_(r[C.ULTUSO - 1]),
      fotoId: String(r[C.FOTO_ID - 1] || '').trim(), notas: String(r[C.NOTAS - 1] || '').trim(),
      activo: _wa_bool_(r[C.ACTIVO - 1])
    });
  });
  return out;
}

function _maq_buscar_(id, soloActivas) {
  id = String(id || '');
  if (!MAQ_.ID_RE.test(id)) _wa_fail_('Esa máquina no existe.');
  var m = _maq_todas_().filter(function(x) { return x.id === id; })[0];
  if (!m || (soloActivas && !m.activo)) _wa_fail_('Esa máquina no existe o está de baja.');
  return m;
}

function _maq_qrUrl_(id) { return _wa_baseUrl_() + '?m=' + id; }

function _maq_pub_(m) {
  return {
    id: m.id, nombre: m.nombre, tipo: m.tipo, marca: m.marca, serie: m.serie, estado: m.estado,
    ubic: m.ubic, cliente: m.cliente, operario: m.operario, ultUso: m.ultUso, notas: m.notas, activo: m.activo,
    foto: m.fotoId ? 'https://drive.google.com/thumbnail?id=' + m.fotoId + '&sz=w640' : '',
    qr: _maq_qrUrl_(m.id)
  };
}

function _maq_ordenadas_(lista) {
  var prio = { 'Rota': 0, 'En reparación': 1, 'En uso': 2, 'Disponible': 3 };
  return lista.sort(function(a, b) {
    return (prio[a.estado] - prio[b.estado]) || a.nombre.localeCompare(b.nombre);
  });
}

/** Lista corta (id, nombre, estado) de máquinas activas: para los selectores de trabajos y partes. */
function _maq_activasCorto_() {
  try {
    return _maq_ordenadas_(_maq_todas_().filter(function(m) { return m.activo; }))
      .map(function(m) { return { id: m.id, nombre: m.nombre, estado: m.estado }; });
  } catch (e) {
    logEvent('WARN', '_maq_activasCorto_', e.message, '');
    return [];
  }
}

function _maq_resumen_() {
  var res = { total: 0, 'Disponible': 0, 'En uso': 0, 'En reparación': 0, 'Rota': 0 };
  try {
    _maq_todas_().forEach(function(m) { if (m.activo) { res.total++; res[m.estado]++; } });
  } catch (e) { logEvent('WARN', '_maq_resumen_', e.message, ''); }
  return res;
}

// ---- Fotos (carpeta propia en Drive; visible con el enlace para poder mostrarla en la app) ----

function _maq_subirFoto_(foto, id) {
  var folder = null, fid = getConfig(MAQ_.CFG_CARPETA);
  if (fid) { try { folder = DriveApp.getFolderById(fid); } catch (_) { folder = null; } }
  if (!folder) {
    folder = DriveApp.createFolder('Maquinaria - fotos');
    setConfig(MAQ_.CFG_CARPETA, folder.getId());
  }
  var file = _wa_guardaFoto_(foto, folder, id + '_' + Date.now());
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (_) {}
  return file.getId();
}

function _maq_limpia_(v, max) { return _wa_s_(v, max).replace(/,/g, ' ').replace(/\s+/g, ' ').trim(); }

function _maq_nombreLibre_(nombre, idPropio) {
  var n = _wa_norm_(nombre);
  _maq_todas_().forEach(function(m) {
    if (m.activo && m.id !== idPropio && _wa_norm_(m.nombre) === n) {
      _wa_fail_('Ya hay una máquina activa con ese nombre. Añade la marca o un número para distinguirlas (por ejemplo "Desbrozadora Stihl 2").');
    }
  });
}

function _maq_registraUso_(m, idTrabajo, cliente, operario, dejada, incidencia, descripcion, resultado, origen) {
  _maq_hojaUsos_().appendRow([
    nextId(MAQ_.SHEET_USOS, 'USO'), new Date(), m.id, m.nombre, idTrabajo || '', cliente || '', operario || '',
    dejada || '', incidencia ? true : false, descripcion || '', resultado || '', origen || 'App'
  ]);
}

function _maq_usos_(idMaq, n) {
  var sh = _maq_hojaUsos_(), last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, MAQ_.HEAD_USOS.length).getValues(), out = [];
  for (var i = vals.length - 1; i >= 0 && out.length < n; i--) {
    var r = vals[i];
    if (String(r[2]) !== idMaq) continue;
    out.push({
      fecha: _maq_fecha_(r[1]), operario: String(r[6] || ''), cliente: String(r[5] || ''), dejada: String(r[7] || ''),
      incidencia: _wa_bool_(r[8]), descripcion: String(r[9] || ''), resultado: String(r[10] || ''), origen: String(r[11] || '')
    });
  }
  return out;
}

// ---- Acciones: lectura ----

function wa_maqLista_(user, p) {
  var admin = user.rol === 'admin';
  var todas = _maq_todas_().filter(function(m) { return admin || m.activo; });
  var activas = _maq_ordenadas_(todas.filter(function(m) { return m.activo; }));
  var bajas = todas.filter(function(m) { return !m.activo; }).sort(function(a, b) { return a.nombre.localeCompare(b.nombre); });
  return { maquinas: activas.concat(bajas).map(_maq_pub_), resumen: _maq_resumen_() };
}

function wa_maqFicha_(user, p) {
  var m = _maq_buscar_(p.id, user.rol !== 'admin');
  return { maquina: _maq_pub_(m), usos: _maq_usos_(m.id, user.rol === 'admin' ? 15 : 3) };
}

// ---- Acciones: alta y edición (Álvaro) ----

function wa_maqCrear_(user, p) {
  var nombre = _maq_limpia_(p.nombre, 80);
  if (!nombre) _wa_fail_('Escribe el nombre de la máquina.');
  var tipo = _maq_limpia_(p.tipo, 40), marca = _maq_limpia_(p.marca, 80), serie = _maq_limpia_(p.serie, 40), notas = _wa_s_(p.notas, 500);
  var fotos = _wa_validaFotos_(p.foto ? [p.foto] : []);
  return _wa_withLock_(function() {
    _maq_nombreLibre_(nombre, '');
    var sh = _maq_hoja_(), C = MAQ_.COL, id = nextId(MAQ_.SHEET, 'MAQ');
    var fotoId = fotos.length ? _maq_subirFoto_(fotos[0], id) : '';
    var fila = new Array(MAQ_.HEAD.length).fill('');
    fila[C.ID - 1] = id; fila[C.ALTA - 1] = new Date(); fila[C.NOMBRE - 1] = nombre; fila[C.TIPO - 1] = tipo;
    fila[C.MARCA - 1] = marca; fila[C.SERIE - 1] = serie; fila[C.ESTADO - 1] = 'Disponible';
    fila[C.UBIC - 1] = 'Depósito'; fila[C.FOTO_ID - 1] = fotoId; fila[C.NOTAS - 1] = notas; fila[C.ACTIVO - 1] = true;
    sh.appendRow(fila);
    logEvent('INFO', 'wa_maqCrear', 'Máquina creada: ' + id, nombre);
    return { id: id, qr: _maq_qrUrl_(id) };
  });
}

function wa_maqEditar_(user, p) {
  var d = p.datos || {};
  var nombre = _maq_limpia_(d.nombre, 80);
  if (!nombre) _wa_fail_('Escribe el nombre de la máquina.');
  var fotos = _wa_validaFotos_(d.foto ? [d.foto] : []);
  return _wa_withLock_(function() {
    var m = _maq_buscar_(p.id, false), C = MAQ_.COL, sh = _maq_hoja_();
    if (m.activo) _maq_nombreLibre_(nombre, m.id);
    sh.getRange(m.row, C.NOMBRE).setValue(nombre);
    sh.getRange(m.row, C.TIPO).setValue(_maq_limpia_(d.tipo, 40));
    sh.getRange(m.row, C.MARCA).setValue(_maq_limpia_(d.marca, 80));
    sh.getRange(m.row, C.SERIE).setValue(_maq_limpia_(d.serie, 40));
    sh.getRange(m.row, C.NOTAS).setValue(_wa_s_(d.notas, 500));
    if (fotos.length) sh.getRange(m.row, C.FOTO_ID).setValue(_maq_subirFoto_(fotos[0], m.id));
    return { ok: true };
  });
}

/** Álvaro cambia el estado a mano (p. ej. "ya está reparada") y, si quiere, su ubicación. Queda en el historial. */
function wa_maqEstado_(user, p) {
  var estado = String(p.estado || '');
  if (MAQ_.ESTADOS.indexOf(estado) === -1) _wa_fail_('El estado no es válido.');
  var ubic = String(p.ubicacion || '');
  if (ubic && MAQ_.UBICACIONES_ADMIN.indexOf(ubic) === -1) _wa_fail_('La ubicación no es válida.');
  var nota = _wa_s_(p.nota, 300);
  return _wa_withLock_(function() {
    var m = _maq_buscar_(p.id, true), C = MAQ_.COL, sh = _maq_hoja_();
    sh.getRange(m.row, C.ESTADO).setValue(estado);
    if (ubic) { sh.getRange(m.row, C.UBIC).setValue(ubic); sh.getRange(m.row, C.CLIENTE).setValue(''); }
    _maq_registraUso_(m, '', '', user.nombre, ubic || m.ubic, false, nota, estado, 'Cambio de estado');
    logEvent('INFO', 'wa_maqEstado', m.id + ' → ' + estado, nota);
    return { ok: true, estado: estado };
  });
}

function wa_maqBaja_(user, p) {
  return _wa_withLock_(function() {
    var m = _maq_buscar_(p.id, false);
    _maq_hoja_().getRange(m.row, MAQ_.COL.ACTIVO).setValue(p.activo === true);
    logEvent('INFO', 'wa_maqBaja', m.id + (p.activo === true ? ' reactivada' : ' de baja'), '');
    return { ok: true };
  });
}

// ---- Acción: registrar uso (escaneando el QR) ----

/** Último trabajo del operario para ese cliente que siga abierto o sea de hoy. */
function _maq_buscarTrabajo_(operarioLabel, cliente) {
  try {
    var n = _wa_norm_(operarioLabel), c = _wa_norm_(cliente), hoy = _wa_hoyDMY_();
    var cand = _wa_trabajos_().filter(function(t) {
      if (_wa_norm_(t.cliente) !== c) return false;
      if (!(_wa_norm_(t.operario) === n || (t.op2 && _wa_norm_(t.op2) === n))) return false;
      var abierto = t.estado === 'Pendiente' || t.estado === 'En curso';
      var esHoy = t.fecha && Utilities.formatDate(t.fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy') === hoy;
      return abierto || esHoy;
    });
    cand.sort(function(a, b) { return ((b.fecha ? b.fecha.getTime() : 0) - (a.fecha ? a.fecha.getTime() : 0)) || (b.row - a.row); });
    return cand.length ? cand[0].id : '';
  } catch (e) { return ''; }
}

function _maq_ubicacionDe_(dejada, cliente, operario) {
  if (dejada === 'En casa del cliente') return 'En casa de ' + cliente;
  if (dejada === 'En la camioneta') return 'Camioneta';
  if (dejada === 'En el depósito') return 'Depósito';
  return 'Con ' + operario;
}

function wa_maqUso_(user, p) {
  var cliente = _wa_enActivos_(p.cliente, _wa_listaClientesActivos_());
  if (!cliente) _wa_fail_('Elige el cliente donde la has usado.');
  var dejada = String(p.dejadaEn || '');
  if (MAQ_.DEJADA.indexOf(dejada) === -1) _wa_fail_('Indica dónde has dejado la máquina.');
  var incidencia = p.incidencia === true;
  var desc = '', resultado = '';
  if (incidencia) {
    desc = _wa_s_(p.descripcion, 500);
    if (desc.length < 3) _wa_fail_('Cuenta qué ha pasado con la máquina.');
    resultado = String(p.resultado || '');
    if (MAQ_.RESULTADOS.indexOf(resultado) === -1) _wa_fail_('Indica si la máquina sigue en servicio, en reparación o rota.');
  }
  return _wa_withLock_(function() {
    var m = _maq_buscar_(p.id, true), C = MAQ_.COL, sh = _maq_hoja_();
    var base = (dejada === 'En casa del cliente' || dejada === 'Sigo usándola yo') ? 'En uso' : 'Disponible';
    var estado = base;
    if (incidencia && resultado === 'Rota') estado = 'Rota';
    else if (incidencia && resultado === 'En reparación') estado = 'En reparación';
    else if (m.estado === 'Rota' || m.estado === 'En reparación') estado = m.estado;   // solo Álvaro la saca de ahí
    var idTrabajo = _maq_buscarTrabajo_(user.nombre, cliente);

    sh.getRange(m.row, C.ESTADO).setValue(estado);
    sh.getRange(m.row, C.UBIC).setValue(_maq_ubicacionDe_(dejada, cliente, user.nombre));
    sh.getRange(m.row, C.CLIENTE).setValue(dejada === 'En casa del cliente' ? cliente : '');
    sh.getRange(m.row, C.OPERARIO).setValue(user.nombre);
    sh.getRange(m.row, C.ULTUSO).setValue(new Date());
    _maq_registraUso_(m, idTrabajo, cliente, user.nombre, dejada, incidencia, desc, resultado, 'QR/App');
    if (idTrabajo) _maq_ponerEnTrabajo_(idTrabajo, [m]);
    logEvent(incidencia ? 'WARN' : 'INFO', 'wa_maqUso', m.id + ' usada por ' + user.nombre + ' en ' + cliente + ' → ' + estado, desc);
    return { id: m.id, nombre: m.nombre, estado: estado, ubicacion: _maq_ubicacionDe_(dejada, cliente, user.nombre), trabajo: idTrabajo };
  });
}

// ---- Enlace con TRABAJOS (columna "Maquinaria") ----

function _maq_ensureColTrabajos_(sh) {
  if (sh.getMaxColumns() < MAQ_.COL_TRA) sh.insertColumnsAfter(sh.getMaxColumns(), MAQ_.COL_TRA - sh.getMaxColumns());
  var celda = sh.getRange(1, MAQ_.COL_TRA), cur = String(celda.getValue() || '').trim();
  if (!cur) celda.setValue('Maquinaria');
  else if (cur !== 'Maquinaria') throw new Error('TRABAJOS: la columna ' + MAQ_.COL_TRA + ' ya se usa para "' + cur + '".');
}

/** Valida y devuelve las máquinas de una lista de IDs. soloUsables=true rechaza las rotas o en reparación. */
function _maq_resolver_(ids, soloUsables) {
  if (!ids) return [];
  if (!Array.isArray(ids) || ids.length > 15) _wa_fail_('La lista de maquinaria no es válida.');
  var todas = _maq_todas_(), out = [];
  ids.forEach(function(id) {
    var m = todas.filter(function(x) { return x.id === String(id) && x.activo; })[0];
    if (!m) _wa_fail_('Una de las máquinas ya no existe o está de baja.');
    if (soloUsables && (m.estado === 'Rota' || m.estado === 'En reparación')) {
      _wa_fail_('"' + m.nombre + '" está ' + (m.estado === 'Rota' ? 'rota' : 'en reparación') + ' y no se puede asignar.');
    }
    if (out.indexOf(m) === -1) out.push(m);
  });
  return out;
}

/** Añade los nombres de las máquinas a la fila del trabajo sin repetir. */
function _maq_ponerEnTrabajo_(idTrabajo, maquinas) {
  if (!idTrabajo || !maquinas || !maquinas.length) return;
  var sh = getSheet('TRABAJOS');
  _maq_ensureColTrabajos_(sh);
  var row = findRowById_(sh, idTrabajo);
  if (!row) return;
  var celda = sh.getRange(row, MAQ_.COL_TRA), cur = String(celda.getValue() || '').trim();
  var lista = cur ? cur.split(/\s*,\s*/) : [];
  maquinas.forEach(function(m) { if (lista.indexOf(m.nombre) === -1) lista.push(m.nombre); });
  celda.setValue(lista.join(', '));
}
