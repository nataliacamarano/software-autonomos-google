// =============================================================
// Gestion.gs — Trabajos y horas de un cliente o de un empleado en un periodo
// =============================================================
//
// Para Álvaro, desde "Gestionar empleados y clientes": elige un mes o un rango de fechas y ve
// el resumen (trabajos y horas) y la lista de trabajos que hay detrás. En los clientes, cada
// trabajo dice si está en una factura (o proforma) y si esa factura está pendiente de pago.
//
// El vínculo trabajo↔factura es la columna IDs_Trabajos de FACTURAS y PROFORMAS; el cobro sale de
// Estado / Estado_Cobro / Fecha_Pago (mismas reglas que "Pendiente de cobro" del inicio).
// =============================================================

var GES_ = { MAX_DIAS: 1100 };

function _ges_hdr_(h) { return String(h).toLowerCase().replace(/[\s_]/g, ''); }

/** idTrabajo → lista de documentos (facturas y proformas) que lo incluyen. */
function _ges_documentos_() {
  var porTrabajo = {}, docs = {};
  [['FACTURAS', 'Factura', 'idfactura'], ['PROFORMAS', 'Proforma', 'idproforma']].forEach(function(par) {
    var sh;
    try { sh = getSheet(par[0]); } catch (e) { return; }
    var last = sh.getLastRow();
    if (last < 2) return;
    var vals = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
    var hd = vals[0].map(_ges_hdr_);
    function ix(n) { return hd.indexOf(n); }
    var cId = ix(par[2]), cIds = ix('idstrabajos');
    if (cId < 0 || cIds < 0) return;
    var cSerie = ix('serie'), cNum = ix('numero'), cTot = ix('total'), cEst = ix('estado'), cCobro = ix('estadocobro'),
        cPago = ix('fechapago'), cFec = ix('fechaemision');
    for (var i = 1; i < vals.length; i++) {
      var r = vals[i], id = String(r[cId] || '').trim();
      if (!id) continue;
      var est = cEst >= 0 ? String(r[cEst] || '').toLowerCase() : '', cob = cCobro >= 0 ? String(r[cCobro] || '').toLowerCase() : '';
      var pagada = (cPago >= 0 && String(r[cPago] || '').trim() !== '') || /pagad|cobrad/.test(est + ' ' + cob);
      var anulada = /anulad/.test(est + ' ' + cob);
      var doc = {
        id: id, tipo: par[1], total: cTot >= 0 ? Math.round((Number(r[cTot]) || 0) * 100) / 100 : 0,
        fecha: cFec >= 0 ? _ag_iso_(r[cFec]) : '',
        pago: anulada ? 'anulada' : (pagada ? 'pagada' : 'pendiente'),
        trabajos: 0
      };
      docs[par[1] + ':' + id] = doc;
      String(r[cIds] || '').split(/[,;\s]+/).filter(Boolean).forEach(function(tid) {
        (porTrabajo[tid] = porTrabajo[tid] || []).push(doc);
        doc.trabajos++;
      });
    }
  });
  return porTrabajo;
}

/** De los documentos de un trabajo, el que cuenta: la factura antes que la proforma, y nunca una anulada. */
function _ges_docDe_(lista) {
  var buenos = (lista || []).filter(function(d) { return d.pago !== 'anulada'; });
  return buenos.filter(function(d) { return d.tipo === 'Factura'; })[0] || buenos[0] || null;
}

function _ges_horasTotales_(t) {
  return _calcHoras_(t.hIni, t.hFin) + (t.op2 ? _calcHoras_(t.hIni2, t.hFin2) : 0);
}

/**
 * Álvaro: trabajos y horas de un cliente (tipo 'c') o de un empleado (tipo 'e') entre dos fechas.
 * p = { tipo, id, desde: 'yyyy-mm-dd', hasta: 'yyyy-mm-dd' }
 */
function wa_gestionPeriodo_(user, p) {
  var tipo = String(p.tipo || ''), id = String(p.id || '');
  if (tipo !== 'c' && tipo !== 'e') _wa_fail_('Tipo no válido.');
  var desde = _ag_fechaIso_(p.desde), hasta = _ag_fechaIso_(p.hasta);
  if (desde > hasta) _wa_fail_('La fecha de inicio no puede ser posterior a la de fin.');
  var dias = (new Date(hasta + 'T00:00:00Z').getTime() - new Date(desde + 'T00:00:00Z').getTime()) / 86400000;
  if (dias > GES_.MAX_DIAS) _wa_fail_('El periodo es demasiado largo (máximo unos 3 años).');

  var persona = (tipo === 'c' ? _wa_clientes_() : _wa_operarios_()).filter(function(x) { return x.id === id; })[0];
  if (!persona) _wa_fail_('No encuentro esa persona.');
  var n = _wa_norm_(persona.label);

  var docsPorTrabajo = tipo === 'c' ? _ges_documentos_() : {};
  var out = [], docsVistos = {};
  _wa_trabajos_().forEach(function(t) {
    var f = _ag_iso_(t.fecha);
    if (!f || f < desde || f > hasta || t.estado === 'Cancelado') return;
    if (tipo === 'c') { if (_wa_norm_(t.cliente) !== n) return; }
    else if (!(_wa_norm_(t.operario) === n || (t.op2 && _wa_norm_(t.op2) === n))) return;
    var fila = { id: t.id, fecha: f, estado: t.estado, trabajo: t.trabajo, hIni: _wa_horaTxt_(t.hIni), hFin: _wa_horaTxt_(t.hFin) };
    if (tipo === 'c') {
      fila.horas = _wa_redondea_(_ges_horasTotales_(t));
      fila.operarios = [t.operario, t.op2].filter(Boolean).join(' y ');
      var d = t.estado === 'Completado' ? _ges_docDe_(docsPorTrabajo[t.id]) : null;
      fila.factura = d ? { id: d.id, tipo: d.tipo, pago: d.pago, total: d.total, fecha: d.fecha } : null;
      if (d) docsVistos[d.tipo + ':' + d.id] = d;
    } else {
      fila.horas = _wa_redondea_(_wa_horasDe_(t, n));
      fila.cliente = t.cliente;
    }
    out.push(fila);
  });
  out.sort(function(a, b) { return a.fecha !== b.fecha ? (a.fecha < b.fecha ? 1 : -1) : (a.id < b.id ? 1 : -1); });

  var hechos = out.filter(function(x) { return x.estado === 'Completado'; });
  var res = {
    tipo: tipo, persona: persona.label, desde: desde, hasta: hasta, trabajos: out,
    resumen: {
      hechos: hechos.length, horas: _wa_redondea_(hechos.reduce(function(a, x) { return a + x.horas; }, 0)),
      sinCompletar: out.length - hechos.length
    }
  };
  if (tipo === 'c') {
    var docs = Object.keys(docsVistos).map(function(k) { return docsVistos[k]; });
    // cuántos trabajos hechos hay en cada situación de facturación
    res.resumen.sinFacturar = hechos.filter(function(x) { return !x.factura; }).length;
    res.resumen.facturados = hechos.filter(function(x) { return !!x.factura; }).length;
    res.resumen.pagados = hechos.filter(function(x) { return x.factura && x.factura.pago === 'pagada'; }).length;
    res.resumen.pendientesPago = hechos.filter(function(x) { return x.factura && x.factura.pago === 'pendiente'; }).length;
    res.documentos = docs.sort(function(a, b) { return a.fecha < b.fecha ? 1 : -1; }).map(function(d) { return { id: d.id, tipo: d.tipo, total: d.total, pago: d.pago, fecha: d.fecha }; });
    res.resumen.importePendiente = Math.round(docs.filter(function(d) { return d.pago === 'pendiente'; }).reduce(function(a, d) { return a + d.total; }, 0) * 100) / 100;
  }
  return res;
}
