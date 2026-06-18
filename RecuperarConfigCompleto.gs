// ============================================================
// RecuperarConfigCompleto.gs — Reconstrucción completa de CONFIG
//
// REGLAS:
//   - NUNCA borra carpetas, archivos ni formularios.
//   - NUNCA crea recursos si ya existen.
//   - Solo localiza y escribe IDs en CONFIG.
//   - Genera DIAGNOSTICO_CONFIG_COMPLETO y CONFIG_VALIDADO.
// ============================================================

// ── IDs conocidos del proyecto (Install.gs legacy, como pistas) ──
var RCC_KNOWN_ = {
  FORM_CLIENTES_ID:  '1HO2xa0hdbvmBNOM8FwDhfSf41A496djIDYRJhyYKOio',
  FORM_OPERARIOS_ID: '1YhQn0mXIaMKaBCu0TNk_xyQgYO4ymnuLcPLTG9E8Hec'
};

// ── Todos los campos CONFIG que este módulo gestiona ─────────────
var RCC_CAMPOS_ = {
  // ── Carpetas ─────────────────────────────────
  CARPETAS: [
    { key: 'carpeta_facturas_id',     tipo: 'carpeta',  buscar: ['FACTURAS'],
      altBuscar: ['Facturas'],        label: 'Facturas'     },
    { key: 'carpeta_trabajos_id',     tipo: 'carpeta',  buscar: ['TRABAJOS'],
      altBuscar: ['Trabajos'],        label: 'Trabajos'     },
    { key: 'carpeta_presupuestos_id', tipo: 'carpeta',  buscar: ['PRESUPUESTOS'],
      altBuscar: ['Presupuestos'],    label: 'Presupuestos' },
    { key: 'carpeta_proformas_id',    tipo: 'carpeta',  buscar: ['PROFORMAS'],
      altBuscar: ['Proformas'],       label: 'Proformas'    },
    { key: 'carpeta_informes_id',     tipo: 'carpeta',  buscar: ['INFORMES'],
      altBuscar: ['Informes'],        label: 'Informes'     },
    { key: 'gastos_folder_id',        tipo: 'carpeta',
      buscar: ['Álvaro Gestión — Tickets y Facturas de Gastos', 'GASTOS'],
      altBuscar: ['Gastos', 'Tickets'],  label: 'Gastos (carpeta tickets)' }
  ],
  // ── Formularios ──────────────────────────────
  FORMS: [
    { keyId:  'form_clientes_id',    keyUrl: 'form_clientes_url',
      pista:  RCC_KNOWN_.FORM_CLIENTES_ID,
      titulo: 'cliente', label: 'Formulario Clientes' },
    { keyId:  'form_operarios_id',   keyUrl: 'form_operarios_url',
      pista:  RCC_KNOWN_.FORM_OPERARIOS_ID,
      titulo: 'operario', label: 'Formulario Operarios' },
    { keyId:  'form_trabajos_id',    keyUrl: 'form_trabajos_url',
      pista:  null,
      titulo: 'trabajo', label: 'Formulario Trabajos' },
    { keyId:  'gastos_form_id',      keyUrl: 'gastos_form_url',
      pista:  null,
      titulo: 'gasto', label: 'Formulario Gastos' }
  ]
};

// ════════════════════════════════════════════════════════════════
// FUNCIÓN PRINCIPAL
// ════════════════════════════════════════════════════════════════

function reconstruirConfigCompleto() {
  var ss  = getTargetSS_();
  var ui  = SpreadsheetApp.getUi();
  var ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  var NCOLS = 6;

  var log = [];   // filas para DIAGNOSTICO_CONFIG_COMPLETO
  var resumen = { recuperados: 0, noEncontrados: [], dudosos: [] };

  function F() {
    var args = Array.prototype.slice.call(arguments);
    while (args.length < NCOLS) args.push('');
    return args.slice(0, NCOLS);
  }

  // ── Cabecera ────────────────────────────────────────────────
  log.push(F('DIAGNÓSTICO CONFIG COMPLETO — ' + ts));
  log.push(F('Objetivo: localizar en Drive y Forms los recursos del proyecto Álvaro y rellenar CONFIG'));
  log.push(F(''));

  // ══════════════════════════════════════════════════════════════
  // BLOQUE 1: Estado actual de CONFIG
  // ══════════════════════════════════════════════════════════════
  log.push(F('══ 1. ESTADO ACTUAL DE CONFIG (antes de operación) ══════════════'));
  log.push(F('Campo CONFIG', 'Tipo', 'Valor actual', 'Estado inicial'));

  var estadoAntes = {};

  RCC_CAMPOS_.CARPETAS.forEach(function(c) {
    var val = getConfig(c.key) || '';
    estadoAntes[c.key] = val;
    log.push(F(c.key, 'carpeta', val || '(vacío)', val ? '✅ Tiene valor' : '❌ VACÍO'));
  });
  RCC_CAMPOS_.FORMS.forEach(function(f) {
    var valId  = getConfig(f.keyId)  || '';
    var valUrl = getConfig(f.keyUrl) || '';
    estadoAntes[f.keyId]  = valId;
    estadoAntes[f.keyUrl] = valUrl;
    log.push(F(f.keyId,  'formulario ID',  valId  || '(vacío)', valId  ? '✅ Tiene valor' : '❌ VACÍO'));
    log.push(F(f.keyUrl, 'formulario URL', valUrl || '(vacío)', valUrl ? '✅ Tiene valor' : '❌ VACÍO'));
  });
  log.push(F(''));

  // ══════════════════════════════════════════════════════════════
  // BLOQUE 2: Recuperar carpetas
  // ══════════════════════════════════════════════════════════════
  log.push(F('══ 2. BÚSQUEDA Y RECUPERACIÓN DE CARPETAS ═══════════════════════'));
  log.push(F('Campo CONFIG', 'Valor anterior', 'Valor recuperado', 'Estado', 'Origen encontrado'));

  RCC_CAMPOS_.CARPETAS.forEach(function(carp) {
    var anterior = estadoAntes[carp.key] || '';

    // Si ya está configurado y accesible → mantener
    if (anterior) {
      try {
        var ef = DriveApp.getFolderById(anterior);
        log.push(F(carp.key, anterior, anterior, '✅ Ya configurado — sin cambios',
          ef.getName() + ' (' + _rcc_buildPath_(ef) + ')'));
        return;
      } catch (_) {
        log.push(F(carp.key, anterior, '', '⚠️ ID guardada pero no accesible — rebuscando', ''));
      }
    }

    // Buscar por nombres (en orden de prioridad)
    var encontrada = null;
    var origenDesc = '';

    for (var ni = 0; ni < carp.buscar.length && !encontrada; ni++) {
      var nombre = carp.buscar[ni];
      var exactas = _rcc_buscarCarpetas_(nombre);

      if (exactas.length === 1) {
        encontrada = exactas[0];
        origenDesc = 'Búsqueda exacta "' + nombre + '"';
      } else if (exactas.length > 1) {
        // Preferir la que tiene subcarpetas de año
        var conAnios = exactas.filter(_rcc_tieneAnios_);
        encontrada = conAnios.length > 0 ? conAnios[0] : exactas[0];
        origenDesc = 'Exacta "' + nombre + '" (' + exactas.length + ' encontradas, elegida ' +
          (conAnios.length > 0 ? 'por estructura año' : 'la primera') + ')';
        // Listar todas como referencia
        exactas.forEach(function(f) {
          var marca = f.getId() === encontrada.getId() ? ' ← ELEGIDA' : '';
          log.push(F('  → ' + f.getName(), f.getId(), _rcc_buildPath_(f), '', marca));
        });
      }
    }

    // Si no encontrada exacta → buscar con contains
    if (!encontrada) {
      var altTerminos = (carp.altBuscar || []).concat(carp.buscar);
      for (var ai = 0; ai < altTerminos.length && !encontrada; ai++) {
        var prefijo = altTerminos[ai].substring(0, 6);
        var amplia = _rcc_searchFolders_('title contains "' + prefijo + '"');
        if (amplia.length > 0) {
          encontrada = amplia[0];
          origenDesc = 'Búsqueda amplia contains "' + prefijo + '" (' + amplia.length + ' candidata/s)';
          amplia.slice(1).forEach(function(f) {
            log.push(F('  candidata: ' + f.getName(), f.getId(), _rcc_buildPath_(f), '', '(no elegida)'));
          });
        }
      }
    }

    if (encontrada) {
      var nuevoId = encontrada.getId();
      setConfig(carp.key, nuevoId);
      var ruta = _rcc_buildPath_(encontrada);
      log.push(F(carp.key, anterior || '(vacío)', nuevoId,
        '✅ Asignado', encontrada.getName() + ' | ' + ruta + ' | ' + origenDesc));
      resumen.recuperados++;
      Logger.log('[RCC] ✅ ' + carp.key + ' → ' + nuevoId);
    } else {
      log.push(F(carp.key, anterior || '(vacío)', '(no encontrado)',
        '❌ No encontrada', 'Pega el ID manualmente en CONFIG → ' + carp.key));
      resumen.noEncontrados.push(carp.label);
      Logger.log('[RCC] ❌ ' + carp.key + ' — sin resultado');
    }
  });
  log.push(F(''));

  // ══════════════════════════════════════════════════════════════
  // BLOQUE 3: Recuperar formularios
  // ══════════════════════════════════════════════════════════════
  log.push(F('══ 3. BÚSQUEDA Y RECUPERACIÓN DE FORMULARIOS ════════════════════'));
  log.push(F('Campo CONFIG', 'Valor anterior', 'Valor recuperado', 'Estado', 'Origen encontrado'));

  // Obtener todos los forms accesibles en Drive (máx 200 para no agotar tiempo)
  var todosLosFormIds = _rcc_listarTodosLosFormIds_();
  var ssId = ss.getId();

  // Para cada form conocido, clasificar por: pista → mismo-spreadsheet-destino → título
  var formsEncontrados = {}; // keyId → {id, url, origen}

  // Paso A: Verificar pistas conocidas
  RCC_CAMPOS_.FORMS.forEach(function(fCfg) {
    if (!fCfg.pista) return;
    try {
      var f = FormApp.openById(fCfg.pista);
      formsEncontrados[fCfg.keyId] = {
        id:     fCfg.pista,
        url:    f.getPublishedUrl(),
        titulo: f.getTitle(),
        origen: 'ID conocido (Install.gs)'
      };
      Logger.log('[RCC] Pista OK: ' + fCfg.keyId + ' → ' + fCfg.pista);
    } catch (_) {
      Logger.log('[RCC] Pista inaccesible: ' + fCfg.pista + ' para ' + fCfg.keyId);
    }
  });

  // Paso B: Clasificar forms encontrados en Drive por destino y título
  todosLosFormIds.forEach(function(info) {
    var fId    = info.id;
    var fTitle = (info.title || '').toLowerCase();
    var destId = info.destId || '';

    RCC_CAMPOS_.FORMS.forEach(function(fCfg) {
      if (formsEncontrados[fCfg.keyId]) return; // ya resuelto

      var tituloCoincide = fTitle.indexOf(fCfg.titulo) >= 0;
      var destinoCoincide = destId === ssId;

      if (destinoCoincide && tituloCoincide) {
        formsEncontrados[fCfg.keyId] = {
          id:     fId,
          url:    info.url,
          titulo: info.title,
          origen: 'Drive (destino=esta hoja + título "' + fCfg.titulo + '")'
        };
        Logger.log('[RCC] Por destino+título: ' + fCfg.keyId + ' → ' + fId);
      } else if (tituloCoincide) {
        // Solo título → candidato dudoso
        if (!formsEncontrados[fCfg.keyId + '_dudoso']) {
          formsEncontrados[fCfg.keyId + '_dudoso'] = {
            id:     fId,
            url:    info.url,
            titulo: info.title,
            origen: 'Drive (solo título "' + fCfg.titulo + '", sin destino verificado)'
          };
        }
      }
    });
  });

  // Paso C: Si no se encontró con certeza, usar dudoso como fallback
  RCC_CAMPOS_.FORMS.forEach(function(fCfg) {
    if (!formsEncontrados[fCfg.keyId] && formsEncontrados[fCfg.keyId + '_dudoso']) {
      formsEncontrados[fCfg.keyId] = formsEncontrados[fCfg.keyId + '_dudoso'];
      formsEncontrados[fCfg.keyId].origen += ' — DUDOSO (verificar)';
      resumen.dudosos.push(fCfg.label);
    }
  });

  // Paso D: Escribir en CONFIG y log
  RCC_CAMPOS_.FORMS.forEach(function(fCfg) {
    var anteriorId  = estadoAntes[fCfg.keyId]  || '';
    var anteriorUrl = estadoAntes[fCfg.keyUrl] || '';

    // Si ya tenía ID y es accesible → mantener
    if (anteriorId) {
      try {
        var fExist = FormApp.openById(anteriorId);
        log.push(F(fCfg.keyId, anteriorId, anteriorId,
          '✅ Ya configurado — sin cambios', fExist.getTitle() + ' (accesible)'));
        log.push(F(fCfg.keyUrl, anteriorUrl, anteriorUrl, '✅ Ya configurado', ''));
        return;
      } catch (_) {
        log.push(F(fCfg.keyId, anteriorId, '', '⚠️ ID guardada pero no accesible — rebuscando', ''));
      }
    }

    var hallado = formsEncontrados[fCfg.keyId];
    if (hallado) {
      setConfig(fCfg.keyId,  hallado.id);
      setConfig(fCfg.keyUrl, hallado.url || '');
      log.push(F(fCfg.keyId,  anteriorId  || '(vacío)', hallado.id,
        hallado.origen.indexOf('DUDOSO') >= 0 ? '⚠️ Asignado (DUDOSO)' : '✅ Asignado',
        hallado.titulo + ' | ' + hallado.origen));
      log.push(F(fCfg.keyUrl, anteriorUrl || '(vacío)', hallado.url || '',
        '✅ URL asignada', ''));
      if (hallado.origen.indexOf('DUDOSO') < 0) resumen.recuperados++;
      Logger.log('[RCC] ✅ ' + fCfg.keyId + ' → ' + hallado.id);
    } else {
      log.push(F(fCfg.keyId, anteriorId || '(vacío)', '(no encontrado)',
        '❌ No encontrado',
        'Crea el formulario con installGastosModule() / setupForms() o pega el ID manualmente'));
      log.push(F(fCfg.keyUrl, anteriorUrl || '(vacío)', '(no encontrado)', '❌', ''));
      resumen.noEncontrados.push(fCfg.label);
      Logger.log('[RCC] ❌ ' + fCfg.keyId + ' — no encontrado');
    }
  });
  log.push(F(''));

  // ══════════════════════════════════════════════════════════════
  // BLOQUE 4: Verificar accesibilidad de todo lo configurado
  // ══════════════════════════════════════════════════════════════
  log.push(F('══ 4. VERIFICACIÓN DE ACCESIBILIDAD ═════════════════════════════'));
  log.push(F('Campo CONFIG', 'ID/Valor', 'Accesible', 'Nombre del recurso', 'Estado'));

  RCC_CAMPOS_.CARPETAS.forEach(function(carp) {
    var val = getConfig(carp.key) || '';
    if (!val) {
      log.push(F(carp.key, '(vacío)', '—', '—', '❌ Sin configurar'));
      return;
    }
    try {
      var f = DriveApp.getFolderById(val);
      log.push(F(carp.key, val, '✅ Sí', f.getName() + ' | ' + _rcc_buildPath_(f), '✅ OK'));
    } catch (e) {
      log.push(F(carp.key, val, '❌ No', 'Error: ' + e.message, '❌ ID inválida'));
    }
  });
  RCC_CAMPOS_.FORMS.forEach(function(fCfg) {
    var val = getConfig(fCfg.keyId) || '';
    if (!val) {
      log.push(F(fCfg.keyId, '(vacío)', '—', '—', '❌ Sin configurar'));
      return;
    }
    try {
      var form = FormApp.openById(val);
      log.push(F(fCfg.keyId, val, '✅ Sí', form.getTitle(), '✅ OK'));
    } catch (e) {
      log.push(F(fCfg.keyId, val, '❌ No', 'Error: ' + e.message, '❌ ID inválida'));
    }
  });
  log.push(F(''));

  // ══════════════════════════════════════════════════════════════
  // BLOQUE 5: Resumen
  // ══════════════════════════════════════════════════════════════
  log.push(F('══ 5. RESUMEN ═══════════════════════════════════════════════════'));
  log.push(F('Recuperados/asignados: ' + resumen.recuperados));
  if (resumen.noEncontrados.length > 0)
    log.push(F('❌ No encontrados: ' + resumen.noEncontrados.join(', ')));
  if (resumen.dudosos.length > 0)
    log.push(F('⚠️ Dudosos (verificar): ' + resumen.dudosos.join(', ')));
  log.push(F(''));

  // ══════════════════════════════════════════════════════════════
  // BLOQUE 6: Verificar flujos de generación / subida (solo checks, sin generar)
  // ══════════════════════════════════════════════════════════════
  log.push(F('══ 6. VERIFICACIÓN DE FLUJOS ════════════════════════════════════'));
  log.push(F('Flujo', 'Requisitos', 'Estado'));

  var flujos = _rcc_verificarFlujos_(ss);
  flujos.forEach(function(fl) {
    log.push(F(fl.nombre, fl.requisitos, fl.estado));
  });
  log.push(F(''));

  // ── Crear DIAGNOSTICO_CONFIG_COMPLETO ───────────────────────
  _rcc_crearHoja_(ss, 'DIAGNOSTICO_CONFIG_COMPLETO', log, NCOLS, ts);

  // ── Crear CONFIG_VALIDADO ────────────────────────────────────
  _rcc_crearConfigValidado_(ss, flujos, ts);

  // ── Alerta final ─────────────────────────────────────────────
  var factOk    = !!getConfig('carpeta_facturas_id');
  var todoOk    = resumen.noEncontrados.length === 0 && resumen.dudosos.length === 0;
  var veredicto = factOk ? '✅ Factura lista para generarse' : '❌ Sigue faltando configuración';

  var msg = veredicto + '\n\n';
  msg += '✅ Recuperados/asignados: ' + resumen.recuperados + '\n';
  if (resumen.noEncontrados.length > 0)
    msg += '❌ No encontrados: ' + resumen.noEncontrados.join(', ') + '\n';
  if (resumen.dudosos.length > 0)
    msg += '⚠️ Dudosos (verificar): ' + resumen.dudosos.join(', ') + '\n';
  msg += '\nVer:\n• DIAGNOSTICO_CONFIG_COMPLETO — informe detallado\n• CONFIG_VALIDADO — resumen ejecutivo';

  logEvent('INFO', 'reconstruirConfigCompleto',
    'recuperados=' + resumen.recuperados +
    ' noEncontrados=' + resumen.noEncontrados.length +
    ' dudosos=' + resumen.dudosos.length, '');

  ui.alert('⚙️ CONFIG — Reconstrucción completa', msg, ui.ButtonSet.OK);
}

// ════════════════════════════════════════════════════════════════
// VERIFICACIÓN DE FLUJOS
// ════════════════════════════════════════════════════════════════

function _rcc_verificarFlujos_(ss) {
  var flujos = [];

  // ── Facturas ─────────────────────────────────────────────────
  var carpFac  = getConfig('carpeta_facturas_id') || '';
  var empresa  = String(getConfig('empresa_nombre') || '').trim();
  var cif      = String(getConfig('empresa_cif')    || '').trim();
  var facOk    = carpFac && empresa && empresa.toLowerCase() !== 'mi empresa' &&
                 cif    && cif.toLowerCase()    !== 'a00000000';
  flujos.push({
    nombre:    'Generación de facturas',
    requisitos: 'carpeta_facturas_id + empresa_nombre + empresa_cif (no prueba)',
    estado:    facOk ? '✅ Listo' :
      '❌ Falta: ' + [
        !carpFac  ? 'carpeta_facturas_id' : '',
        (!empresa || empresa.toLowerCase() === 'mi empresa') ? 'empresa_nombre real' : '',
        (!cif     || cif.toLowerCase()    === 'a00000000')  ? 'empresa_cif real'    : ''
      ].filter(Boolean).join(', ')
  });

  // ── Presupuestos ──────────────────────────────────────────────
  var carpPre = getConfig('carpeta_presupuestos_id') || getConfig('carpeta_facturas_id') || '';
  flujos.push({
    nombre:    'Generación de presupuestos',
    requisitos: 'carpeta_presupuestos_id (o fallback a carpeta_facturas_id)',
    estado:    carpPre ? '✅ Listo (' + carpPre + ')' : '❌ Falta carpeta_presupuestos_id y carpeta_facturas_id'
  });

  // ── Proformas ─────────────────────────────────────────────────
  var carpPro = getConfig('carpeta_proformas_id') || getConfig('carpeta_facturas_id') || '';
  flujos.push({
    nombre:    'Generación de proformas',
    requisitos: 'carpeta_proformas_id (o fallback a carpeta_facturas_id)',
    estado:    carpPro ? '✅ Listo (' + carpPro + ')' : '❌ Falta carpeta_proformas_id y carpeta_facturas_id'
  });

  // ── Trabajos (formulario de subida) ──────────────────────────
  var formTrab = getConfig('form_trabajos_id') || '';
  var trabOk   = false;
  if (formTrab) {
    try { FormApp.openById(formTrab); trabOk = true; } catch (_) {}
  }
  flujos.push({
    nombre:    'Subida de trabajos (formulario)',
    requisitos: 'form_trabajos_id accesible',
    estado:    trabOk ? '✅ Listo' :
      (formTrab ? '❌ ID configurado pero no accesible: ' + formTrab :
                  '❌ Falta form_trabajos_id — ejecuta setupForms() o crea el formulario')
  });

  // ── Gastos (formulario + carpeta) ─────────────────────────────
  var formGas   = getConfig('gastos_form_id')   || '';
  var carpGas   = getConfig('gastos_folder_id') || '';
  var gasFormOk = false;
  var gasCarpOk = false;
  if (formGas) { try { FormApp.openById(formGas);       gasFormOk = true; } catch (_) {} }
  if (carpGas) { try { DriveApp.getFolderById(carpGas); gasCarpOk = true; } catch (_) {} }
  flujos.push({
    nombre:    'Subida de gastos',
    requisitos: 'gastos_form_id accesible + gastos_folder_id accesible',
    estado:    (gasFormOk && gasCarpOk) ? '✅ Listo' :
      '❌ Falta: ' + [
        !gasFormOk ? 'gastos_form_id (ejecuta installGastosModule())' : '',
        !gasCarpOk ? 'gastos_folder_id (carpeta de tickets)'           : ''
      ].filter(Boolean).join(' + ')
  });

  return flujos;
}

// ════════════════════════════════════════════════════════════════
// HOJA CONFIG_VALIDADO
// ════════════════════════════════════════════════════════════════

function _rcc_crearConfigValidado_(ss, flujos, ts) {
  var sh = getOrCreateSheet('CONFIG_VALIDADO');
  sh.clearContents();
  sh.clearFormats();

  var filas = [];
  filas.push(['CONFIG_VALIDADO — ' + ts, '', '', '']);
  filas.push(['Campo CONFIG', 'Valor', 'Accesible', 'Estado']);

  var todos = RCC_CAMPOS_.CARPETAS.map(function(c) {
    return { key: c.key, label: c.label, tipo: 'carpeta' };
  }).concat(
    RCC_CAMPOS_.FORMS.map(function(f) {
      return { key: f.keyId, label: f.label + ' (ID)', tipo: 'formulario' };
    }).concat(
      RCC_CAMPOS_.FORMS.map(function(f) {
        return { key: f.keyUrl, label: f.label + ' (URL)', tipo: 'url', noVerify: true };
      })
    )
  );

  todos.forEach(function(t) {
    var val = getConfig(t.key) || '';
    if (!val) {
      filas.push([t.key, '(vacío)', '—', '❌ Sin configurar']);
      return;
    }
    if (t.noVerify) {
      filas.push([t.key, val, '—', '✅ Guardado']);
      return;
    }
    try {
      var nombre = '';
      if (t.tipo === 'carpeta') {
        nombre = DriveApp.getFolderById(val).getName();
      } else {
        nombre = FormApp.openById(val).getTitle();
      }
      filas.push([t.key, val, '✅ ' + nombre, '✅ OK']);
    } catch (e) {
      filas.push([t.key, val, '❌ ' + e.message, '❌ Inválido']);
    }
  });

  filas.push(['', '', '', '']);
  filas.push(['══ FLUJOS DE TRABAJO ══', '', '', '']);
  flujos.forEach(function(fl) {
    filas.push([fl.nombre, fl.estado, '', '']);
  });

  var NCOLS = 4;
  sh.getRange(1, 1, filas.length, NCOLS).setValues(filas);
  sh.getRange(1, 1, 1, NCOLS).setBackground('#0d652d').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(2, 1, 1, NCOLS).setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold');

  for (var r = 2; r < filas.length; r++) {
    var txt = filas[r].join('');
    var bg;
    if (txt.indexOf('✅') >= 0) bg = '#e6f4ea';
    else if (txt.indexOf('❌') >= 0) bg = '#fce8e6';
    else if (txt.indexOf('⚠️') >= 0) bg = '#fff8e1';
    else if (txt.indexOf('══') >= 0) bg = '#e8eaed';
    if (bg) sh.getRange(r + 1, 1, 1, NCOLS).setBackground(bg);
  }
  sh.autoResizeColumns(1, NCOLS);
}

// ════════════════════════════════════════════════════════════════
// HELPERS PRIVADOS
// ════════════════════════════════════════════════════════════════

function _rcc_buscarCarpetas_(nombre) {
  var results = [];
  try {
    var it = DriveApp.getFoldersByName(nombre);
    while (it.hasNext()) results.push(it.next());
  } catch (e) {
    Logger.log('[RCC] Error buscando carpeta "' + nombre + '": ' + e.message);
  }
  return results;
}

function _rcc_searchFolders_(query) {
  var results = [];
  try {
    var it = DriveApp.searchFolders(query);
    var n = 0;
    while (it.hasNext() && n < 20) { results.push(it.next()); n++; }
  } catch (e) {
    Logger.log('[RCC] Error searchFolders "' + query + '": ' + e.message);
  }
  return results;
}

function _rcc_tieneAnios_(folder) {
  var anos = ['2024', '2025', '2026', '2027', '2028'];
  for (var i = 0; i < anos.length; i++) {
    try { if (folder.getFoldersByName(anos[i]).hasNext()) return true; } catch (_) {}
  }
  return false;
}

function _rcc_buildPath_(folder) {
  try {
    var parts   = [folder.getName()];
    var current = folder;
    var rootId  = DriveApp.getRootFolder().getId();
    for (var i = 0; i < 5; i++) {
      var parents = current.getParents();
      if (!parents.hasNext()) break;
      var p = parents.next();
      if (p.getId() === rootId) { parts.unshift('Mi unidad'); break; }
      parts.unshift(p.getName());
      current = p;
    }
    return parts.join(' / ');
  } catch (_) {
    try { return folder.getName(); } catch (_2) { return '—'; }
  }
}

/**
 * Lista todos los Google Forms accesibles en Drive del usuario.
 * Devuelve array de {id, title, url, destId}.
 * Máx 150 formularios para no agotar tiempo de ejecución.
 */
function _rcc_listarTodosLosFormIds_() {
  var infos = [];
  try {
    var it = DriveApp.searchFiles('mimeType = "application/vnd.google-apps.form"');
    var n  = 0;
    while (it.hasNext() && n < 150) {
      var file = it.next();
      var info = { id: file.getId(), title: file.getName(), url: '', destId: '' };
      try {
        var form = FormApp.openById(file.getId());
        info.url    = form.getPublishedUrl() || '';
        info.destId = String(form.getDestinationId() || '');
        info.title  = form.getTitle() || info.title;
      } catch (_) {}
      infos.push(info);
      n++;
    }
    Logger.log('[RCC] Forms encontrados en Drive: ' + infos.length);
  } catch (e) {
    Logger.log('[RCC] Error listando forms: ' + e.message);
  }
  return infos;
}

function _rcc_crearHoja_(ss, nombre, log, NCOLS, ts) {
  var sh = getOrCreateSheet(nombre);
  sh.clearContents();
  sh.clearFormats();

  var datos = log.map(function(row) {
    var r = row.slice();
    while (r.length < NCOLS) r.push('');
    return r.slice(0, NCOLS);
  });
  sh.getRange(1, 1, datos.length, NCOLS).setValues(datos);
  sh.getRange(1, 1, 1, NCOLS).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(2, 1, 1, NCOLS).setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');

  for (var r = 2; r < datos.length; r++) {
    var txt = datos[r].join('');
    var bg;
    if (txt.indexOf('✅') >= 0) bg = '#e6f4ea';
    else if (txt.indexOf('❌') >= 0) bg = '#fce8e6';
    else if (txt.indexOf('⚠️') >= 0) bg = '#fff8e1';
    else if (txt.indexOf('══') >= 0) bg = '#e8eaed';
    if (bg) sh.getRange(r + 1, 1, 1, NCOLS).setBackground(bg);
  }
  sh.autoResizeColumns(1, NCOLS);
  ss.setActiveSheet(sh);
}

// ════════════════════════════════════════════════════════════════
// FORMULARIOS — Desactivar "Editar respuesta"
//
// Por defecto Google Forms envía al usuario un enlace "Editar respuesta"
// tras cada envío. Si el usuario lo abre, ve el relleno anterior.
// Esta función desactiva esa opción en todos los formularios del sistema.
// ════════════════════════════════════════════════════════════════

function desactivarEdicionRespuestasFormularios() {
  var ui  = SpreadsheetApp.getUi();
  var log = [];
  var ok  = 0;
  var err = 0;

  var forms = [
    { key: 'form_clientes_id',  label: 'Clientes'  },
    { key: 'form_operarios_id', label: 'Operarios' },
    { key: 'form_trabajos_id',  label: 'Trabajos'  },
    { key: 'gastos_form_id',    label: 'Gastos'    }
  ];

  forms.forEach(function(f) {
    var id = getConfig(f.key) || '';
    if (!id) {
      log.push('⚪ ' + f.label + ': sin configurar en CONFIG (' + f.key + ')');
      return;
    }
    try {
      var form = FormApp.openById(id);
      var titulo = form.getTitle();

      // Desactivar edición de respuestas (esto elimina el enlace "Editar tu respuesta")
      form.setAllowResponseEdits(false);

      // Desactivar confirmación con resumen de respuestas (evita que se muestren las anteriores)
      form.setShowLinkToRespondAgain(true);

      log.push('✅ ' + f.label + ' (' + titulo + '): edición desactivada');
      ok++;
      Logger.log('[FORMS] ✅ ' + f.label + ' → setAllowResponseEdits(false)');
    } catch (e) {
      log.push('❌ ' + f.label + ' (ID: ' + id + '): ' + e.message);
      err++;
      Logger.log('[FORMS] ❌ ' + f.label + ': ' + e.message);
    }
  });

  var msg = '📋 FORMULARIOS — Edición de respuestas\n\n' +
    log.join('\n') + '\n\n' +
    (ok  > 0 ? '✅ ' + ok  + ' formulario(s) actualizado(s)\n' : '') +
    (err > 0 ? '❌ ' + err + ' error(es)\n'                     : '') +
    '\nLos usuarios ya no verán sus respuestas anteriores al abrir el formulario.';

  ui.alert('📋 Formularios — Sin residuos', msg, ui.ButtonSet.OK);
}
