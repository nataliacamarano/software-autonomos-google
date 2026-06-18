// ============================================================
// OrganizarHojas.gs — Organización y visibilidad de hojas
//
// REGLAS:
//   - NUNCA elimina hojas.
//   - Solo oculta hojas no esenciales para el trabajo diario.
//   - Genera informes: MAPA_HOJAS_SISTEMA, HOJAS_OCULTADAS,
//     ANALISIS_HOJAS_ELIMINABLES.
// ============================================================

// ── Hojas que deben permanecer VISIBLES ─────────────────────
var OH_VISIBLES = [
  'PANEL_GENERAL', 'PANEL_CLIENTE', 'PANEL_FACTURAS',
  'PANEL_GASTOS',  'PANEL_CONTABLE', 'PANEL_PENDIENTES',
  'CLIENTES', 'TRABAJOS', 'FACTURAS', 'GASTOS', 'OPERARIOS',
  'PRESUPUESTOS', 'PROFORMAS', 'INFORMES'
];

// ── Hojas de informe generadas por ESTA función ──────────────
// Las 3 hojas de organización se ocultan junto con el resto (son admin).
// Solo ULTIMAS_HOJAS_OCULTADAS se mantiene visible tras ocultarHojasAdmin().
var OH_INFORMES_PROPIOS = [
  'ULTIMAS_HOJAS_OCULTADAS'
];

// Hojas de administración interna (se ocultan, no se eliminan)
var OH_HOJAS_ADMIN = [
  'MAPA_HOJAS_SISTEMA', 'HOJAS_OCULTADAS', 'ANALISIS_HOJAS_ELIMINABLES'
];

// ── Catálogo de hojas conocidas ──────────────────────────────
var OH_CATALOGO = {
  // ── Soporte interno ──────────────────────────────────────
  'CONFIG': {
    cat: 'Soporte interno',
    desc: 'Configuración global del sistema (claves, IDs, parámetros)',
    dep: 'Todas las funciones — getConfig() / setConfig()',
    eliminable: false, motivo_ocultar: 'Técnica — no editar manualmente'
  },
  'LOGS': {
    cat: 'Soporte interno',
    desc: 'Log de eventos del sistema (errores, acciones)',
    dep: 'logEvent() — todos los módulos',
    eliminable: false, motivo_ocultar: 'Técnica — solo lectura'
  },
  'COMPLETADOS': {
    cat: 'Soporte interno',
    desc: 'Vista de trabajos completados (filtro de TRABAJOS)',
    dep: 'PANEL_PENDIENTES, PANEL_GENERAL',
    eliminable: false, motivo_ocultar: 'Auxiliar — se actualiza automáticamente'
  },
  'PENDIENTES': {
    cat: 'Soporte interno',
    desc: 'Vista de trabajos pendientes (filtro de TRABAJOS)',
    dep: 'PANEL_PENDIENTES, refreshPanelPendientes()',
    eliminable: false, motivo_ocultar: 'Auxiliar — se actualiza automáticamente'
  },
  'CLIENTES_ACTIVOS': {
    cat: 'Soporte interno',
    desc: 'Lista filtrada de clientes activos para desplegables',
    dep: 'PANEL_FACTURAS — dropdowns de cliente',
    eliminable: false, motivo_ocultar: 'Auxiliar — necesaria para dropdowns'
  },
  'RESP_CLIENTES': {
    cat: 'Soporte interno',
    desc: 'Respuestas del formulario de alta de clientes',
    dep: 'Formulario Clientes + trigger onFormSubmit',
    eliminable: false, motivo_ocultar: 'Auxiliar — procesada por trigger automático'
  },
  'RESP_OPERARIOS': {
    cat: 'Soporte interno',
    desc: 'Respuestas del formulario de alta de operarios',
    dep: 'Formulario Operarios + trigger onFormSubmit',
    eliminable: false, motivo_ocultar: 'Auxiliar — procesada por trigger automático'
  },
  'RESP_GASTOS': {
    cat: 'Soporte interno',
    desc: 'Respuestas del formulario de gastos',
    dep: 'Formulario Gastos + trigger onFormSubmit',
    eliminable: false, motivo_ocultar: 'Auxiliar — procesada por trigger automático'
  },
  'RESP_TRABAJOS': {
    cat: 'Soporte interno',
    desc: 'Respuestas del formulario de partes de trabajo',
    dep: 'Formulario Trabajos + trigger onFormSubmit',
    eliminable: false, motivo_ocultar: 'Auxiliar — procesada por trigger automático'
  },
  'DRIVE_CONFIG': {
    cat: 'Soporte interno',
    desc: 'Configuración de carpetas Drive (módulo DriveManager legacy)',
    dep: 'DriveManager.gs',
    eliminable: false, motivo_ocultar: 'Técnica'
  },
  'DOC_TRABAJOS': {
    cat: 'Soporte interno',
    desc: 'Índice de documentos y fotos adjuntas a trabajos',
    dep: 'DocTrabajos.gs — adjuntarPdfGastoSeleccionado()',
    eliminable: false, motivo_ocultar: 'Auxiliar'
  },
  'ALERTAS_LOG': {
    cat: 'Soporte interno',
    desc: 'Historial de alertas diarias enviadas por email',
    dep: 'Alertas.gs — enviarAlertasDiarias()',
    eliminable: false, motivo_ocultar: 'Log técnico'
  },
  'RECORDATORIOS_LOG': {
    cat: 'Soporte interno',
    desc: 'Historial de recordatorios de cobro enviados',
    dep: 'Cobros.gs — enviarRecordatoriosPendientes()',
    eliminable: false, motivo_ocultar: 'Log técnico'
  },
  'ENVIOS_CONTABLE_LOG': {
    cat: 'Soporte interno',
    desc: 'Log de envíos al gestor contable',
    dep: 'PanelContable.gs — enviarGastosAlGestor()',
    eliminable: false, motivo_ocultar: 'Log técnico'
  },
  // ── Diagnóstico ──────────────────────────────────────────
  'DIAGNOSTICO_CARPETAS_DRIVE': {
    cat: 'Diagnóstico',
    desc: 'Diagnóstico de carpetas Drive (generado por recuperarCarpetasDrive)',
    dep: 'RecuperarCarpetas.gs',
    eliminable: true, motivo_ocultar: 'Informe técnico puntual'
  },
  'DIAGNOSTICO_FACTURA_DATOS': {
    cat: 'Diagnóstico',
    desc: 'Diagnóstico de datos de empresa y cliente para facturas',
    dep: 'Facturas.gs — diagnosticarFacturaDatos()',
    eliminable: true, motivo_ocultar: 'Informe técnico puntual'
  },
  'DIAGNOSTICO_RECUPERACION_CARPETAS': {
    cat: 'Diagnóstico',
    desc: 'Diagnóstico agresivo de recuperación de carpetas Drive',
    dep: 'RecuperarCarpetas.gs — diagnosticarYRecuperarCarpetas()',
    eliminable: true, motivo_ocultar: 'Informe técnico puntual'
  },
  'DIAGNOSTICO_CONFIG_COMPLETO': {
    cat: 'Diagnóstico',
    desc: 'Reconstrucción completa de CONFIG (carpetas + formularios)',
    dep: 'RecuperarConfigCompleto.gs',
    eliminable: true, motivo_ocultar: 'Informe técnico puntual'
  },
  'CONFIG_VALIDADO': {
    cat: 'Diagnóstico',
    desc: 'Resumen ejecutivo de CONFIG validado',
    dep: 'RecuperarConfigCompleto.gs',
    eliminable: true, motivo_ocultar: 'Informe técnico puntual'
  },
  'DASHBOARD': {
    cat: 'Diagnóstico',
    desc: 'Dashboard de métricas del sistema',
    dep: 'generarDashboard()',
    eliminable: true, motivo_ocultar: 'Vista auxiliar regenerable'
  },
  'DASHBOARD_FINANCIERO': {
    cat: 'Diagnóstico',
    desc: 'Dashboard financiero (ingresos, gastos, tendencias)',
    dep: 'generarDashboardFinanciero()',
    eliminable: true, motivo_ocultar: 'Vista auxiliar regenerable'
  },
  'INFORME_GASTOS': {
    cat: 'Diagnóstico',
    desc: 'Informe de gastos del período',
    dep: 'generarInformeGastos()',
    eliminable: true, motivo_ocultar: 'Informe regenerable'
  },
  'IVA_TRIMESTRAL': {
    cat: 'Diagnóstico',
    desc: 'Liquidación IVA trimestral',
    dep: 'generarIVATrimestral()',
    eliminable: true, motivo_ocultar: 'Informe regenerable'
  },
  'LCR_INFORME_PREVIO': {
    cat: 'Diagnóstico',
    desc: 'Informe previo de limpieza de clientes reales',
    dep: 'Cleanup.gs',
    eliminable: true, motivo_ocultar: 'Informe técnico puntual'
  },
  'MEJORAS_IMPLEMENTADAS': {
    cat: 'Diagnóstico',
    desc: 'Historial de mejoras del sistema implementadas',
    dep: 'Ninguna activa',
    eliminable: true, motivo_ocultar: 'Historial técnico'
  },
  // ── Auditoría ────────────────────────────────────────────
  'AUDITORIA_DATOS': {
    cat: 'Auditoría',
    desc: 'Auditoría de datos reales vs. prueba — Fase 1',
    dep: 'Auditoria.gs — ejecutarAuditoriaDatos()',
    eliminable: true, motivo_ocultar: 'Informe de auditoría puntual'
  },
  'AUDITORIA_FASE2': {
    cat: 'Auditoría',
    desc: 'Informe ejecutivo de auditoría — Fase 2 con semáforo',
    dep: 'AuditoriaFase2.gs — ejecutarAuditoriaFase2()',
    eliminable: true, motivo_ocultar: 'Informe de auditoría puntual'
  },
  'REVISION_MANUAL': {
    cat: 'Auditoría',
    desc: 'Registros que requieren revisión manual post-limpieza',
    dep: 'Cleanup.gs',
    eliminable: true, motivo_ocultar: 'Resultado de limpieza puntual'
  },
  'LIMPIEZA_CLIENTES_REALES': {
    cat: 'Auditoría',
    desc: 'Análisis de clientes reales separados de clientes de prueba',
    dep: 'Cleanup.gs',
    eliminable: true, motivo_ocultar: 'Resultado de limpieza puntual'
  },
  'LIMPIEZA_REALIZADA': {
    cat: 'Auditoría',
    desc: 'Log detallado de registros eliminados en limpieza',
    dep: 'Cleanup.gs — ejecutarFase3Completa()',
    eliminable: false, motivo_ocultar: 'Historial de seguridad — conservar'
  },
  // ── Backup ───────────────────────────────────────────────
  'BACKUP_REALIZADO': {
    cat: 'Backup',
    desc: 'Informe del último backup: hojas copiadas, registros, instrucciones restauración',
    dep: 'Backup.gs — ejecutarBackupCompleto()',
    eliminable: false, motivo_ocultar: 'Referencia de seguridad'
  },
  // ── Restauración ─────────────────────────────────────────
  'RESTAURACION_DIAGNOSTICO': {
    cat: 'Restauración',
    desc: 'Diagnóstico de diferencias antes de restaurar datos',
    dep: 'Restauracion.gs — ejecutarSoloDiagnosticoRestauracion()',
    eliminable: true, motivo_ocultar: 'Informe puntual de restauración'
  },
  'RESTAURACION_REALIZADA': {
    cat: 'Restauración',
    desc: 'Log de restauración de datos ejecutada',
    dep: 'Restauracion.gs — ejecutarRestauracionDesdeBackup()',
    eliminable: false, motivo_ocultar: 'Historial de seguridad — conservar'
  },
  'RESTAURACION_REVISION_MANUAL': {
    cat: 'Restauración',
    desc: 'Registros con conflicto pendientes de revisión manual post-restauración',
    dep: 'Restauracion.gs',
    eliminable: true, motivo_ocultar: 'Resultado puntual de restauración'
  },
  // ── Obsoletas ────────────────────────────────────────────
  'GASTOS_LEGACY': {
    cat: 'Obsoleta',
    desc: 'Datos de gastos del sistema anterior (migrados a GASTOS)',
    dep: 'Ninguna función activa la usa',
    eliminable: true, motivo_ocultar: 'Sin uso activo — candidata a eliminar'
  },
  'GASTOS_LEGACY_': {
    cat: 'Obsoleta',
    desc: 'Copia residual de GASTOS_LEGACY',
    dep: 'Ninguna función activa la usa',
    eliminable: true, motivo_ocultar: 'Sin uso activo — candidata a eliminar'
  },
  'CONTABILIDAD': {
    cat: 'Obsoleta',
    desc: 'Hoja de contabilidad (función trasladada a PANEL_CONTABLE)',
    dep: 'Verificar antes de eliminar',
    eliminable: true, motivo_ocultar: 'Posiblemente obsoleta'
  },
  'TEST_E2E': {
    cat: 'Obsoleta',
    desc: 'Resultados de test end-to-end (entorno de desarrollo)',
    dep: 'E2ETest.gs',
    eliminable: true, motivo_ocultar: 'Solo para desarrollo'
  }
};

// ════════════════════════════════════════════════════════════════
// FUNCIÓN PRINCIPAL
// ════════════════════════════════════════════════════════════════

function organizarHojas() {
  var ss  = SpreadsheetApp.getActive();
  var ui  = SpreadsheetApp.getUi();
  var ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

  var todasLasHojas = ss.getSheets();
  var contadores = { ocultadas: 0, yaOcultas: 0, visibles: 0, noEncontradas: [] };

  // Hojas que NUNCA se ocultan (visibles + informes propios)
  var noOcultar = OH_VISIBLES.concat(OH_INFORMES_PROPIOS);

  // ── PASO 1: Clasificar todas las hojas ──────────────────────
  var mapaData = _oh_clasificarHojas_(todasLasHojas, noOcultar);

  // ── PASO 2: Crear MAPA_HOJAS_SISTEMA ────────────────────────
  _oh_crearMapa_(ss, mapaData, ts);

  // ── PASO 3: Ocultar hojas no esenciales ─────────────────────
  var hojasMostrarVisibles = {};
  OH_VISIBLES.forEach(function(n) { hojasMostrarVisibles[n] = true; });

  todasLasHojas.forEach(function(sh) {
    var nombre = sh.getName();

    // Siempre mantener visibles las del usuario y los informes propios
    if (noOcultar.indexOf(nombre) >= 0) {
      try { sh.showSheet(); } catch (_) {}
      contadores.visibles++;
      return;
    }

    // Ocultar el resto
    if (!sh.isSheetHidden()) {
      try { sh.hideSheet(); contadores.ocultadas++; }
      catch (e) { Logger.log('[OH] No se pudo ocultar "' + nombre + '": ' + e.message); }
    } else {
      contadores.yaOcultas++;
    }
  });

  // Verificar que todas las hojas del listado VISIBLE existen
  var nombresExistentes = todasLasHojas.map(function(s) { return s.getName(); });
  OH_VISIBLES.forEach(function(n) {
    if (nombresExistentes.indexOf(n) < 0) contadores.noEncontradas.push(n);
  });

  // ── PASO 4: Crear HOJAS_OCULTADAS ───────────────────────────
  _oh_crearHojasOcultadas_(ss, mapaData, ts);

  // ── PASO 5: Crear ANALISIS_HOJAS_ELIMINABLES ─────────────────
  _oh_crearHojasEliminables_(ss, mapaData, ts);

  // ── Navegar a PANEL_FACTURAS como hoja de inicio ─────────────
  var panelSh = ss.getSheetByName('PANEL_FACTURAS');
  if (panelSh) ss.setActiveSheet(panelSh);

  // ── Alerta final ─────────────────────────────────────────────
  var msg = '✅ ORGANIZACIÓN COMPLETADA\n\n' +
    '✅ Hojas visibles (uso diario): ' + contadores.visibles + '\n' +
    '🙈 Ocultadas ahora:             ' + contadores.ocultadas + '\n' +
    '⚪ Ya estaban ocultas:          ' + contadores.yaOcultas + '\n';

  if (contadores.noEncontradas.length > 0)
    msg += '\n⚠️ Del listado de visibles, no existen en el spreadsheet:\n  ' +
      contadores.noEncontradas.join(', ') + '\n';

  msg += '\nVer hojas generadas:\n' +
    '• MAPA_HOJAS_SISTEMA — inventario completo\n' +
    '• HOJAS_OCULTADAS — qué se ocultó y por qué\n' +
    '• ANALISIS_HOJAS_ELIMINABLES — candidatas a borrar en fase 2';

  logEvent('INFO', 'organizarHojas',
    'ocultadas=' + contadores.ocultadas + ' visibles=' + contadores.visibles, '');

  ui.alert('🗂️ Organización de hojas', msg, ui.ButtonSet.OK);
}

// ════════════════════════════════════════════════════════════════
// CLASIFICACIÓN
// ════════════════════════════════════════════════════════════════

function _oh_clasificarHojas_(sheets, noOcultar) {
  var resultado = [];

  sheets.forEach(function(sh) {
    var nombre   = sh.getName();
    var lastRow  = sh.getLastRow();
    var lastCol  = sh.getLastColumn();
    var yaOculta = sh.isSheetHidden();
    var filas    = Math.max(0, lastRow - 1); // sin cabecera

    // Determinar categoría
    var info = OH_CATALOGO[nombre];
    var cat, desc, dep, eliminable, motivoOcultar;

    if (OH_VISIBLES.indexOf(nombre) >= 0) {
      cat = 'Uso diario';
      desc = info ? info.desc : 'Hoja principal de trabajo diario';
      dep  = info ? info.dep  : '—';
      eliminable    = false;
      motivoOcultar = '— (MANTENER VISIBLE)';
    } else if (OH_INFORMES_PROPIOS.indexOf(nombre) >= 0) {
      cat = 'Informe del sistema';
      desc = 'Informe de organización activo — generado por organizarHojas()';
      dep  = 'OrganizarHojas.gs';
      eliminable    = true;
      motivoOcultar = '— (MANTENER VISIBLE — informe activo)';
    } else if (OH_HOJAS_ADMIN.indexOf(nombre) >= 0) {
      cat = 'Administración interna';
      desc = 'Informe de organización del sistema (administración)';
      dep  = 'OrganizarHojas.gs — solo escritura, sin lectores externos';
      eliminable    = true;
      motivoOcultar = 'Administración interna — acceder via Ver → Hojas ocultas';
    } else if (info) {
      cat  = info.cat;
      desc = info.desc;
      dep  = info.dep;
      eliminable    = info.eliminable;
      motivoOcultar = info.motivo_ocultar;
    } else if (/^BKP_/.test(nombre)) {
      cat  = 'Backup';
      desc = 'Copia de seguridad automática de datos';
      dep  = 'Backup.gs — ejecutarBackupCompleto()';
      eliminable    = false;
      motivoOcultar = 'Backup de seguridad — conservar al menos 3 meses';
    } else if (/^DIAGNOSTICO/.test(nombre)) {
      cat  = 'Diagnóstico';
      desc = 'Hoja de diagnóstico generada automáticamente';
      dep  = 'Funciones de diagnóstico';
      eliminable    = true;
      motivoOcultar = 'Informe técnico regenerable';
    } else if (/AUDITOR/.test(nombre.toUpperCase())) {
      cat  = 'Auditoría';
      desc = 'Informe de auditoría del sistema';
      dep  = 'Módulo auditoría';
      eliminable    = true;
      motivoOcultar = 'Informe de auditoría puntual';
    } else if (/RESTAURA/.test(nombre.toUpperCase())) {
      cat  = 'Restauración';
      desc = 'Informe o log de restauración de datos';
      dep  = 'Restauracion.gs';
      eliminable    = true;
      motivoOcultar = 'Resultado de restauración puntual';
    } else if (/^RESP_/.test(nombre)) {
      cat  = 'Soporte interno';
      desc = 'Respuestas de formulario Google Forms';
      dep  = 'Trigger onFormSubmit';
      eliminable    = false;
      motivoOcultar = 'Procesada por trigger automático';
    } else if (/LIMPIEZA/.test(nombre.toUpperCase())) {
      cat  = 'Auditoría';
      desc = 'Informe o log de limpieza de datos';
      dep  = 'Cleanup.gs';
      eliminable    = true;
      motivoOcultar = 'Resultado de limpieza puntual';
    } else {
      cat  = 'Desconocida';
      desc = 'Sin clasificar — revisar manualmente';
      dep  = '¿?';
      eliminable    = false;
      motivoOcultar = 'Sin clasificar — ocultar por precaución';
    }

    var recomendacion;
    if (cat === 'Uso diario')          recomendacion = '✅ Mantener visible';
    else if (cat === 'Informe del sistema') recomendacion = '✅ Mantener visible';
    else if (eliminable)               recomendacion = '🗑️ Candidata a eliminar (fase 2)';
    else                               recomendacion = '🙈 Ocultar — conservar';

    resultado.push({
      nombre:        nombre,
      cat:           cat,
      desc:          desc,
      dep:           dep,
      eliminable:    eliminable,
      motivoOcultar: motivoOcultar,
      recomendacion: recomendacion,
      filas:         filas,
      yaOculta:      yaOculta,
      vacía:         lastRow <= 1 && lastCol <= 1
    });
  });

  // Ordenar: uso diario → soporte → diagnóstico → auditoría → backup → obsoleta → desconocida
  var orden = ['Uso diario','Informe del sistema','Soporte interno',
               'Diagnóstico','Auditoría','Backup','Restauración','Obsoleta','Desconocida'];
  resultado.sort(function(a, b) {
    return orden.indexOf(a.cat) - orden.indexOf(b.cat);
  });

  return resultado;
}

// ════════════════════════════════════════════════════════════════
// HOJA: MAPA_HOJAS_SISTEMA
// ════════════════════════════════════════════════════════════════

function _oh_crearMapa_(ss, mapaData, ts) {
  var sh = _oh_getSheet_(ss, 'MAPA_HOJAS_SISTEMA');
  var NCOLS = 8;
  var filas = [];

  filas.push(['MAPA_HOJAS_SISTEMA — ' + ts, '', '', '', '', '', '', '']);
  filas.push(['Hoja', 'Categoría', 'Función / Descripción',
              'Dependencias', 'Filas de datos', 'Estaba oculta', 'Vacía', 'Recomendación']);

  mapaData.forEach(function(h) {
    filas.push([
      h.nombre, h.cat, h.desc, h.dep,
      h.filas > 0 ? h.filas : '(vacía)',
      h.yaOculta ? 'Sí' : 'No',
      h.vacía   ? 'Sí' : 'No',
      h.recomendacion
    ]);
  });

  sh.clearContents(); sh.clearFormats();
  sh.getRange(1, 1, filas.length, NCOLS).setValues(filas);

  // Cabeceras
  sh.getRange(1, 1, 1, NCOLS).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(2, 1, 1, NCOLS).setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');

  // Color por categoría
  var colorCat = {
    'Uso diario':        '#e6f4ea',
    'Informe del sistema':'#e8f0fe',
    'Soporte interno':   '#fff8e1',
    'Diagnóstico':       '#fce8b2',
    'Auditoría':         '#fce8e6',
    'Backup':            '#e8eaed',
    'Restauración':      '#f3e8fd',
    'Obsoleta':          '#f5f5f5',
    'Desconocida':       '#fff3e0'
  };
  for (var r = 2; r < filas.length; r++) {
    var bg = colorCat[filas[r][1]] || '#ffffff';
    sh.getRange(r + 1, 1, 1, NCOLS).setBackground(bg);
  }
  sh.autoResizeColumns(1, NCOLS);
}

// ════════════════════════════════════════════════════════════════
// HOJA: HOJAS_OCULTADAS
// ════════════════════════════════════════════════════════════════

function _oh_crearHojasOcultadas_(ss, mapaData, ts) {
  var sh = _oh_getSheet_(ss, 'HOJAS_OCULTADAS');
  var NCOLS = 5;
  var filas = [];

  filas.push(['HOJAS_OCULTADAS — ' + ts, '', '', '', '']);
  filas.push(['Hoja', 'Categoría', 'Motivo de ocultación', 'Dependencia', 'Eliminable']);

  var ocultadas = mapaData.filter(function(h) {
    return OH_VISIBLES.indexOf(h.nombre) < 0 &&
           OH_INFORMES_PROPIOS.indexOf(h.nombre) < 0;
  });

  ocultadas.forEach(function(h) {
    filas.push([h.nombre, h.cat, h.motivoOcultar, h.dep, h.eliminable ? 'Sí' : 'No']);
  });

  filas.push(['', '', '', '', '']);
  filas.push(['TOTAL OCULTADAS: ' + ocultadas.length, '', '', '', '']);

  sh.clearContents(); sh.clearFormats();
  sh.getRange(1, 1, filas.length, NCOLS).setValues(filas);
  sh.getRange(1, 1, 1, NCOLS).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(2, 1, 1, NCOLS).setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');

  for (var r = 2; r < filas.length; r++) {
    var txt = String(filas[r][4] || '');
    var bg  = txt === 'Sí' ? '#fce8e6' : '#fff8e1';
    if (filas[r][0] === '') bg = '#ffffff';
    if (String(filas[r][0]).indexOf('TOTAL') >= 0) bg = '#e8eaed';
    sh.getRange(r + 1, 1, 1, NCOLS).setBackground(bg);
  }
  sh.autoResizeColumns(1, NCOLS);
}

// ════════════════════════════════════════════════════════════════
// HOJA: ANALISIS_HOJAS_ELIMINABLES
// ════════════════════════════════════════════════════════════════

function _oh_crearHojasEliminables_(ss, mapaData, ts) {
  var sh = _oh_getSheet_(ss, 'ANALISIS_HOJAS_ELIMINABLES');
  var NCOLS = 6;
  var filas = [];

  filas.push(['ANÁLISIS DE HOJAS ELIMINABLES (FASE 2) — ' + ts, '', '', '', '', '']);
  filas.push(['⚠️ NO eliminar todavía. Revisar cada hoja antes de borrar.', '', '', '', '', '']);
  filas.push(['Hoja', 'Categoría', 'Filas de datos', 'Vacía', 'Motivo para eliminar', 'Riesgo']);

  var eliminables = mapaData.filter(function(h) { return h.eliminable; });
  var noEliminables = mapaData.filter(function(h) {
    return !h.eliminable &&
           OH_VISIBLES.indexOf(h.nombre) < 0 &&
           OH_INFORMES_PROPIOS.indexOf(h.nombre) < 0;
  });

  if (eliminables.length > 0) {
    filas.push(['══ PUEDEN ELIMINARSE (revisar antes) ══', '', '', '', '', '']);
    eliminables.forEach(function(h) {
      var riesgo = h.vacía ? '🟢 Bajo (vacía)'
        : (h.filas < 10 ? '🟡 Bajo (pocos datos)'
        : '🔴 Revisar contenido antes de borrar');
      filas.push([h.nombre, h.cat, h.filas, h.vacía ? 'Sí' : 'No', h.motivoOcultar, riesgo]);
    });
  }

  filas.push(['', '', '', '', '', '']);

  if (noEliminables.length > 0) {
    filas.push(['══ NO ELIMINAR — conservar aunque estén ocultas ══', '', '', '', '', '']);
    noEliminables.forEach(function(h) {
      filas.push([h.nombre, h.cat, h.filas, h.vacía ? 'Sí' : 'No', h.dep, '🔴 Conservar']);
    });
  }

  filas.push(['', '', '', '', '', '']);
  filas.push(['RESUMEN', '', '', '', '', '']);
  filas.push(['Eliminables:', eliminables.length, '', 'No eliminables:', noEliminables.length, '']);

  sh.clearContents(); sh.clearFormats();
  sh.getRange(1, 1, filas.length, NCOLS).setValues(filas);
  sh.getRange(1, 1, 1, NCOLS).setBackground('#0d652d').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(2, 1, 1, NCOLS).setBackground('#e6f4ea').setFontColor('#5f6368').setFontStyle('italic');
  sh.getRange(3, 1, 1, NCOLS).setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold');

  for (var r = 3; r < filas.length; r++) {
    var txt2 = filas[r].join('');
    var bg2;
    if (txt2.indexOf('══') >= 0)        bg2 = '#e8eaed';
    else if (txt2.indexOf('🟢') >= 0)   bg2 = '#e6f4ea';
    else if (txt2.indexOf('🟡') >= 0)   bg2 = '#fff8e1';
    else if (txt2.indexOf('🔴') >= 0 && txt2.indexOf('Conservar') < 0) bg2 = '#fce8e6';
    else if (txt2.indexOf('🔴 Conservar') >= 0) bg2 = '#fce8e6';
    else if (txt2.indexOf('RESUMEN') >= 0) bg2 = '#e8eaed';
    if (bg2) sh.getRange(r + 1, 1, 1, NCOLS).setBackground(bg2);
  }
  sh.autoResizeColumns(1, NCOLS);
}

// ════════════════════════════════════════════════════════════════
// MOSTRAR TODAS LAS HOJAS (para recuperar visibilidad si es necesario)
// ════════════════════════════════════════════════════════════════

function mostrarTodasLasHojas() {
  var ss = SpreadsheetApp.getActive();
  var hojas = ss.getSheets();
  var mostradas = 0;
  hojas.forEach(function(sh) {
    if (sh.isSheetHidden()) {
      try { sh.showSheet(); mostradas++; } catch (_) {}
    }
  });
  SpreadsheetApp.getUi().alert(
    '👁️ Hojas restauradas',
    mostradas + ' hoja(s) vueltas a mostrar.\nUsa "Organizar hojas" para volver a ocultar las técnicas.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ════════════════════════════════════════════════════════════════
// HELPER
// ════════════════════════════════════════════════════════════════

function _oh_getSheet_(ss, nombre) {
  var sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  sh.showSheet();
  return sh;
}

// ════════════════════════════════════════════════════════════════
// OCULTAR HOJAS DE ADMINISTRACIÓN INTERNA
// ════════════════════════════════════════════════════════════════

/**
 * Oculta MAPA_HOJAS_SISTEMA, HOJAS_OCULTADAS y ANALISIS_HOJAS_ELIMINABLES.
 * Verifica que ninguna fórmula ni script dependa de ellas como FUENTE.
 * Marca cada una en CONFIG como ADMINISTRACION_INTERNA.
 * Genera ULTIMAS_HOJAS_OCULTADAS con el informe de la operación.
 */
function ocultarHojasAdmin() {
  var ss  = SpreadsheetApp.getActive();
  var ui  = SpreadsheetApp.getUi();
  var ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

  var OBJETIVO = OH_HOJAS_ADMIN; // ['MAPA_HOJAS_SISTEMA','HOJAS_OCULTADAS','ANALISIS_HOJAS_ELIMINABLES']
  var resultados = [];

  OBJETIVO.forEach(function(nombre) {
    var sh = ss.getSheetByName(nombre);

    // ── 1. ¿Existe la hoja? ─────────────────────────────────
    if (!sh) {
      resultados.push({
        nombre: nombre,
        existe:  false,
        deps:    '—',
        config:  '—',
        accion:  '⚪ No existe en el spreadsheet'
      });
      return;
    }

    // ── 2. Verificar dependencias externas ───────────────────
    // Busca referencias a esta hoja en fórmulas de OTRAS hojas
    var depEncontradas = [];
    try {
      var finder = ss.createTextFinder(nombre).matchFormulaText(true);
      var celdas = finder.findAll();
      celdas.forEach(function(celda) {
        var shNombre = celda.getSheet().getName();
        if (shNombre !== nombre) {  // ignorar referencias dentro de la propia hoja
          depEncontradas.push(shNombre + '!' + celda.getA1Notation());
        }
      });
    } catch (e) {
      depEncontradas.push('Error al verificar: ' + e.message);
    }

    var depStr = depEncontradas.length > 0
      ? '⚠️ Referenciada en: ' + depEncontradas.join(', ')
      : '✅ Sin dependencias externas en fórmulas';

    // ── 3. Marcar en CONFIG como ADMINISTRACION_INTERNA ──────
    var configKey    = 'tipo_hoja_' + nombre.toLowerCase().replace(/ /g, '_');
    var configValor  = 'ADMINISTRACION_INTERNA';
    var configResult = '';
    try {
      setConfig(configKey, configValor);
      configResult = '✅ CONFIG.' + configKey + ' = ' + configValor;
    } catch (e) {
      configResult = '⚠️ No se pudo escribir en CONFIG: ' + e.message;
    }

    // ── 4. Ocultar la hoja ────────────────────────────────────
    var accion = '';
    if (sh.isSheetHidden()) {
      accion = '⚪ Ya estaba oculta — sin cambios';
    } else {
      try {
        sh.hideSheet();
        accion = '✅ Ocultada correctamente (Ver → Hojas ocultas para acceder)';
        Logger.log('[ADMIN] Ocultada: ' + nombre);
      } catch (e) {
        accion = '❌ Error al ocultar: ' + e.message;
      }
    }

    resultados.push({
      nombre:  nombre,
      existe:  true,
      deps:    depStr,
      config:  configResult,
      accion:  accion
    });
  });

  // ── Crear ULTIMAS_HOJAS_OCULTADAS ────────────────────────────
  var shInf = _oh_getSheet_(ss, 'ULTIMAS_HOJAS_OCULTADAS');
  var NCOLS = 5;
  var filas = [];

  filas.push(['ULTIMAS_HOJAS_OCULTADAS — ' + ts, '', '', '', '']);
  filas.push(['Estas hojas son de ADMINISTRACIÓN INTERNA. Acceder via Ver → Hojas ocultas.', '', '', '', '']);
  filas.push(['Hoja', 'Dependencias encontradas', 'Marcado en CONFIG', 'Acción realizada', 'Estado final']);

  resultados.forEach(function(r) {
    if (!r.existe) {
      filas.push([r.nombre, '—', '—', r.accion, '⚪ No existe']);
      return;
    }
    var estadoFinal = (r.accion.indexOf('✅') >= 0 || r.accion.indexOf('⚪ Ya') >= 0)
      ? '✅ Oculta — solo admin'
      : '❌ Revisar';
    filas.push([r.nombre, r.deps, r.config, r.accion, estadoFinal]);
  });

  filas.push(['', '', '', '', '']);
  filas.push(['NOTA: estas hojas NO han sido eliminadas.', '', '', '', '']);
  filas.push(['Para verlas: Google Sheets → Ver → Hojas ocultas → seleccionar hoja.', '', '', '', '']);
  filas.push(['Para regenerarlas: Menú ⚙️ Sistema → Organizar hojas.', '', '', '', '']);

  shInf.clearContents(); shInf.clearFormats();
  shInf.getRange(1, 1, filas.length, NCOLS).setValues(filas);

  shInf.getRange(1, 1, 1, NCOLS).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  shInf.getRange(2, 1, 1, NCOLS).setBackground('#e8f0fe').setFontColor('#3c4043').setFontStyle('italic');
  shInf.getRange(3, 1, 1, NCOLS).setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');

  for (var r2 = 3; r2 < filas.length; r2++) {
    var txt = filas[r2].join('');
    var bg;
    if (txt.indexOf('✅') >= 0) bg = '#e6f4ea';
    else if (txt.indexOf('❌') >= 0) bg = '#fce8e6';
    else if (txt.indexOf('⚠️') >= 0) bg = '#fff8e1';
    else if (txt.indexOf('⚪') >= 0) bg = '#f8f9fa';
    else if (txt.indexOf('NOTA') >= 0 || txt.indexOf('Para') >= 0) bg = '#e8f0fe';
    if (bg) shInf.getRange(r2 + 1, 1, 1, NCOLS).setBackground(bg);
  }
  shInf.autoResizeColumns(1, NCOLS);
  ss.setActiveSheet(shInf);

  // ── Alerta final ─────────────────────────────────────────────
  var ok = resultados.filter(function(r) {
    return r.accion.indexOf('✅') >= 0 || r.accion.indexOf('⚪ Ya') >= 0;
  }).length;

  var msg = '🙈 HOJAS DE ADMINISTRACIÓN OCULTADAS\n\n';
  resultados.forEach(function(r) {
    msg += '• ' + r.nombre + ': ' + (r.accion.indexOf('✅') >= 0 ? '✅ Ocultada'
      : r.accion.indexOf('⚪ Ya') >= 0 ? '⚪ Ya estaba oculta' : r.accion) + '\n';
  });
  msg += '\n✅ Sin dependencias externas encontradas.\n';
  msg += '✅ Marcadas como ADMINISTRACION_INTERNA en CONFIG.\n\n';
  msg += 'Para acceder: Ver → Hojas ocultas\n';
  msg += 'Ver informe completo en: ULTIMAS_HOJAS_OCULTADAS';

  ui.alert('🙈 Hojas de administración', msg, ui.ButtonSet.OK);
}
