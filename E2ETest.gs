// ============================================================
// E2ETest.gs — Prueba End-to-End completa (SIN tocar datos reales)
// ============================================================
//
// Orden de ejecución recomendado:
//   1. validateEndToEndFlow()   — verificar que el sistema está listo
//   2. generateTestData()       — crear datos de prueba
//   3. generateClientReportPDF()— generar PDF (ya en Informes.gs)
//   4. cleanupTestData()        — eliminar datos de prueba
// ============================================================

// ─────────────────────────────────────────────────────────────
// validateEndToEndFlow
// ─────────────────────────────────────────────────────────────

/**
 * Verifica que todos los módulos del sistema están operativos.
 * Comprueba: CLIENTES, OPERARIOS, RESP_TRABAJOS, TRABAJOS,
 * carpeta de fotos y módulo PDF.
 * SOLO LECTURA — no modifica datos.
 */
function validateEndToEndFlow() {
  Logger.log('================================================');
  Logger.log('validateEndToEndFlow()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var allOk = true;
  var ss    = getTargetSS_();

  // ── 1. CLIENTES ───────────────────────────────────────────────
  Logger.log('\n[1] Hoja CLIENTES:');
  try {
    var shCli  = getSheet('CLIENTES');
    var numCli = Math.max(0, shCli.getLastRow() - 1);
    Logger.log('  ✓ GID=' + shCli.getSheetId() + '  registros=' + numCli);
    if (numCli === 0) Logger.log('  ⚠️  Sin registros — añade al menos 1 cliente activo.');
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    allOk = false;
  }

  // ── 2. OPERARIOS ──────────────────────────────────────────────
  Logger.log('\n[2] Hoja OPERARIOS:');
  try {
    var shOpe  = getSheet('OPERARIOS');
    var numOpe = Math.max(0, shOpe.getLastRow() - 1);
    Logger.log('  ✓ GID=' + shOpe.getSheetId() + '  registros=' + numOpe);
    if (numOpe === 0) Logger.log('  ⚠️  Sin registros — añade al menos 1 operario activo.');
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    allOk = false;
  }

  // ── 3. RESP_TRABAJOS ──────────────────────────────────────────
  Logger.log('\n[3] Hoja RESP_TRABAJOS:');
  var respSheet = ss.getSheetByName('RESP_TRABAJOS');
  if (respSheet) {
    Logger.log('  ✓ GID=' + respSheet.getSheetId() +
               '  filas=' + Math.max(0, respSheet.getLastRow() - 1));
  } else {
    Logger.log('  ❌ NO EXISTE.');
    Logger.log('     Ejecuta installTrabajosModule() para crearla.');
    allOk = false;
  }

  var formTrabId = getConfig('form_trabajos_id');
  if (formTrabId) {
    try {
      var formTrab = FormApp.openById(formTrabId);
      var destId   = formTrab.getDestinationId();
      Logger.log('  Formulario: "' + formTrab.getTitle() + '"');
      Logger.log('  Destino:    ' + (destId === ss.getId()
        ? '✓ este spreadsheet'
        : '⚠️  ' + (destId || 'sin destino')));
    } catch (e) {
      Logger.log('  ⚠️  form_trabajos_id configurado pero no accesible: ' + e.message);
    }
  } else {
    Logger.log('  ⚠️  form_trabajos_id no configurado en CONFIG.');
  }

  // ── 4. TRABAJOS (columnas clave del formulario) ───────────────
  Logger.log('\n[4] Hoja TRABAJOS:');
  try {
    var shTra  = getSheet('TRABAJOS');
    var numTra = Math.max(0, shTra.getLastRow() - 1);
    Logger.log('  ✓ GID=' + shTra.getSheetId() + '  registros=' + numTra);

    var hdrs      = shTra.getRange(1, 1, 1, shTra.getLastColumn()).getValues()[0];
    var requeridos = ['ID', 'Cliente', 'Operario', 'Estado', 'Fotos'];
    Logger.log('  Columnas requeridas:');
    for (var f = 0; f < requeridos.length; f++) {
      var campo     = requeridos[f];
      var presente  = false;
      for (var h = 0; h < hdrs.length; h++) {
        if (String(hdrs[h]).toLowerCase().indexOf(campo.toLowerCase()) !== -1) {
          presente = true; break;
        }
      }
      Logger.log('    ' + campo + ': ' + (presente ? '✓' : '❌ NO ENCONTRADO'));
      if (!presente) allOk = false;
    }
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    allOk = false;
  }

  // ── 5. Carpeta de fotos ───────────────────────────────────────
  Logger.log('\n[5] Carpeta de fotos (carpeta_trabajos_id):');
  var carpetaTrabId = getConfig('carpeta_trabajos_id');
  if (!carpetaTrabId) {
    Logger.log('  ⚠️  No configurada. Las fotos se almacenan como texto (enlaces Drive).');
    Logger.log('     Configura carpeta_trabajos_id en CONFIG si quieres una carpeta dedicada.');
  } else {
    try {
      var carpetaTrab = DriveApp.getFolderById(carpetaTrabId);
      Logger.log('  ✓ "' + carpetaTrab.getName() + '"  ID=' + carpetaTrabId);
    } catch (e) {
      Logger.log('  ❌ ID configurado pero no accesible: ' + e.message);
      allOk = false;
    }
  }

  // ── 6. Módulo PDF ─────────────────────────────────────────────
  Logger.log('\n[6] Módulo PDF (DocumentApp + Drive):');
  try {
    var testDoc = DocumentApp.create('_e2e_validate_pdf_test_');
    testDoc.saveAndClose();
    DriveApp.getFileById(testDoc.getId()).setTrashed(true);
    Logger.log('  ✓ DocumentApp y DriveApp accesibles.');
  } catch (e) {
    Logger.log('  ❌ DocumentApp no accesible: ' + e.message);
    allOk = false;
  }

  var carpetaFacId = getConfig('carpeta_facturas_id');
  if (carpetaFacId) {
    try {
      var carpetaFac = DriveApp.getFolderById(carpetaFacId);
      Logger.log('  ✓ Carpeta PDF: "' + carpetaFac.getName() + '"  ID=' + carpetaFacId);
    } catch (e) {
      Logger.log('  ⚠️  carpeta_facturas_id configurado pero no accesible.');
      Logger.log('     El PDF se guardará en la raíz de Drive.');
    }
  } else {
    Logger.log('  ⚠️  carpeta_facturas_id no configurado. PDFs en raíz de Drive.');
    Logger.log('     Configura carpeta_facturas_id en CONFIG para organizar los PDFs.');
  }

  // ── Resumen ───────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('RESULTADO E2E:');
  Logger.log('  ' + (allOk
    ? '✓ SISTEMA LISTO PARA PRUEBA E2E'
    : '❌ HAY PROBLEMAS — revisa los ❌ anteriores'));
  if (allOk) Logger.log('  Siguiente paso: ejecuta generateTestData()');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// generateTestData
// ─────────────────────────────────────────────────────────────

/**
 * Crea 1 cliente de prueba, 1 operario de prueba y 1 trabajo de prueba.
 * NO modifica datos reales. Todo lleva el texto "TEST_E2E" para identificarlo.
 * Configura automáticamente CONFIG para que generateClientReportPDF()
 * genere el informe de prueba sin ajustes manuales.
 *
 * Para eliminar los datos de prueba: ejecuta cleanupTestData()
 */
function generateTestData() {
  Logger.log('================================================');
  Logger.log('generateTestData()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('SIN TOCAR DATOS REALES — identificador: TEST_E2E');
  Logger.log('================================================');

  var hoy    = new Date();
  var hoyStr = Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  // ── 1. Cliente de prueba ──────────────────────────────────────
  Logger.log('\n[1] Creando cliente de prueba...');
  var rCli = createCliente({
    nombre:    'TEST_E2E',
    apellidos: 'Cliente Prueba',
    empresa:   'Empresa TEST E2E',
    email:     'test_e2e_cliente@prueba.local',
    telefono:  '000000001',
    direccion: 'Calle Prueba E2E, 1',
    cifNif:    'TEST00001'
  });

  if (!rCli.success) {
    Logger.log('  ❌ Error al crear cliente: ' + rCli.error);
    return;
  }
  Logger.log('  ✓ ' + rCli.id);

  // Etiqueta del cliente tal como la genera _getTrabajosDropdownClientes_():
  // "Nombre Apellidos (Empresa)"
  var cliLabel = 'TEST_E2E Cliente Prueba (Empresa TEST E2E)';
  Logger.log('  Etiqueta dropdown: "' + cliLabel + '"');

  // ── 2. Operario de prueba ─────────────────────────────────────
  Logger.log('\n[2] Creando operario de prueba...');
  var rOpe = createOperario({
    nombre:    'TEST_E2E',
    apellidos: 'Operario Prueba',
    email:     'test_e2e_operario@prueba.local',
    telefono:  '000000002'
  });

  if (!rOpe.success) {
    Logger.log('  ❌ Error al crear operario: ' + rOpe.error);
    Logger.log('  Revirtiendo cliente de prueba ya creado...');
    _deleteTestRows_('CLIENTES', 'TEST_E2E');
    return;
  }
  Logger.log('  ✓ ' + rOpe.id);

  // Etiqueta del operario tal como la genera _getTrabajosDropdownOperarios_():
  // "Nombre Apellidos"
  var opeLabel = 'TEST_E2E Operario Prueba';
  Logger.log('  Etiqueta dropdown: "' + opeLabel + '"');

  // ── 3. Trabajo de prueba ──────────────────────────────────────
  Logger.log('\n[3] Creando trabajo de prueba...');
  var rTra = createTrabajo({
    cliente:          cliLabel,
    operario:         opeLabel,
    fechaTrabajo:     hoy,     // Date object — evita que new Date("DD/MM/YYYY") lo interprete como MM/DD
    horaInicio:       '08:00',
    horaFin:          '10:00',
    trabajoRealizado: 'TEST_E2E — Trabajo de prueba. No es un dato real. ' +
                      'Eliminar con cleanupTestData().',
    incidencias:      'Sin incidencias (TEST_E2E)',
    estado:           'Completado',
    fotos:            'TEST_E2E_enlace_ficticio_no_real',
    ubicacion:        'Ubicación de prueba E2E'
  });

  if (!rTra.success) {
    Logger.log('  ❌ Error al crear trabajo: ' + rTra.error);
    return;
  }
  Logger.log('  ✓ ' + rTra.id);

  // ── 4. Configurar CONFIG para generateClientReportPDF() ───────
  Logger.log('\n[4] Configurando CONFIG para el informe de prueba...');
  setConfig('cliente_reporte',      cliLabel);
  setConfig('fecha_inicio_reporte', hoyStr);
  setConfig('fecha_fin_reporte',    hoyStr);
  Logger.log('  ✓ cliente_reporte:      "' + cliLabel + '"');
  Logger.log('  ✓ fecha_inicio_reporte: ' + hoyStr);
  Logger.log('  ✓ fecha_fin_reporte:    ' + hoyStr);

  // ── Resumen ───────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('DATOS DE PRUEBA CREADOS:');
  Logger.log('  Cliente:  ' + rCli.id + ' — "' + cliLabel + '"');
  Logger.log('  Operario: ' + rOpe.id + ' — "' + opeLabel + '"');
  Logger.log('  Trabajo:  ' + rTra.id + ' — fecha ' + hoyStr);
  Logger.log('');
  Logger.log('SIGUIENTE PASO: ejecuta generateClientReportPDF()');
  Logger.log('  Lee CONFIG y genera el PDF con el trabajo de prueba.');
  Logger.log('  (función en Informes.gs — ya configurada)');
  Logger.log('');
  Logger.log('LIMPIEZA (cuando termines la prueba):');
  Logger.log('  Ejecuta cleanupTestData()');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// generateTestInvoice
// ─────────────────────────────────────────────────────────────

/**
 * Ejecuta el flujo completo de factura de prueba de forma automática:
 *   1. Crea datos de prueba (cliente, operario, trabajo) via generateTestData()
 *   2. Configura las claves de factura en CONFIG con valores de prueba
 *   3. Llama a generateInvoicePDF() para generar el PDF
 *   4. Muestra un resumen con la URL del PDF generado
 *
 * Para eliminar los datos de prueba: ejecuta cleanupTestData()
 */
function generateTestInvoice() {
  Logger.log('================================================');
  Logger.log('generateTestInvoice()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('SIN TOCAR DATOS REALES — identificador: TEST_E2E');
  Logger.log('================================================');

  // ── 1. Crear datos de prueba ──────────────────────────────────
  Logger.log('\n[PASO 1] Creando datos de prueba...');
  generateTestData();

  // ── 2. Configurar claves de factura ──────────────────────────
  Logger.log('\n[PASO 2] Configurando claves de factura en CONFIG...');
  var hoy    = new Date();
  var hoyStr = Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var cliLabel = 'TEST_E2E Cliente Prueba (Empresa TEST E2E)';

  setConfig('cliente_factura',      cliLabel);
  setConfig('fecha_inicio_factura', hoyStr);
  setConfig('fecha_fin_factura',    hoyStr);
  setConfig('precio_hora',          '35');
  setConfig('materiales',           '0');

  Logger.log('  ✓ cliente_factura:       "' + cliLabel + '"');
  Logger.log('  ✓ fecha_inicio_factura:  ' + hoyStr);
  Logger.log('  ✓ fecha_fin_factura:     ' + hoyStr);
  Logger.log('  ✓ precio_hora:           35');
  Logger.log('  ✓ materiales:            0');

  // ── 3. Generar factura PDF ────────────────────────────────────
  Logger.log('\n[PASO 3] Generando factura PDF...');
  generateInvoicePDF();

  // ── Resumen ───────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('FACTURA DE PRUEBA COMPLETADA.');
  Logger.log('  Busca el PDF en Drive (o en la carpeta carpeta_facturas_id si está configurada).');
  Logger.log('  Busca la fila en la hoja FACTURAS con el número de factura más reciente.');
  Logger.log('');
  Logger.log('LIMPIEZA (cuando termines la prueba):');
  Logger.log('  Ejecuta cleanupTestData()');
  Logger.log('  (También limpia las claves de factura de CONFIG)');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// cleanupTestData
// ─────────────────────────────────────────────────────────────

/**
 * Elimina todos los datos de prueba E2E de CLIENTES, OPERARIOS y TRABAJOS.
 * Identifica las filas por el texto "TEST_E2E" en cualquier columna.
 * Resetea las claves de informe en CONFIG.
 * NO toca datos reales.
 */
function cleanupTestData() {
  Logger.log('================================================');
  Logger.log('cleanupTestData()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var total = 0;

  Logger.log('\n[1] Limpiando CLIENTES...');
  total += _deleteTestRows_('CLIENTES', 'TEST_E2E');

  Logger.log('\n[2] Limpiando OPERARIOS...');
  total += _deleteTestRows_('OPERARIOS', 'TEST_E2E');

  Logger.log('\n[3] Limpiando TRABAJOS...');
  total += _deleteTestRows_('TRABAJOS', 'TEST_E2E');

  Logger.log('\n[4] Reseteando CONFIG de informe...');
  setConfig('cliente_reporte',      '');
  setConfig('fecha_inicio_reporte', '');
  setConfig('fecha_fin_reporte',    '');
  Logger.log('  ✓ cliente_reporte, fecha_inicio_reporte, fecha_fin_reporte → (vacías)');

  Logger.log('\n[5] Reseteando CONFIG de factura...');
  setConfig('cliente_factura',      '');
  setConfig('fecha_inicio_factura', '');
  setConfig('fecha_fin_factura',    '');
  setConfig('precio_hora',          '');
  setConfig('materiales',           '');
  Logger.log('  ✓ cliente_factura, fecha_inicio_factura, fecha_fin_factura, precio_hora, materiales → (vacías)');

  Logger.log('\n================================================');
  Logger.log('LIMPIEZA COMPLETADA. Total filas eliminadas: ' + total);
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// Helpers privados
// ─────────────────────────────────────────────────────────────

/**
 * Elimina filas que contengan el tag en las primeras 12 columnas.
 * Recorre de abajo hacia arriba para no desplazar índices.
 * Devuelve el número de filas eliminadas.
 */
function _deleteTestRows_(sheetName, tag) {
  try {
    var sheet   = getSheet(sheetName);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('  ' + sheetName + ': sin filas de datos.');
      return 0;
    }
    var numCols = Math.min(sheet.getLastColumn(), 12);
    var data    = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var deleted = 0;
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i].join('|').indexOf(tag) !== -1) {
        sheet.deleteRow(i + 2);
        deleted++;
      }
    }
    Logger.log('  ' + sheetName + ': eliminadas ' + deleted + ' fila(s) de prueba.');
    return deleted;
  } catch (e) {
    Logger.log('  ❌ Error en ' + sheetName + ': ' + e.message);
    return 0;
  }
}
