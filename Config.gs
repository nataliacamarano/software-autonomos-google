// ============================================================
// Config.gs — Constantes del sistema y gestión de configuración
// ============================================================

const SHEET_NAMES = {
  CONFIG:      'CONFIG',
  CLIENTES:    'CLIENTES',
  OPERARIOS:   'OPERARIOS',
  TRABAJOS:    'TRABAJOS',
  FACTURAS:    'FACTURAS',
  PROFORMAS:   'PROFORMAS',
  INFORMES:    'INFORMES',
  PENDIENTES:  'PENDIENTES',
  COMPLETADOS: 'COMPLETADOS',
  LOGS:        'LOGS'
};

const SHEET_HEADERS = {
  CONFIG: [
    'Clave', 'Valor'
  ],
  CLIENTES: [
    'ID_Cliente', 'Timestamp', 'Nombre', 'Apellidos', 'Empresa',
    'Email', 'Telefono', 'Direccion', 'CIF_NIF', 'Activo',
    'Metodos_Pago'
  ],
  OPERARIOS: [
    'ID_Operario', 'Timestamp', 'Nombre', 'Apellidos',
    'Email', 'Telefono', 'Especialidad', 'Activo'
  ],
  TRABAJOS: [
    'ID_Trabajo', 'Fecha_Creacion', 'ID_Cliente', 'Nombre_Cliente',
    'ID_Operario', 'Nombre_Operario', 'Descripcion', 'Tipo_Trabajo',
    'Direccion_Trabajo', 'Fecha_Inicio', 'Fecha_Fin_Estimada',
    'Estado', 'Horas_Trabajadas', 'Coste_Materiales',
    'Precio_Mano_Obra', 'Precio_Total', 'Observaciones',
    'Facturado', 'ID_Factura'
  ],
  FACTURAS: [
    'ID_Factura', 'Fecha_Emision', 'Serie', 'Numero',
    'ID_Cliente', 'Nombre_Cliente', 'Email_Cliente',
    'Direccion_Cliente', 'CIF_Cliente', 'IDs_Trabajos', 'Concepto',
    'Base_Imponible', 'IVA_Porcentaje', 'IVA_Importe',
    'IRPF_Porcentaje', 'IRPF_Importe', 'Total',
    'Estado', 'URL_PDF', 'Fecha_Envio', 'Fecha_Pago',
    'Metodo_Pago_Usado', 'Datos_Pago_Usados'
  ],
  PENDIENTES: [
    'ID_Trabajo', 'Fecha_Creacion', 'Nombre_Cliente', 'Nombre_Operario',
    'Descripcion', 'Tipo_Trabajo', 'Fecha_Inicio', 'Fecha_Fin_Estimada',
    'Estado', 'Precio_Total'
  ],
  COMPLETADOS: [
    'ID_Trabajo', 'Fecha_Creacion', 'Nombre_Cliente', 'Nombre_Operario',
    'Descripcion', 'Tipo_Trabajo', 'Fecha_Inicio', 'Fecha_Fin_Estimada',
    'Estado', 'Precio_Total', 'Facturado', 'ID_Factura'
  ],
  PROFORMAS: [
    'ID_Proforma', 'Fecha_Emision', 'Serie', 'Numero',
    'ID_Cliente', 'Nombre_Cliente', 'Email_Cliente',
    'Direccion_Cliente', 'CIF_Cliente', 'IDs_Trabajos', 'Concepto',
    'Base_Imponible', 'Total', 'Estado', 'URL_PDF', 'Fecha_Envio', 'Fecha_Pago',
    'ID_Presupuesto_Origen', 'Estado_Cobro',
    'Metodo_Pago_Usado', 'Datos_Pago_Usados'
  ],
  INFORMES: [
    'ID_Informe', 'Fecha_Emision', 'ID_Cliente', 'Nombre_Cliente',
    'Periodo_Desde', 'Periodo_Hasta', 'Trabajos_Contados', 'Horas_Totales',
    'URL_PDF', 'Estado', 'IDs_Trabajos'
  ],
  LOGS: [
    'Timestamp', 'Nivel', 'Funcion', 'Mensaje', 'Detalle'
  ]
};

// Pares [clave, valor] escritos en CONFIG solo si la clave no existe aún
const DEFAULT_CONFIG_ROWS = [
  ['empresa_nombre',      'Mi Empresa'],
  ['empresa_cif',         'A00000000'],
  ['empresa_direccion',   'Calle Ejemplo, 1 - 28001 Madrid'],
  ['empresa_email',       ''],
  ['empresa_telefono',    '600 000 000'],
  ['iva_porcentaje',      '21'],
  ['irpf_porcentaje',     '15'],
  ['serie_factura',       'FAC'],
  ['siguiente_numero',    '1'],
  ['carpeta_facturas_id', ''],
  ['carpeta_trabajos_id', '']
];

// ---- Acceso a configuración ----

/**
 * Lee el valor de una clave de la hoja CONFIG.
 * Devuelve null si no existe.
 */
function getConfig(key) {
  const sheet = getSheet(SHEET_NAMES.CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return String(data[i][1]);
  }
  return null;
}

/**
 * Actualiza el valor de una clave en CONFIG.
 * Si la clave no existe la añade al final.
 */
function setConfig(key, value) {
  const sheet = getSheet(SHEET_NAMES.CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      const cell = sheet.getRange(i + 1, 2);
      cell.setNumberFormat('@');   // evita que Sheets convierta fechas automáticamente
      cell.setValue(String(value));
      return;
    }
  }
  sheet.appendRow([key, String(value)]);
  sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('@');
}

/**
 * Devuelve toda la configuración como objeto { clave: valor }.
 */
function getAllConfig() {
  try {
    const sheet = getSheet(SHEET_NAMES.CONFIG);
    const data = sheet.getDataRange().getValues();
    const config = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) config[String(data[i][0])] = data[i][1];
    }
    return config;
  } catch (_) {
    return {};
  }
}

/**
 * Inicializa la hoja CONFIG: crea la cabecera si está vacía
 * y escribe los valores por defecto para las claves que falten.
 */
function initConfig() {
  const sheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SHEET_HEADERS.CONFIG);
    applyHeaderStyle_(sheet, SHEET_HEADERS.CONFIG.length);
  }
  const existing = getAllConfig();
  DEFAULT_CONFIG_ROWS.forEach(([key, value]) => {
    if (!(key in existing)) sheet.appendRow([key, value]);
  });
}
