# Sistema de Gestión para Autónomos — Google Workspace

Plantilla completa de gestión empresarial para autónomos, construida íntegramente sobre **Google Workspace** (Apps Script, Sheets, Drive, Gmail, Forms). Sin servidores externos. Sin dependencias de terceros. 100% en la nube de Google.

---

## Arquitectura general

```
Google Workspace
│
├── Google Sheets          ← Base de datos + paneles visuales
├── Google Apps Script     ← Lógica de negocio (este repositorio)
├── Google Drive           ← Almacenamiento de PDFs, facturas, fotos
├── Google Forms           ← Alta de clientes, operarios y gastos
└── Gmail                  ← Envío de facturas, recordatorios, informes
```

Todo el sistema se ejecuta dentro del ecosistema de Google. No requiere VPS, servidores propios ni APIs externas.

---

## Estructura del Spreadsheet

| Hoja | Descripción |
|------|-------------|
| `CONFIG` | Configuración general del sistema (nombre empresa, CIF, email gestor, etc.) |
| `CLIENTES` | Registro de clientes con ID único |
| `OPERARIOS` | Registro de operarios/empleados |
| `TRABAJOS` | Trabajos realizados con estado, cliente y operario |
| `FACTURAS` | Facturas emitidas con estado de cobro e IVA |
| `GASTOS` | Gastos con proveedor, IVA soportado y estado de pago |
| `PENDIENTES` | Trabajos pendientes de cobro/facturación |
| `COMPLETADOS` | Histórico de trabajos completados |
| `LOGS` | Log de eventos del sistema |
| `RESP_CLIENTES` | Respuestas formulario de clientes (enlazado a Google Forms) |
| `RESP_OPERARIOS` | Respuestas formulario de operarios |
| `RESP_TRABAJOS` | Respuestas formulario de partes de trabajo |
| `RESP_GASTOS` | Respuestas formulario de registro de gastos |
| `DRIVE_CONFIG` | IDs de carpetas Drive estructuradas por año y categoría |

---

## Módulos (archivos Apps Script)

### Núcleo del sistema

| Archivo | Función |
|---------|---------|
| `Code.gs` | Setup inicial, gestión de hojas, envío al gestor, informes de gastos, `doGet()` |
| `Config.gs` | Lectura y escritura de la hoja CONFIG. Función `getConfig()` / `setConfig()` |
| `Utils.gs` | Utilidades compartidas: `getTargetSS_()`, `getSheet()`, `logEvent()`, helpers de fecha |
| `Panel.gs` | Menú principal `onOpen()`, router de eventos `onEdit()`, búsqueda de trabajos en paneles |
| `Triggers.gs` | Instalación de triggers, creación de formularios Google Forms |
| `Install.gs` | Referencia histórica de instalación inicial (LEGACY — no activo) |

### Entidades de negocio

| Archivo | Función |
|---------|---------|
| `Clientes.gs` | Alta, edición y búsqueda de clientes |
| `Operarios.gs` | Alta y gestión de operarios |
| `Trabajos.gs` | Registro de partes de trabajo, asignación de fotos Drive, gestión de estado |
| `Facturas.gs` | Emisión de facturas, cálculo de IVA repercutido, PDF |
| `FacturasPresupuesto.gs` | Conversión presupuesto → factura, proformas |
| `Presupuestos.gs` | Gestión de presupuestos |
| `Gastos.gs` | Registro de gastos, subida de tickets, OCR automático via Drive |
| `Cobros.gs` | Control de cobros y estado de facturas |

### Paneles visuales (dentro del Spreadsheet)

| Archivo | Panel | Descripción |
|---------|-------|-------------|
| `PanelGeneral.gs` | `PANEL_GENERAL` | KPIs globales: facturación, cobros, gastos, trabajos pendientes |
| `PanelCliente.gs` | `PANEL_CLIENTE` | Vista por cliente: trabajos, facturas, cobros y pendientes de facturar |
| `PanelGastos.gs` | `PANEL_GASTOS` | Control de gastos con IVA soportado, estado de pago y exportación al gestor |
| `PanelContable.gs` | `PANEL_CONTABLE` | Panel contable por período: IVA repercutido vs soportado, resultado fiscal |
| `PanelContableEmail.gs` | — | Envío del informe contable por email al gestor |
| `PanelContableExport.gs` | — | Exportación a CSV/Drive del cierre contable |
| `PanelContableHistorial.gs` | — | Historial de cierres contables anteriores |
| `PanelContableSelector.gs` | — | Selector de período para PANEL_CONTABLE |
| `PanelContableTest.gs` | — | Test de integración del módulo contable |
| `PanelPendientes.gs` | `PANEL_PENDIENTES` | Vista global de clientes con trabajos sin facturar |

### Automatizaciones

| Archivo | Función |
|---------|---------|
| `Recordatorios.gs` | Envío automático diario de recordatorios de pago a clientes con facturas vencidas |
| `InformeMensual.gs` | Generación de informe mensual en Sheets |
| `Informes.gs` | Generación de informes PDF de cliente |
| `IVAFiscal.gs` | Cálculo de liquidaciones de IVA por trimestre |
| `DocTrabajos.gs` | Generación de documentos Google Docs por trabajo |
| `Alertas.gs` | Sistema de alertas internas |
| `Dashboard.gs` | Integración con Looker Studio (instrucciones) |
| `DashboardFinanciero.gs` | Dashboard financiero en Sheets |

### Drive

| Archivo | Función |
|---------|---------|
| `DriveManager.gs` | Creación y gestión de la estructura de carpetas Drive (`ALVARO_EMPRESA/YYYY/...`) |
| `DriveMigrate185.gs` | Migración de archivos entre carpetas Drive (M1–M4) |
| `DriveScan185.gs` | Diagnóstico de archivos Drive: audita IDs, detecta huérfanos, valida estructura |

### Herramientas de desarrollo

| Archivo | Función |
|---------|---------|
| `Diagnostics.gs` | Diagnóstico de hojas, formularios y fotos. Solo lectura |
| `E2ETest.gs` | Tests end-to-end del flujo completo. No toca datos reales |
| `Cleanup.gs` | Utilidades de limpieza de datos de prueba |

### HTML (sidebars y diálogos)

| Archivo | Función |
|---------|---------|
| `EditarGasto.html` | Formulario de edición de un gasto desde el panel |
| `EditarPago.html` | Formulario de edición de estado de cobro de una factura |

---

## Flujo de datos principal

```
Google Forms (cliente/operario/trabajo/gasto)
        ↓  onFormSubmit trigger
Google Sheets (CLIENTES / OPERARIOS / TRABAJOS / GASTOS)
        ↓  lógica de negocio en Apps Script
Google Drive (PDFs, fotos, tickets)
        ↓  envío por email
Gmail (facturas, recordatorios, informes al gestor)
```

---

## Triggers instalados

| Trigger | Función | Cuándo |
|---------|---------|--------|
| `onFormSubmitHandler` | Router de formularios | Al enviar cualquier formulario vinculado |
| `onOpen` | Crea menú | Al abrir el Spreadsheet |
| `onPanelEdit` | Gestiona checkboxes y botones de paneles | Al editar el Spreadsheet |
| `enviarRecordatoriosPendientes` | Recordatorios de cobro | Diariamente a las 9:00 |

---

## Permisos OAuth requeridos

```
https://www.googleapis.com/auth/spreadsheets
https://www.googleapis.com/auth/drive
https://www.googleapis.com/auth/documents
https://www.googleapis.com/auth/gmail.send
https://www.googleapis.com/auth/forms
https://www.googleapis.com/auth/script.container.ui
https://www.googleapis.com/auth/script.scriptapp
https://www.googleapis.com/auth/userinfo.email
```

---

## Configuración inicial (para nuevas instalaciones)

1. Crear un nuevo Google Spreadsheet.
2. Abrir **Extensiones → Apps Script**.
3. Copiar todos los archivos `.gs` y `.html` en el editor.
4. Copiar `appsscript.json` (activar vista de archivos de manifiesto primero).
5. En `Install.gs`, sustituir los placeholders:
   - `SPREADSHEET_ID_HERE` → ID del Spreadsheet creado
   - `FORM_CLIENTES_ID_HERE` → ID del formulario de clientes (creado automáticamente por `setupForms()`)
   - `FORM_OPERARIOS_ID_HERE` → ID del formulario de operarios
6. Ejecutar `setup()` desde el editor de Apps Script.
7. Ejecutar `setupTriggers()`.
8. Ejecutar `setupForms()`.
9. Rellenar la hoja `CONFIG` con los datos de la empresa.

---

## Estructura de carpetas Drive generada

```
ALVARO_EMPRESA/
├── SISTEMA/
│   ├── CONFIG/
│   ├── LOGS/
│   ├── EXPORTACIONES/
│   ├── TEMP/
│   └── PLANTILLAS/
├── CLIENTES/
│   └── {nombre_cliente}/
│       ├── FACTURAS/
│       ├── PRESUPUESTOS/
│       ├── PROFORMAS/
│       ├── FOTOS/
│       ├── INFORMES/
│       └── HISTORIAL/
└── {YYYY}/
    ├── FACTURAS/
    │   └── {01_Enero .. 12_Diciembre}/
    ├── GASTOS/
    │   └── {01_Enero .. 12_Diciembre}/
    ├── FOTOS_TRABAJOS/
    ├── PRESUPUESTOS/
    ├── PROFORMAS/
    ├── INFORMES/
    ├── IVA/
    ├── CONTABILIDAD/
    └── BACKUP/
```

---

## Seguridad

- **Sin tokens externos**: el sistema no llama a ningún servidor fuera de Google.
- **Sin API keys de terceros**: usa exclusivamente Google OAuth.
- **Autenticación de configuración**: el `TARGET_SS_ID` se almacena en `PropertiesService` (no en el código).
- **Sin VPS ni infraestructura propia**: 100% Google Workspace.

---

## Licencia

Plantilla de uso libre para autónomos. Adaptar a las necesidades de cada negocio antes de usar en producción.

---

*Generado el 2026-06-05 — Sistema de gestión para autónomos sobre Google Workspace.*
