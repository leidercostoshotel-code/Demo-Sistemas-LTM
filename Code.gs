// ============================================================
// GOOGLE APPS SCRIPT - Backend para Gestor de Consumos Mensuales
// ============================================================
// INSTRUCCIONES DE INSTALACION:
// 1. Ve a https://script.google.com y crea un nuevo proyecto
// 2. Copia TODO este codigo en el archivo Code.gs
// 3. Haz clic en "Implementar" > "Nueva implementacion"
// 4. Selecciona tipo: "Aplicacion web"
// 5. Ejecutar como: "Yo" (tu cuenta)
// 6. Acceso: "Cualquier persona"
// 7. Copia la URL generada y pegala en el index.html (variable APPS_SCRIPT_URL)
// 8. Cada vez que modifiques el script, crea una NUEVA implementacion
// ============================================================

const SPREADSHEET_ID = ''; // <-- PEGA AQUI el ID de tu Google Sheet (opcional, si esta vacio usa el vinculado)

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  // Si no hay ID, crear o usar uno nuevo
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SHEET_ID');
  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch(e) {
      // Si fue borrado, crear uno nuevo
    }
  }
  const ss = SpreadsheetApp.create('Consumos Mensuales - Historial');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getOrCreateSheet(mesKey) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(mesKey);
  if (!sheet) {
    sheet = ss.insertSheet(mesKey);
    // Crear cabeceras
    const headers = [
      'abreviatura', 'pdv', 'nroChk', 'importe',
      'comestibles', 'teCafe', 'vinos', 'gaseosas',
      'cerveza', 'licores', 'fecha', 'invitadoMotivo',
      'firmante', 'observaciones', 'guardadoEn'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a1a2e')
      .setFontColor('#d4a853');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================================
// ENDPOINT GET - Cargar datos de un mes o listar meses
// ============================================================
function doGet(e) {
  try {
    const action = e.parameter.action || 'cargar';

    if (action === 'listarMeses') {
      return listarMeses();
    }

    if (action === 'cargar') {
      const mes = e.parameter.mes;
      if (!mes) {
        return jsonResponse({ success: false, error: 'Parametro "mes" requerido (ej: 2024-01)' });
      }
      return cargarMes(mes);
    }

    return jsonResponse({ success: false, error: 'Accion no reconocida' });
  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

// ============================================================
// ENDPOINT POST - Guardar datos de un mes
// ============================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'guardar';

    if (action === 'guardar') {
      const mes = payload.mes;
      const datos = payload.datos;
      if (!mes || !datos) {
        return jsonResponse({ success: false, error: 'Se requiere "mes" y "datos"' });
      }
      return guardarMes(mes, datos);
    }

    if (action === 'eliminarMes') {
      const mes = payload.mes;
      if (!mes) {
        return jsonResponse({ success: false, error: 'Se requiere "mes"' });
      }
      return eliminarMes(mes);
    }

    return jsonResponse({ success: false, error: 'Accion no reconocida' });
  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

// ============================================================
// FUNCIONES DE DATOS
// ============================================================

function listarMeses() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const meses = [];

  sheets.forEach(sheet => {
    const nombre = sheet.getName();
    // Solo incluir hojas con formato de mes (YYYY-MM)
    if (/^\d{4}-\d{2}$/.test(nombre)) {
      const lastRow = sheet.getLastRow();
      const filas = lastRow > 1 ? lastRow - 1 : 0; // Restar cabecera
      meses.push({
        mes: nombre,
        filas: filas,
        bloqueado: false // Puedes agregar logica de bloqueo aqui
      });
    }
  });

  // Ordenar por mes descendente (mas reciente primero)
  meses.sort((a, b) => b.mes.localeCompare(a.mes));

  return jsonResponse({ success: true, meses: meses });
}

function cargarMes(mesKey) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(mesKey);

  if (!sheet) {
    return jsonResponse({ success: true, datos: [], mesNuevo: true });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return jsonResponse({ success: true, datos: [], mesNuevo: false });
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  const datos = data.map(row => ({
    abreviatura: row[0] || '',
    pdv: row[1] || '',
    nroChk: row[2] || '',
    importe: row[3] || '',
    comestibles: row[4] || '',
    teCafe: row[5] || '',
    vinos: row[6] || '',
    gaseosas: row[7] || '',
    cerveza: row[8] || '',
    licores: row[9] || '',
    fecha: row[10] ? Utilities.formatDate(new Date(row[10]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    invitadoMotivo: row[11] || '',
    firmante: row[12] || '',
    observaciones: row[13] || ''
  }));

  return jsonResponse({ success: true, datos: datos, mesNuevo: false });
}

function guardarMes(mesKey, datos) {
  const sheet = getOrCreateSheet(mesKey);

  // Limpiar datos anteriores (mantener cabecera)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 15).clearContent();
  }

  if (datos.length === 0) {
    return jsonResponse({ success: true, mensaje: 'Mes limpiado correctamente', filas: 0 });
  }

  const ahora = new Date();
  const values = datos.map(d => [
    d.abreviatura || '',
    d.pdv || '',
    d.nroChk || '',
    d.importe || '',
    d.comestibles || '',
    d.teCafe || '',
    d.vinos || '',
    d.gaseosas || '',
    d.cerveza || '',
    d.licores || '',
    d.fecha || '',
    d.invitadoMotivo || '',
    d.firmante || '',
    d.observaciones || '',
    ahora
  ]);

  sheet.getRange(2, 1, values.length, 15).setValues(values);

  // Formato
  sheet.autoResizeColumns(1, 15);

  return jsonResponse({
    success: true,
    mensaje: 'Guardado exitosamente',
    filas: datos.length,
    guardadoEn: ahora.toISOString()
  });
}

function eliminarMes(mesKey) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(mesKey);

  if (!sheet) {
    return jsonResponse({ success: false, error: 'Mes no encontrado' });
  }

  // No eliminar si es la unica hoja
  if (ss.getSheets().length <= 1) {
    // Solo limpiar contenido
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 15).clearContent();
    }
  } else {
    ss.deleteSheet(sheet);
  }

  return jsonResponse({ success: true, mensaje: 'Mes eliminado' });
}

// ============================================================
// UTILIDADES
// ============================================================
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Test manual (ejecutar desde el editor para verificar)
function testListarMeses() {
  const result = listarMeses();
  Logger.log(result.getContent());
}

function testGuardar() {
  const datos = [
    { abreviatura: 'ATEN. VyM', pdv: 'La Locanda', nroChk: '001', importe: 150, comestibles: 80, teCafe: 20, vinos: 30, gaseosas: 10, cerveza: 0, licores: 10, fecha: '2024-01-15', invitadoMotivo: 'Test', firmante: 'Admin', observaciones: '' }
  ];
  const sheet = getOrCreateSheet('2024-01');
  Logger.log('Sheet creado: ' + sheet.getName());
}
