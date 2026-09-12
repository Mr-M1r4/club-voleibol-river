// ===== NUCLEO APLICACION =====
const CONFIG = {
  SHEET_NAME_AFILIADOS: 'Afiliados',
  SHEET_NAME_PAGOS: 'Pagos',
  SHEET_NAME_ASISTENCIA: 'Asistencia',
  SHEET_NAME_CONFIG: 'Config',
  COL_AFILIADO: {
    ID: 0, NOMBRES: 1, APELLIDOS: 2, TIPO_DOC: 3, NUM_DOC: 4,
    FECHA_NAC: 5, EDAD: 6, TELEFONO: 7, EMAIL: 8, DIRECCION: 9,
    ACUDIENTE: 10, TEL_ACUDIENTE: 11, CATEGORIA: 12,
    FECHA_INGRESO: 13, FECHA_VENCIMIENTO: 14, ESTADO: 15,
    QR_URL: 16, NOTAS: 17
  },
  COL_PAGO: {
    ID: 0, ID_AFILIADO: 1, NOMBRE: 2, MONTO: 3, FECHA_PAGO: 4,
    MES_PAGADO: 5, METODO_PAGO: 6, COMPROBANTE: 7, NOTAS: 8
  },
  COL_ASISTENCIA: {
    ID: 0, ID_AFILIADO: 1, NOMBRE: 2, FECHA: 3, HORA: 4,
    TIPO: 5, QR_ESCANEADO: 6
  },
  ESTADOS: {
    ACTIVO: 'Activo',
    VENCIDO: 'Vencido',
    SUSPENDIDO: 'Suspendido',
    RETIRADO: 'Retirado'
  },
  TIPOS_ASISTENCIA: ['Entrenamiento', 'Partido', 'Evento', 'Otro'],
  MENSAJES: {
    BIENVENIDA: '¡Bienvenido al Club Deportivo de Voleibol River! Ya estás registrado en nuestro sistema.',
    RENOVACION: 'Tu membresía en el Club Voleibol River ha sido renovada exitosamente. Vigencia: ',
    VENCIMIENTO: 'Tu membresía está próxima a vencer. Realiza tu pago para continuar activo.',
    ASISTENCIA: 'Asistencia registrada correctamente en el Club Voleibol River.',
    RECORDATORIO_PAGO: 'Recordatorio: Tu cuota del Club Voleibol River está pendiente de pago.'
  }
};

function onOpen() {
  crearMenu();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.source.getActiveSheet();
  const name = sheet.getName();

  if (name === CONFIG.SHEET_NAME_PAGOS && e.range.getColumn() === CONFIG.COL_PAGO.FECHA_PAGO + 1) {
    procesarPago(e.range.getRow());
  }
}

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;

  if (page === 'asistencia') return renderAsistenciaQR(e);
  if (page === 'qr') return generarQRDeportista(e);
  if (page === 'dashboard') return renderDashboard();
  if (page === 'encuesta') {
    const template = HtmlService.createTemplateFromFile('Encuesta');
    return template.evaluate()
      .setTitle('Encuesta de clases de vóleibol - Club Voleibol River')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Club Voleibol River - Sistema de Gestión')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  if (!e) {
    return jsonOut({ status: 'error', message: 'Solicitud inválida' });
  }

  const action = (e.parameter && e.parameter.action) || '';

  switch (action) {
    case 'registrar_asistencia':
      return registrarAsistenciaWeb(e);
    case 'webhook_pago':
      return webhookPago(e);
    case 'encuesta':
      return registrarRespuestaEncuesta(e);
    case 'preinscripcion':
      return responderPreinscripcion(e);
    default:
      if (e.postData && e.postData.contents) {
        return webhookPago(e);
      }
      return jsonOut({ status: 'error', message: 'Acción no reconocida' });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function crearMenu() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    return;
  }
  if (!ui) return;

  try {
    ui.createMenu('🏐 Club Voleibol River')
      .addItem('Inicializar Sistema', 'inicializarSistema')
      .addSeparator()
      .addItem('Generar QR para todos los afiliados', 'generarQRTodos')
      .addItem('Actualizar estados de afiliados', 'actualizarEstadoAfiliados')
      .addSeparator()
      .addItem('Configurar automatizaciones diarias', 'configurarAutomatizaciones')
      .addItem('Ver automatizaciones activas', 'verAutomatizaciones')
      .addSeparator()
      .addItem('Crear Formulario de Registro', 'crearFormularioRegistro')
      .addItem('Crear Formulario de Asistencia', 'crearFormularioAsistencia')
      .addSeparator()
      .addItem('Probar Notificaciones', 'probarNotificacion')
      .addItem('Ver Dashboard', 'abrirDashboard')
      .addToUi();
  } catch (e) {
    Logger.log('Menú no disponible en modo standalone: ' + e);
  }
}

function abrirDashboard() {
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { return; }
  if (!ui) return;
  const url = ScriptApp.getService().getUrl() + '?page=dashboard';
  const html = '<html><body><script>window.open("' + url + '");google.script.host.close();</script></body></html>';
  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setHeight(10).setWidth(10),
    'Abriendo Dashboard...'
  );
}
// ===== BASE =====
const SPREADSHEET_ID_KEY = 'CLUB_VOLEIBOL_RIVER_SS_ID';

function getSpreadsheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_KEY);
  if (id) {
    ss = SpreadsheetApp.openById(id);
    if (ss) return ss;
  }

  throw new Error(
    'No hay una hoja de cálculo vinculada.\n\n' +
    'Opción 1 (recomendada):\n' +
    '1. Crea un Google Sheet en drive.google.com\n' +
    '2. Extensiones → Apps Script\n' +
    '3. Pega allí todo el código\n\n' +
    'Opción 2:\n' +
    'Ejecuta primero: inicializarSistema()'
  );
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function inicializarSistema() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    ss = SpreadsheetApp.create('Club Voleibol River - Sistema de Gestión');
    PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_KEY, ss.getId());
  }

  crearEstructuraAfiliados(ss);
  crearEstructuraPagos(ss);
  crearEstructuraAsistencia(ss);
  crearEstructuraConfig(ss);
  crearEstructuraEncuestas(ss);
  crearEstructuraInscripciones(ss);

  crearMenu();

  const url = ss.getUrl();
  Logger.log('Sheet creado: ' + url);

  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  if (ui) {
    ui.alert(
      '✅ Sistema inicializado correctamente.\n\n' +
      'Sheet: ' + ss.getName() + '\n' +
      'URL: ' + url + '\n\n' +
      'Pestañas creadas:\n' +
      '- Afiliados (registro de deportistas)\n' +
      '- Pagos (control de pagos y renovación automática)\n' +
      '- Asistencia (registro vía QR)\n' +
      '- Config (credenciales y parámetros)\n' +
      '- Encuestas (respuestas del estudio de mercado)\n' +
      '- Inscripciones (pre-inscripciones desde la web)\n\n' +
      '👉 Abre el sheet y recarga para ver el menú "🏐 Club Voleibol River"'
    );
  }

  return url;
}

function crearEstructuraAfiliados(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_AFILIADOS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME_AFILIADOS);
  sheet.clear();

  const headers = [
    'ID Afiliado', 'Nombres', 'Apellidos', 'Tipo Documento', 'Número Documento',
    'Fecha de Nacimiento', 'Edad', 'Teléfono', 'Email', 'Dirección',
    'Nombre Acudiente', 'Teléfono Acudiente', 'Categoría',
    'Fecha de Ingreso', 'Fecha de Vencimiento', 'Estado',
    'URL Código QR', 'Notas'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#2E75B6');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF');

  sheet.setFrozenRows(1);

  const widths = [10, 20, 20, 15, 18, 15, 8, 15, 25, 25, 20, 18, 15, 15, 15, 12, 40, 30];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.ESTADOS.ACTIVO, CONFIG.ESTADOS.VENCIDO, CONFIG.ESTADOS.SUSPENDIDO, CONFIG.ESTADOS.RETIRADO], true)
    .build();
  sheet.getRange('P:P').setDataValidation(validation);
}

function crearEstructuraPagos(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_PAGOS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME_PAGOS);
  sheet.clear();

  const headers = [
    'ID Pago', 'ID Afiliado', 'Nombre Afiliado', 'Monto', 'Fecha de Pago',
    'Mes Pagado', 'Método de Pago', 'Comprobante URL', 'Notas'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#2E75B6');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);

  const widths = [10, 12, 25, 12, 15, 15, 15, 30, 20];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const methodValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Efectivo', 'Transferencia', 'Nequi', 'Daviplata', 'Bancolombia', 'PayPal', 'Otro'], true)
    .build();
  sheet.getRange('G:G').setDataValidation(methodValidation);
}

function crearEstructuraAsistencia(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_ASISTENCIA);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME_ASISTENCIA);
  sheet.clear();

  const headers = [
    'ID Registro', 'ID Afiliado', 'Nombre Afiliado', 'Fecha', 'Hora',
    'Tipo', 'Escaneado por QR'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#2E75B6');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);

  const widths = [12, 12, 25, 15, 10, 16, 12];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const tipoValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.TIPOS_ASISTENCIA, true)
    .build();
  sheet.getRange('F:F').setDataValidation(tipoValidation);
}

function crearEstructuraConfig(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME_CONFIG);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME_CONFIG);
  sheet.clear();

  const configData = [
    ['Parámetro', 'Valor', 'Descripción'],
    ['TELEGRAM_BOT_TOKEN', '', 'Token del bot de Telegram (de @BotFather)'],
    ['TELEGRAM_CHAT_ID', '', 'Chat ID para notificaciones grupales'],
    ['WHATSAPP_API_URL', '', 'URL de API de WhatsApp Business/tercero'],
    ['WHATSAPP_API_KEY', '', 'API Key de WhatsApp'],
    ['WHATSAPP_GROUP_ID', '', 'ID del grupo de WhatsApp para notificaciones grupales'],
    ['TWILIO_ACCOUNT_SID', '', 'Account SID de Twilio (para SMS)'],
    ['TWILIO_AUTH_TOKEN', '', 'Auth Token de Twilio'],
    ['TWILIO_PHONE_NUMBER', '', 'Número Twilio (formato +57...)'],
    ['WEBHOOK_TOKEN', '', 'Token secreto para proteger el webhook de pagos'],
    ['NOTIFICACION_HABILITADA', 'SI', 'SI/NO - Activar notificaciones automáticas'],
    ['CANAL_PRINCIPAL', 'TELEGRAM', 'Canal principal: TELEGRAM, WHATSAPP, SMS'],
    ['CUOTA_MENSUAL', '120000', 'Valor cuota mensual en COP'],
    ['META_MENSAJE', 'Club Voleibol River', 'Firma para mensajes'],
  ];

  sheet.getRange(1, 1, 1, 3).setValues([configData[0]]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  sheet.getRange(1, 1, 1, 3).setBackground('#2E75B6');
  sheet.getRange(1, 1, 1, 3).setFontColor('#FFFFFF');

  const rows = configData.slice(1).map(r => [r[0], r[1], r[2]]);
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.getRange(2, 1, rows.length, 1).setFontWeight('bold');

  sheet.setColumnWidth(1, 35);
  sheet.setColumnWidth(2, 35);
  sheet.setColumnWidth(3, 50);
  sheet.setFrozenRows(1);
  sheet.protect().setWarningOnly(true);
}

function crearEstructuraEncuestas(ss) {
  let sheet = ss.getSheetByName('Encuestas');
  if (!sheet) sheet = ss.insertSheet('Encuestas');
  sheet.clear();

  const headers = [
    'Fecha', 'Plan', 'Nombre', 'Teléfono', 'Email', 'Zona',
    'Precio Justo', 'Por qué no', 'Nivel', 'Horario', 'Categoría Edad',
    'Motivación', 'Comentarios'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#2E75B6');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

function crearEstructuraInscripciones(ss) {
  let sheet = ss.getSheetByName('Inscripciones');
  if (!sheet) sheet = ss.insertSheet('Inscripciones');
  sheet.clear();

  const headers = [
    'Fecha', 'Nombre', 'Celular', 'Categoría', 'Origen'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#2E75B6');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

function obtenerConfig(key) {
  const sheet = getSheet(CONFIG.SHEET_NAME_CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return '';
}

function generarId() {
  const ts = Date.now().toString(36).toUpperCase();
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let suf = '';
  for (let i = 0; i < 2; i++) {
    suf += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return ts + suf;
}

function parseFechaLocal(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'string') {
    let m = value.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = value.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  }
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function esFechaVencida(value) {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const fecha = parseFechaLocal(value);
  if (!fecha) return false;
  fecha.setHours(0, 0, 0, 0);
  return fecha < hoy;
}

function diasParaVencer(value) {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const fecha = parseFechaLocal(value);
  if (!fecha) return null;
  fecha.setHours(0, 0, 0, 0);
  return Math.round((fecha - hoy) / (1000 * 60 * 60 * 24));
}

function escaparHTML(value) {
  if (value === null || value === undefined) return '';
  return HtmlService.createHtmlOutput(String(value)).getContent();
}

function configurarAutomatizaciones() {
  eliminarAutomatizaciones();
  ScriptApp.newTrigger('actualizarEstadoAfiliados')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(0)
    .inTimezone('America/Bogota')
    .create();

  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  const mensaje = '✅ Automatización diaria configurada: todos los días a las 06:00 (America/Bogota) se verificarán los vencimientos de membresía.';
  if (ui) {
    ui.alert(mensaje);
  } else {
    Logger.log(mensaje);
  }
  return { success: true, mensaje };
}

function eliminarAutomatizaciones() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'actualizarEstadoAfiliados') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function verAutomatizaciones() {
  const triggers = ScriptApp.getProjectTriggers().filter(t =>
    t.getHandlerFunction() === 'actualizarEstadoAfiliados'
  );

  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}

  const mensaje = triggers.length > 0
    ? '✅ Hay ' + triggers.length + ' automatización(es) diaria(s) activa(s).'
    : '⚠️ No hay automatizaciones activas. Usa "Configurar automatizaciones diarias".';

  if (ui) {
    ui.alert(mensaje);
  } else {
    Logger.log(mensaje);
  }
  return { success: true, activas: triggers.length };
}
// ===== AFILIADOS =====
function registrarAfiliado(data) {
  data = data || {};
  const nombres = String(data.nombres || '').trim();
  const apellidos = String(data.apellidos || '').trim();

  if (!nombres || !apellidos) {
    return { success: false, mensaje: 'Nombres y apellidos son obligatorios.' };
  }
  if (!data.numDoc) {
    return { success: false, mensaje: 'El número de documento es obligatorio.' };
  }

  const existente = getAfiliadoPorDoc(data.tipoDoc || 'CC', data.numDoc);
  if (existente) {
    return {
      success: false,
      mensaje: 'Ya existe un afiliado con ese documento: ' +
        existente.data['Nombres'] + ' ' + existente.data['Apellidos']
    };
  }

  const sheet = getSheet(CONFIG.SHEET_NAME_AFILIADOS);
  const id = generarId();
  const now = new Date();

  const vencimiento = new Date(now);
  vencimiento.setMonth(vencimiento.getMonth() + 1);

  let edad = '';
  if (data.fechaNac) {
    const nac = parseFechaLocal(data.fechaNac);
    if (nac) {
      edad = Math.max(0, Math.floor((now - nac) / (365.25 * 24 * 60 * 60 * 1000)));
    }
  }

  const fila = [
    id,
    nombres.toUpperCase(),
    apellidos.toUpperCase(),
    data.tipoDoc || 'CC',
    String(data.numDoc),
    data.fechaNac || '',
    edad,
    data.telefono || '',
    data.email || '',
    data.direccion || '',
    data.acudiente || '',
    data.telAcudiente || '',
    data.categoria || 'General',
    Utilities.formatDate(now, 'America/Bogota', 'yyyy-MM-dd'),
    Utilities.formatDate(vencimiento, 'America/Bogota', 'yyyy-MM-dd'),
    CONFIG.ESTADOS.ACTIVO,
    generarURLPaginaQR(id),
    data.notas || ''
  ];

  sheet.appendRow(fila);

  if (obtenerConfig('NOTIFICACION_HABILITADA') === 'SI') {
    const mensaje = CONFIG.MENSAJES.BIENVENIDA +
      '\nNombre: ' + nombres + ' ' + apellidos +
      '\nID: ' + id +
      '\nVigencia hasta: ' + Utilities.formatDate(vencimiento, 'America/Bogota', 'yyyy-MM-dd');
    enviarNotificacion(data.telefono, mensaje);
  }

  return { success: true, id: id, mensaje: 'Afiliado registrado exitosamente' };
}

function getAfiliados(filtro) {
  const sheet = getSheet(CONFIG.SHEET_NAME_AFILIADOS);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  const headers = data[0];
  const rows = data.slice(1);

  let result = rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });
    return obj;
  });

  if (filtro) {
    if (filtro.estado) {
      result = result.filter(r => r['Estado'] === filtro.estado);
    }
    if (filtro.search) {
      const s = filtro.search.toLowerCase();
      result = result.filter(r =>
        (r['Nombres'] || '').toLowerCase().includes(s) ||
        (r['Apellidos'] || '').toLowerCase().includes(s) ||
        (r['Número Documento'] || '').includes(s)
      );
    }
  }

  return result;
}

function getAfiliadoPorId(id) {
  const sheet = getSheet(CONFIG.SHEET_NAME_AFILIADOS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][CONFIG.COL_AFILIADO.ID] === id) {
      const obj = {};
      data[0].forEach((h, j) => { obj[h] = data[i][j]; });
      return { row: i + 1, data: obj };
    }
  }
  return null;
}

function getAfiliadoPorDoc(tipo, numero) {
  const sheet = getSheet(CONFIG.SHEET_NAME_AFILIADOS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][CONFIG.COL_AFILIADO.TIPO_DOC] === (tipo || 'CC') &&
        String(data[i][CONFIG.COL_AFILIADO.NUM_DOC]) === String(numero)) {
      const obj = {};
      data[0].forEach((h, j) => { obj[h] = data[i][j]; });
      return { row: i + 1, data: obj };
    }
  }
  return null;
}

function actualizarAfiliado(id, campos) {
  const sheet = getSheet(CONFIG.SHEET_NAME_AFILIADOS);
  const result = getAfiliadoPorId(id);
  if (!result) return { success: false, mensaje: 'Afiliado no encontrado' };

  const headers = sheet.getDataRange().getValues()[0];
  Object.keys(campos).forEach(key => {
    const colIndex = headers.indexOf(key);
    if (colIndex >= 0) {
      sheet.getRange(result.row, colIndex + 1).setValue(campos[key]);
    }
  });

  return { success: true, mensaje: 'Afiliado actualizado correctamente' };
}

function actualizarEstadoAfiliados() {
  const sheet = getSheet(CONFIG.SHEET_NAME_AFILIADOS);
  const data = sheet.getDataRange().getValues();
  let actualizados = 0;

  for (let i = 1; i < data.length; i++) {
    const id = data[i][CONFIG.COL_AFILIADO.ID];
    if (!id) continue;
    const estado = data[i][CONFIG.COL_AFILIADO.ESTADO];
    if (estado === CONFIG.ESTADOS.RETIRADO) continue;

    const vencimientoStr = data[i][CONFIG.COL_AFILIADO.FECHA_VENCIMIENTO];
    if (!vencimientoStr) continue;

    if (esFechaVencida(vencimientoStr) && estado === CONFIG.ESTADOS.ACTIVO) {
      sheet.getRange(i + 1, CONFIG.COL_AFILIADO.ESTADO + 1).setValue(CONFIG.ESTADOS.VENCIDO);
      actualizados++;

      const telefono = data[i][CONFIG.COL_AFILIADO.TELEFONO];
      if (telefono && obtenerConfig('NOTIFICACION_HABILITADA') === 'SI') {
        enviarNotificacion(telefono, CONFIG.MENSAJES.VENCIMIENTO);
      }
    }

    const dias = diasParaVencer(vencimientoStr);
    if (dias === 5 && estado === CONFIG.ESTADOS.ACTIVO) {
      const telefono = data[i][CONFIG.COL_AFILIADO.TELEFONO];
      if (telefono && obtenerConfig('NOTIFICACION_HABILITADA') === 'SI') {
        enviarNotificacion(telefono, CONFIG.MENSAJES.RECORDATORIO_PAGO);
      }
    }
  }

  return { success: true, actualizados };
}

function generarContenidoQR(idAfiliado) {
  return ScriptApp.getService().getUrl() + '?page=asistencia&id=' + idAfiliado;
}

function generarURLPaginaQR(idAfiliado) {
  return ScriptApp.getService().getUrl() + '?page=qr&id=' + idAfiliado;
}
// ===== PAGOS =====
function registrarPago(data) {
  data = data || {};
  const sheet = getSheet(CONFIG.SHEET_NAME_PAGOS);
  const id = generarId();
  const ahora = new Date();

  const afiliado = getAfiliadoPorDoc(data.tipoDoc, data.numDoc);
  if (!afiliado) {
    return { success: false, mensaje: 'Afiliado no encontrado. Verifica tipo y número de documento.' };
  }

  const monto = parseFloat(data.monto) > 0 ? parseFloat(data.monto) : parseFloat(obtenerConfig('CUOTA_MENSUAL'));
  const fila = [
    id,
    afiliado.data['ID Afiliado'],
    afiliado.data['Nombres'] + ' ' + afiliado.data['Apellidos'],
    monto,
    Utilities.formatDate(ahora, 'America/Bogota', 'yyyy-MM-dd'),
    data.mesPagado || Utilities.formatDate(ahora, 'America/Bogota', 'yyyy-MM'),
    data.metodoPago || 'Efectivo',
    data.comprobanteUrl || '',
    data.notas || ''
  ];

  sheet.appendRow(fila);
  renovarMembresia(afiliado);

  return {
    success: true,
    id: id,
    mensaje: 'Pago registrado y membresía renovada automáticamente'
  };
}

function procesarPago(fila) {
  const sheet = getSheet(CONFIG.SHEET_NAME_PAGOS);
  const rowData = sheet.getRange(fila, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idAfiliado = rowData[CONFIG.COL_PAGO.ID_AFILIADO];

  if (!idAfiliado) return;

  const afiliado = getAfiliadoPorId(idAfiliado);
  if (afiliado) {
    renovarMembresia(afiliado);
  }
}

function renovarMembresia(afiliado) {
  const sheet = getSheet(CONFIG.SHEET_NAME_AFILIADOS);
  const hoy = new Date();
  const actual = parseFechaLocal(afiliado.data['Fecha de Vencimiento']);

  let base = hoy;
  if (actual && actual > hoy) base = actual;

  const nueva = new Date(base);
  nueva.setMonth(nueva.getMonth() + 1);

  const fechaStr = Utilities.formatDate(nueva, 'America/Bogota', 'yyyy-MM-dd');
  sheet.getRange(afiliado.row, CONFIG.COL_AFILIADO.FECHA_VENCIMIENTO + 1).setValue(fechaStr);
  sheet.getRange(afiliado.row, CONFIG.COL_AFILIADO.ESTADO + 1).setValue(CONFIG.ESTADOS.ACTIVO);

  const telefono = afiliado.data['Teléfono'];
  if (telefono && obtenerConfig('NOTIFICACION_HABILITADA') === 'SI') {
    const mensaje = CONFIG.MENSAJES.RENOVACION +
      Utilities.formatDate(nueva, 'America/Bogota', 'dd/MM/yyyy');
    enviarNotificacion(telefono, mensaje);
  }
}

function getHistorialPagos(idAfiliado) {
  const sheet = getSheet(CONFIG.SHEET_NAME_PAGOS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const rows = data.slice(1);

  return rows
    .filter(row => row[CONFIG.COL_PAGO.ID_AFILIADO] === idAfiliado)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    .sort((a, b) => {
      if (a['Fecha de Pago'] < b['Fecha de Pago']) return 1;
      if (a['Fecha de Pago'] > b['Fecha de Pago']) return -1;
      return 0;
    });
}

function getResumenPagos() {
  const sheet = getSheet(CONFIG.SHEET_NAME_PAGOS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { total: 0, mes: 0, pendientes: 0 };
  }

  const rows = data.slice(1);
  const ahora = new Date();
  const mesActual = Utilities.formatDate(ahora, 'America/Bogota', 'yyyy-MM');

  let total = 0;
  let mes = 0;

  rows.forEach(row => {
    const monto = parseFloat(row[CONFIG.COL_PAGO.MONTO]) || 0;
    total += monto;
    const mesPago = String(row[CONFIG.COL_PAGO.MES_PAGADO] || '');
    if (mesPago === mesActual) {
      mes += monto;
    }
  });

  const afiliadosSheet = getSheet(CONFIG.SHEET_NAME_AFILIADOS);
  const afiliadosData = afiliadosSheet.getDataRange().getValues();
  const pendientes = afiliadosData.slice(1).filter(r =>
    r[CONFIG.COL_AFILIADO.ESTADO] === CONFIG.ESTADOS.VENCIDO
  ).length;

  return { total, mes, pendientes };
}

function generarReportePagosPorMes() {
  const sheet = getSheet(CONFIG.SHEET_NAME_PAGOS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const rows = data.slice(1);
  const agrupado = {};

  rows.forEach(row => {
    const mes = String(row[CONFIG.COL_PAGO.MES_PAGADO] || 'Desconocido');
    const monto = parseFloat(row[CONFIG.COL_PAGO.MONTO]) || 0;
    if (!agrupado[mes]) agrupado[mes] = 0;
    agrupado[mes] += monto;
  });

  return Object.keys(agrupado)
    .sort()
    .map(mes => ({ mes, total: agrupado[mes] }));
}
// ===== ASISTENCIA =====
function registrarAsistencia(idAfiliado, tipo) {
  const sheet = getSheet(CONFIG.SHEET_NAME_ASISTENCIA);
  const afiliado = getAfiliadoPorId(idAfiliado);
  if (!afiliado) {
    return { success: false, mensaje: 'Afiliado no encontrado. QR inválido.' };
  }

  if (afiliado.data['Estado'] !== CONFIG.ESTADOS.ACTIVO) {
    return {
      success: false,
      mensaje: 'Membresía ' + afiliado.data['Estado'] + '. Renueva tu membresía para registrar asistencia.',
      estado: afiliado.data['Estado']
    };
  }

  tipo = (CONFIG.TIPOS_ASISTENCIA.indexOf(tipo) >= 0) ? tipo : 'Entrenamiento';
  const ahora = new Date();
  const id = generarId();

  const fila = [
    id,
    idAfiliado,
    afiliado.data['Nombres'] + ' ' + afiliado.data['Apellidos'],
    Utilities.formatDate(ahora, 'America/Bogota', 'yyyy-MM-dd'),
    Utilities.formatDate(ahora, 'America/Bogota', 'HH:mm:ss'),
    tipo,
    'SÍ'
  ];

  sheet.appendRow(fila);

  if (obtenerConfig('NOTIFICACION_HABILITADA') === 'SI') {
    const mensaje = CONFIG.MENSAJES.ASISTENCIA +
      '\nFecha: ' + Utilities.formatDate(ahora, 'America/Bogota', 'dd/MM/yyyy') +
      '\nHora: ' + Utilities.formatDate(ahora, 'America/Bogota', 'HH:mm') +
      '\nTipo: ' + tipo;
    enviarNotificacion(afiliado.data['Teléfono'], mensaje);
  }

  const nombre = afiliado.data['Nombres'] + ' ' + afiliado.data['Apellidos'];
  return {
    success: true,
    mensaje: 'Asistencia registrada para ' + nombre,
    nombre: nombre
  };
}

function registrarAsistenciaWeb(e) {
  const id = e.parameter && e.parameter.id ? e.parameter.id : '';
  const tipo = (e.parameter && e.parameter.tipo) || 'Entrenamiento';
  const result = registrarAsistencia(id, tipo);

  if (e.parameter && e.parameter.format === 'json') {
    return jsonOut(result);
  }

  return renderAsistenciaResult(result);
}

function getAsistenciaAfiliado(idAfiliado, dias) {
  const sheet = getSheet(CONFIG.SHEET_NAME_ASISTENCIA);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const rows = data.slice(1);
  const fechaCorte = new Date();
  fechaCorte.setDate(fechaCorte.getDate() - (dias || 30));

  return rows
    .filter(row => row[CONFIG.COL_ASISTENCIA.ID_AFILIADO] === idAfiliado)
    .filter(row => {
      const fecha = parseFechaLocal(row[CONFIG.COL_ASISTENCIA.FECHA]);
      return fecha && fecha >= fechaCorte;
    })
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    .sort((a, b) => {
      if (a['Fecha'] < b['Fecha']) return 1;
      if (a['Fecha'] > b['Fecha']) return -1;
      return 0;
    });
}

function getResumenAsistencia(dias) {
  const sheet = getSheet(CONFIG.SHEET_NAME_ASISTENCIA);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { total: 0, porTipo: {}, diario: [] };

  const rows = data.slice(1);
  const fechaCorte = new Date();
  fechaCorte.setDate(fechaCorte.getDate() - (dias || 30));

  const filtradas = rows.filter(row => {
    const fecha = parseFechaLocal(row[CONFIG.COL_ASISTENCIA.FECHA]);
    return fecha && fecha >= fechaCorte;
  });

  const porTipo = {};
  const diario = {};
  let total = 0;

  filtradas.forEach(row => {
    total++;
    const tipo = row[CONFIG.COL_ASISTENCIA.TIPO] || 'General';
    porTipo[tipo] = (porTipo[tipo] || 0) + 1;
    const fecha = String(row[CONFIG.COL_ASISTENCIA.FECHA]);
    diario[fecha] = (diario[fecha] || 0) + 1;
  });

  return {
    total,
    porTipo,
    diario: Object.keys(diario).sort().map(f => ({ fecha: f, count: diario[f] }))
  };
}

function renderAsistenciaQR(e) {
  const id = escaparHTML(e && e.parameter && e.parameter.id ? e.parameter.id : '');
  const nombre = escaparHTML(e && e.parameter && e.parameter.nombre ? e.parameter.nombre : '');

  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Registro de Asistencia - Club Voleibol River</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; font-family: Arial, sans-serif; }
    body {
      background: linear-gradient(135deg, #2E75B6 0%, #1a4f7a 100%);
      min-height: 100vh; display: flex; align-items: center; justify-content: center;
      padding: 20px;
    }
    .card {
      background: white; border-radius: 20px; padding: 35px; max-width: 420px; width: 100%;
      box-shadow: 0 20px 60px rgba(0,0,0,0.3); text-align: center;
    }
    .logo { font-size: 28px; font-weight: bold; color: #2E75B6; margin-bottom: 5px; }
    .subtitle { color: #666; font-size: 14px; margin-bottom: 25px; }
    .avatar {
      width: 100px; height: 100px; border-radius: 50%; background: #e8f0fe;
      display: flex; align-items: center; justify-content: center;
      margin: 0 auto 20px; font-size: 40px; color: #2E75B6;
    }
    h2 { color: #333; margin-bottom: 8px; }
    .status { color: #888; margin-bottom: 25px; font-size: 14px; }
    .btn {
      display: block; width: 100%; padding: 16px; border: none; border-radius: 12px;
      font-size: 16px; font-weight: bold; cursor: pointer; margin-bottom: 12px;
      transition: transform 0.15s, box-shadow 0.15s;
    }
    .btn:active { transform: scale(0.97); }
    .btn-primary { background: #2E75B6; color: white; }
    .btn-primary:hover { background: #1a5a8a; }
    .btn-success { background: #28a745; color: white; }
    .btn-success:hover { background: #1e7e34; }
    .btn-outline { background: white; color: #2E75B6; border: 2px solid #2E75B6; }
    .mensaje { margin-top: 20px; padding: 15px; border-radius: 10px; display: none; }
    .mensaje.success { background: #d4edda; color: #155724; display: block; }
    .mensaje.error { background: #f8d7da; color: #721c24; display: block; }
    .mensaje.warning { background: #fff3cd; color: #856404; display: block; }
    .spinner { display: none; width: 40px; height: 40px; border: 4px solid #f3f3f3;
      border-top: 4px solid #2E75B6; border-radius: 50%; animation: spin 1s linear infinite;
      margin: 15px auto; }
    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    .input-group { margin-bottom: 20px; text-align: left; }
    label { display: block; color: #555; font-size: 13px; margin-bottom: 5px; font-weight: bold; }
    select, input {
      width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 10px;
      font-size: 15px; transition: border-color 0.3s;
    }
    select:focus, input:focus { outline: none; border-color: #2E75B6; }
    .footer { margin-top: 20px; font-size: 12px; color: #aaa; }
  </style>
</head>
<body>
  <div class="card">
    <div class="logo">🏐 Club Voleibol River</div>
    <div class="subtitle">Registro de Asistencia</div>
    <div class="avatar">👤</div>
    <h2 id="nombreDisplay">${nombre || 'Escanea para registrarte'}</h2>
    <div class="status" id="statusDisplay">${id ? 'ID: ' + id : 'Pendiente de escaneo'}</div>

    <div class="input-group">
      <label>Tipo de actividad</label>
      <select id="tipoAsistencia">
        <option value="Entrenamiento">Entrenamiento</option>
        <option value="Partido">Partido</option>
        <option value="Evento">Evento</option>
        <option value="Otro">Otro</option>
      </select>
    </div>

    <button class="btn btn-success" onclick="registrar()">✅ Registrar Asistencia</button>
    <div class="spinner" id="spinner"></div>
    <div id="mensaje" class="mensaje"></div>
    <div class="footer">Club Deportivo de Voleibol River — Rionegro, Antioquia</div>
  </div>

  <script>
    const params = new URLSearchParams(window.location.search);
    const id = params.get('id') || '${id}';

    function registrar() {
      const btn = document.querySelector('.btn-success');
      const spinner = document.getElementById('spinner');
      const mensaje = document.getElementById('mensaje');

      btn.style.display = 'none';
      spinner.style.display = 'block';
      mensaje.style.display = 'none';

      const tipo = document.getElementById('tipoAsistencia').value;

      google.script.run
        .withSuccessHandler(function(result) {
          spinner.style.display = 'none';
          btn.style.display = 'block';
          mensaje.className = 'mensaje ' + (result.success ? 'success' : 'error');
          mensaje.textContent = result.mensaje;
          if (result.nombre) {
            document.getElementById('nombreDisplay').textContent = result.nombre;
          }
        })
        .withFailureHandler(function(err) {
          spinner.style.display = 'none';
          btn.style.display = 'block';
          mensaje.className = 'mensaje error';
          mensaje.textContent = 'Error: ' + err;
        })
        .registrarAsistencia(id, tipo);
    }
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle('Registro de Asistencia - Club Voleibol River')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderAsistenciaResult(result) {
  const ok = !!(result && result.success);
  const icon = ok ? '✅' : '❌';
  const color = ok ? '#28a745' : '#dc3545';
  const title = ok ? 'Asistencia Registrada' : 'Error';
  const mensaje = escaparHTML(result && result.mensaje ? result.mensaje : '');

  const html = `
<!DOCTYPE html>
<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Resultado - Club Voleibol River</title>
<style>
  body { font-family: Arial; display: flex; align-items: center; justify-content: center;
    min-height: 100vh; background: linear-gradient(135deg, #2E75B6, #1a4f7a); margin: 0; padding: 20px; }
  .card { background: white; border-radius: 20px; padding: 40px; max-width: 400px; text-align: center; box-shadow: 0 20px 60px rgba(0,0,0,0.3); }
  .icon { font-size: 64px; margin-bottom: 15px; }
  h2 { color: ${color}; margin-bottom: 10px; }
  p { color: #666; margin-bottom: 25px; line-height: 1.5; }
  .btn { display: inline-block; padding: 12px 30px; background: #2E75B6; color: white;
    text-decoration: none; border-radius: 10px; font-weight: bold; }
</style></head>
<body>
  <div class="card">
    <div class="icon">${icon}</div>
    <h2>${title}</h2>
    <p>${mensaje}</p>
    <a class="btn" href="javascript:window.close()">Cerrar</a>
  </div>
</body></html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle('Resultado - Club Voleibol River');
}

function renderDashboard() {
  const totalAfiliados = getAfiliados().length;
  const activos = getAfiliados({ estado: CONFIG.ESTADOS.ACTIVO }).length;
  const resumenPagos = getResumenPagos();
  const resumenAsistencia = getResumenAsistencia(30);
  const now = Utilities.formatDate(new Date(), 'America/Bogota', "dd 'de' MMMM 'de' yyyy");

  const html = `
<!DOCTYPE html>
<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard - Club Voleibol River</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; font-family: Arial, sans-serif; }
  body { background: #f0f2f5; padding: 20px; }
  .header { background: #2E75B6; color: white; padding: 25px; border-radius: 15px; margin-bottom: 25px; }
  .header h1 { font-size: 24px; }
  .header p { opacity: 0.9; font-size: 14px; margin-top: 5px; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 20px; margin-bottom: 25px; }
  .card { background: white; border-radius: 15px; padding: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
  .card .num { font-size: 32px; font-weight: bold; color: #2E75B6; }
  .card .label { color: #666; font-size: 13px; margin-top: 5px; }
  .card .sub { color: #999; font-size: 11px; margin-top: 3px; }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: 15px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
  th { background: #2E75B6; color: white; padding: 12px 15px; text-align: left; font-size: 13px; }
  td { padding: 10px 15px; border-bottom: 1px solid #eee; font-size: 13px; }
  .badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 11px; font-weight: bold; }
  .badge-activo { background: #d4edda; color: #155724; }
  .badge-vencido { background: #f8d7da; color: #721c24; }
  .badge-suspendido { background: #fff3cd; color: #856404; }
</style>
</head><body>
  <div class="header"><h1>🏐 Club Voleibol River</h1><p>Panel de control — ${now}</p></div>
  <div class="grid">
    <div class="card"><div class="num">${totalAfiliados}</div><div class="label">Total Afiliados</div></div>
    <div class="card"><div class="num">${activos}</div><div class="label">Activos</div></div>
    <div class="card"><div class="num">$${resumenPagos.mes.toLocaleString()}</div><div class="label">Recaudado este mes</div></div>
    <div class="card"><div class="num">${resumenPagos.pendientes}</div><div class="label">Membresías vencidas</div></div>
    <div class="card"><div class="num">${resumenAsistencia.total}</div><div class="label">Asistencias (30 días)</div></div>
    <div class="card"><div class="num">$${resumenPagos.total.toLocaleString()}</div><div class="label">Total recaudado histórico</div></div>
  </div>
</body></html>`;

  return HtmlService.createHtmlOutput(html).setTitle('Dashboard - Club Voleibol River');
}
// ===== QR =====
function getQRImageURL(contenido, size) {
  const s = (size || 250);
  return 'https://api.qrserver.com/v1/create-qr-code/?size=' + s + 'x' + s +
    '&data=' + encodeURIComponent(contenido);
}

function generarQRDeportista(e) {
  const id = e && e.parameter && e.parameter.id ? e.parameter.id : '';
  const afiliado = id ? getAfiliadoPorId(id) : null;

  if (!afiliado) {
    return ContentService
      .createTextOutput('Afiliado no encontrado')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const data = afiliado.data;
  const qrContent = generarContenidoQR(id);
  const qrImg = getQRImageURL(qrContent, 250);

  const nombre = escaparHTML(data['Nombres'] + ' ' + data['Apellidos']);
  const tipoDoc = escaparHTML(data['Tipo Documento']);
  const numDoc = escaparHTML(data['Número Documento']);
  const categoria = escaparHTML(data['Categoría'] || 'General');
  const vencimiento = escaparHTML(data['Fecha de Vencimiento']);
  const estado = escaparHTML(data['Estado']);

  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>QR - ${nombre}</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; font-family: Arial, sans-serif; }
    body {
      display: flex; align-items: center; justify-content: center;
      min-height: 100vh; background: #f0f2f5; padding: 20px;
    }
    .card {
      background: white; border-radius: 20px; padding: 35px; max-width: 420px; width: 100%;
      text-align: center; box-shadow: 0 10px 40px rgba(0,0,0,0.15);
    }
    .header { font-size: 20px; font-weight: bold; color: #2E75B6; margin-bottom: 5px; }
    .sub { color: #888; font-size: 13px; margin-bottom: 20px; }
    .qr-container {
      background: white; padding: 20px; border-radius: 15px;
      display: inline-block; margin: 10px auto;
    }
    .qr-container img { width: 250px; height: 250px; image-rendering: pixelated; }
    .info { margin-top: 20px; text-align: left; }
    .info-row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #eee; }
    .info-row .lbl { color: #888; font-size: 13px; }
    .info-row .val { color: #333; font-weight: bold; font-size: 13px; }
    .actions { margin-top: 25px; display: flex; gap: 10px; }
    .btn {
      flex: 1; padding: 12px; border: none; border-radius: 10px; font-size: 14px;
      font-weight: bold; cursor: pointer; transition: opacity 0.2s;
    }
    .btn-download { background: #2E75B6; color: white; }
    .btn-print { background: #6c757d; color: white; }
    .btn:hover { opacity: 0.9; }
    .estado {
      display: inline-block; padding: 4px 15px; border-radius: 20px;
      font-size: 12px; font-weight: bold; margin-top: 15px;
    }
    .estado-Activo { background: #d4edda; color: #155724; }
    .estado-Vencido { background: #f8d7da; color: #721c24; }
  </style>
</head>
<body>
  <div class="card">
    <div class="header">🏐 Club Voleibol River</div>
    <div class="sub">Código QR de Afiliado</div>

    <div class="qr-container">
      <img src="${qrImg}" alt="QR">
    </div>

    <div class="estado estado-${estado}">${estado}</div>

    <div class="info">
      <div class="info-row"><span class="lbl">Nombre</span><span class="val">${nombre}</span></div>
      <div class="info-row"><span class="lbl">Documento</span><span class="val">${tipoDoc} ${numDoc}</span></div>
      <div class="info-row"><span class="lbl">Categoría</span><span class="val">${categoria}</span></div>
      <div class="info-row"><span class="lbl">Vigencia</span><span class="val">${vencimiento}</span></div>
      <div class="info-row"><span class="lbl">ID</span><span class="val">${id}</span></div>
    </div>

    <div class="actions">
      <button class="btn btn-download" onclick="descargar()">📥 Descargar</button>
      <button class="btn btn-print" onclick="window.print()">🖨 Imprimir</button>
    </div>
  </div>

  <script>
    function descargar() {
      const img = document.querySelector('.qr-container img');
      const link = document.createElement('a');
      link.download = 'QR_${id}.png';
      link.href = img.src;
      link.click();
    }
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle('QR - ' + nombre)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function generarQRAfiliado(id) {
  const afiliado = getAfiliadoPorId(id);
  if (!afiliado) return null;

  const sheet = getSheet(CONFIG.SHEET_NAME_AFILIADOS);
  sheet.getRange(afiliado.row, CONFIG.COL_AFILIADO.QR_URL + 1).setValue(generarURLPaginaQR(id));

  return generarURLPaginaQR(id);
}

function generarQRTodos() {
  const sheet = getSheet(CONFIG.SHEET_NAME_AFILIADOS);
  const data = sheet.getDataRange().getValues();
  let generados = 0;

  for (let i = 1; i < data.length; i++) {
    const id = data[i][CONFIG.COL_AFILIADO.ID];
    if (!id) continue;
    generarQRAfiliado(id);
    generados++;
  }

  return { success: true, generados };
}
// ===== NOTIFICACIONES =====
function enviarNotificacion(destino, mensaje) {
  const canal = obtenerConfig('CANAL_PRINCIPAL') || 'TELEGRAM';

  switch (canal) {
    case 'WHATSAPP':
      return enviarWhatsApp(destino, mensaje);
    case 'SMS':
      return enviarSMS(destino, mensaje);
    case 'TELEGRAM':
    default:
      return enviarTelegram(destino, mensaje);
  }
}

function enviarTelegram(telefono, mensaje) {
  const token = obtenerConfig('TELEGRAM_BOT_TOKEN');
  const chatId = obtenerConfig('TELEGRAM_CHAT_ID');

  if (!token || !chatId) {
    console.log('Telegram no configurado. Token:' + !!token + ' ChatID:' + !!chatId);
    return false;
  }

  const text = mensaje + '\n\n— ' + obtenerConfig('META_MENSAJE');

  try {
    const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
    const payload = {
      chat_id: chatId,
      text: text,
      parse_mode: 'HTML',
      disable_web_page_preview: true
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.ok) {
      console.log('✅ Telegram: mensaje enviado a chat ' + chatId);
      return true;
    } else {
      console.log('❌ Telegram error: ' + JSON.stringify(result));
      return false;
    }
  } catch (err) {
    console.log('❌ Telegram exception: ' + err);
    return false;
  }
}

function enviarTelegramDirecto(chatId, mensaje) {
  const token = obtenerConfig('TELEGRAM_BOT_TOKEN');
  if (!token || !chatId) return false;

  const text = mensaje + '\n\n— ' + obtenerConfig('META_MENSAJE');

  try {
    const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
    const payload = {
      chat_id: chatId,
      text: text,
      parse_mode: 'HTML'
    };

    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return true;
  } catch (err) {
    console.log('❌ Telegram direct error: ' + err);
    return false;
  }
}

function enviarWhatsApp(telefono, mensaje) {
  const apiUrl = obtenerConfig('WHATSAPP_API_URL');
  const apiKey = obtenerConfig('WHATSAPP_API_KEY');

  if (!apiUrl || !apiKey) {
    console.log('WhatsApp no configurado');
    return false;
  }

  const numero = formatearNumero(telefono);
  const text = mensaje + '\n\n' + obtenerConfig('META_MENSAJE');

  try {
    const payload = {
      to: numero,
      text: text
    };

    UrlFetchApp.fetch(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      headers: {
        'Authorization': 'Bearer ' + apiKey
      }
    });

    console.log('✅ WhatsApp: mensaje enviado a ' + numero);
    return true;
  } catch (err) {
    console.log('❌ WhatsApp error: ' + err);
    return false;
  }
}

function enviarSMS(telefono, mensaje) {
  const accountSid = obtenerConfig('TWILIO_ACCOUNT_SID');
  const authToken = obtenerConfig('TWILIO_AUTH_TOKEN');
  const fromNumber = obtenerConfig('TWILIO_PHONE_NUMBER');

  if (!accountSid || !authToken || !fromNumber) {
    console.log('SMS/Twilio no configurado');
    return false;
  }

  const numero = formatearNumero(telefono);
  const text = mensaje + ' - ' + obtenerConfig('META_MENSAJE');

  try {
    const url = 'https://api.twilio.com/2010-04-01/Accounts/' + accountSid + '/Messages.json';
    const payload = {
      To: numero,
      From: fromNumber,
      Body: text.substring(0, 160)
    };

    const auth = Utilities.base64Encode(accountSid + ':' + authToken);

    UrlFetchApp.fetch(url, {
      method: 'post',
      payload: payload,
      muteHttpExceptions: true,
      headers: {
        'Authorization': 'Basic ' + auth
      }
    });

    console.log('✅ SMS enviado a ' + numero);
    return true;
  } catch (err) {
    console.log('❌ SMS error: ' + err);
    return false;
  }
}

function formatearNumero(telefono) {
  let num = String(telefono).replace(/[^0-9]/g, '');
  if (!num) return '';
  if (num.startsWith('0')) num = num.substring(1);
  if (!num.startsWith('57')) num = '57' + num;
  return '+' + num;
}

function enviarNotificacionGrupal(mensaje) {
  const canal = obtenerConfig('CANAL_PRINCIPAL') || 'TELEGRAM';

  switch (canal) {
    case 'WHATSAPP':
      return enviarWhatsApp(obtenerConfig('WHATSAPP_GROUP_ID'), mensaje);
    case 'TELEGRAM':
    default:
      return enviarTelegram('', mensaje);
  }
}

function probarNotificacion() {
  const fecha = Utilities.formatDate(new Date(), 'America/Bogota', 'dd/MM/yyyy HH:mm');
  const resultado = enviarNotificacionGrupal(
    '🔔 PRUEBA DE NOTIFICACIÓN\n\n' +
    'El sistema de notificaciones del Club Voleibol River funciona correctamente.\n' +
    'Fecha: ' + fecha
  );

  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  if (!ui) {
    Logger.log(resultado ? '✅ Notificación enviada' : '❌ Error al enviar');
    return resultado;
  }

  if (resultado) {
    ui.alert('✅ Notificación de prueba enviada correctamente.');
  } else {
    ui.alert('❌ Error al enviar notificación. Verifica la configuración en la pestaña Config.');
  }
}

function autorizadoWebhook(e) {
  const token = obtenerConfig('WEBHOOK_TOKEN');
  if (!token) return true;

  const given = (e.parameter && e.parameter.token) ||
    (e.parameter && e.parameter['x-api-key']) ||
    (e.parameter && e.parameter.api_key) || '';
  return String(given).trim() === String(token).trim();
}

function webhookPago(e) {
  if (!autorizadoWebhook(e)) {
    return jsonOut({ success: false, mensaje: 'Solicitud no autorizada' });
  }

  let data = {};
  try {
    data = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');
  } catch (err) {
    return jsonOut({ success: false, mensaje: 'Cuerpo JSON inválido' });
  }

  const result = registrarPago({
    tipoDoc: data.tipoDoc || data.tipo_documento || 'CC',
    numDoc: data.numDoc || data.numero_documento || '',
    monto: data.monto || data.mount || obtenerConfig('CUOTA_MENSUAL'),
    metodoPago: data.metodoPago || data.metodo_pago || 'Transferencia',
    comprobanteUrl: data.comprobanteUrl || data.comprobante_url || '',
    notas: data.notas || 'Pago vía webhook'
  });

  return jsonOut(result);
}

function registrarRespuestaEncuesta(e) {
  let data = {};
  try {
    data = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');
  } catch (err) {
    return jsonOut({ status: 'error', message: 'Cuerpo JSON inválido' });
  }
  return jsonOut(guardarEncuesta(data));
}

function guardarEncuesta(data) {
  data = data || {};

  const sheet = getSheet('Encuestas');
  if (sheet.getLastRow() === 0) {
    crearEstructuraEncuestas(sheet.getParent());
  }

  const fila = [
    Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss'),
    data.plan || '',
    data.name || '',
    data.phone || '',
    data.email || '',
    data.zone || '',
    data.fairPrice || '',
    data.whyNot || '',
    data.level || '',
    data.schedule || '',
    data.age || '',
    data.motivation || '',
    data.comments || ''
  ];

  sheet.appendRow(fila);

  return { status: 'success', message: 'Respuesta guardada' };
}

function responderPreinscripcion(e) {
  let data = {};
  try {
    data = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');
  } catch (err) {
    return jsonOut({ status: 'error', message: 'Cuerpo JSON inválido' });
  }
  return jsonOut(guardarInscripcion(data));
}

function guardarInscripcion(data) {
  data = data || {};

  const nombre = String(data.nombre || '').trim();
  const telefono = String(data.telefono || '').trim();
  if (!nombre || !telefono) {
    return { status: 'error', message: 'Nombre y celular son obligatorios' };
  }

  const sheet = getSheet('Inscripciones');
  if (sheet.getLastRow() === 0) {
    crearEstructuraInscripciones(sheet.getParent());
  }

  const fila = [
    Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss'),
    nombre,
    telefono,
    data.categoria || 'General',
    data.origen || 'sitio web'
  ];

  sheet.appendRow(fila);

  if (obtenerConfig('NOTIFICACION_HABILITADA') === 'SI') {
    enviarNotificacionGrupal(
      '🐆 Nuevo pre-inscrito desde la web\n' +
      'Nombre: ' + nombre + '\n' +
      'Celular: ' + telefono + '\n' +
      'Categoría: ' + (data.categoria || 'General')
    );
  }

  return { status: 'success', message: 'Inscripción recibida' };
}
// ===== FORMS =====
const FORM_REGISTRO_KEY = 'FORM_REGISTRO_ID';
const FORM_ASISTENCIA_KEY = 'FORM_ASISTENCIA_ID';

function crearFormularioRegistro() {
  const form = FormApp.create('Club Voleibol River - Registro de Afiliados');
  form.setDescription(
    'Completa este formulario para registrar a un nuevo deportista en el Club Deportivo de Voleibol River.\n\n' +
    'Los datos serán verificados y el afiliado recibirá su código QR al ser aprobado.'
  );
  form.setCollectEmail(true);
  form.setRequireLogin(false);
  form.setShowLinkToRespondAgain(true);

  form.addPageBreakItem().setTitle('Información Personal');

  const tipoDoc = form.addMultipleChoiceItem()
    .setTitle('Tipo de Documento')
    .setChoiceValues(['Tarjeta de Identidad', 'Cédula de Ciudadanía', 'Cédula de Extranjería', 'Pasaporte', 'Registro Civil'])
    .setRequired(true);

  const numDoc = form.addTextItem()
    .setTitle('Número de Documento')
    .setRequired(true)
    .setHelpText('Sin puntos ni espacios');

  const nombres = form.addTextItem()
    .setTitle('Nombres Completos')
    .setRequired(true);

  const apellidos = form.addTextItem()
    .setTitle('Apellidos Completos')
    .setRequired(true);

  const fechaNac = form.addDateItem()
    .setTitle('Fecha de Nacimiento')
    .setRequired(true);

  const telefono = form.addTextItem()
    .setTitle('Teléfono / Celular')
    .setRequired(true)
    .setHelpText('Ej: 3001234567');

  const email = form.addTextItem()
    .setTitle('Correo Electrónico')
    .setRequired(true);

  const direccion = form.addTextItem()
    .setTitle('Dirección de Residencia')
    .setHelpText('Barrio y dirección');

  form.addPageBreakItem().setTitle('Información del Acudiente (menores de edad)');

  const acudiente = form.addTextItem()
    .setTitle('Nombre del Acudiente')
    .setHelpText('Si el deportista es menor de edad');

  const telAcudiente = form.addTextItem()
    .setTitle('Teléfono del Acudiente');

  form.addPageBreakItem().setTitle('Información Deportiva');

  const categoria = form.addMultipleChoiceItem()
    .setTitle('Categoría')
    .setChoiceValues(['Infantil (8-12 años)', 'Juvenil (13-17 años)', 'Mayores (18+ años)', 'General'])
    .setRequired(true);

  const notas = form.addParagraphTextItem()
    .setTitle('Notas adicionales')
    .setHelpText('Alergias, condiciones médicas, etc.');

  const acepta = form.addCheckboxItem()
    .setTitle('Términos y Condiciones')
    .setHelpText('Acepto los términos y condiciones del Club Deportivo de Voleibol River')
    .setRequired(true);

  const terminos = [
    'Acepto cumplir los estatutos y reglamentos del Club.',
    'Autorizo el uso de mi imagen para fines promocionales del Club.',
    'Certifico que los datos proporcionados son verídicos.',
    'Acepto las normas de disciplina deportiva del Club.'
  ];
  acepta.setChoiceValues(terminos);

  const trigger = ScriptApp.newTrigger('procesarFormularioRegistro')
    .forForm(form)
    .onFormSubmit()
    .create();

  form.setConfirmationMessage(
    '¡Registro completado exitosamente!\n\n' +
    'Recibirás una notificación con tu código QR de afiliado.\n' +
    'Bienvenido al Club Deportivo de Voleibol River 🏐'
  );

  PropertiesService.getScriptProperties().setProperty(FORM_REGISTRO_KEY, String(form.getId()));
  return form.getId();
}

function obtenerOCrearFormularioRegistro() {
  const id = PropertiesService.getScriptProperties().getProperty(FORM_REGISTRO_KEY);
  if (id) {
    try {
      FormApp.openById(id);
      return id;
    } catch (e) {
      PropertiesService.getScriptProperties().deleteProperty(FORM_REGISTRO_KEY);
    }
  }
  return crearFormularioRegistro();
}

function crearFormularioAsistencia() {
  const form = FormApp.create('Club Voleibol River - Registro de Asistencia');
  form.setDescription(
    'Registra tu asistencia a entrenamientos, partidos y eventos del Club Deportivo de Voleibol River.'
  );
  form.setRequireLogin(false);
  form.setShowLinkToRespondAgain(true);

  const idAfiliado = form.addTextItem()
    .setTitle('ID de Afiliado')
    .setRequired(true)
    .setHelpText('Ingresa el ID que aparece en tu código QR');

  const tipo = form.addMultipleChoiceItem()
    .setTitle('Tipo de Actividad')
    .setChoiceValues(['Entrenamiento', 'Partido', 'Evento', 'Otro'])
    .setRequired(true);

  const trigger = ScriptApp.newTrigger('procesarFormularioAsistencia')
    .forForm(form)
    .onFormSubmit()
    .create();

  PropertiesService.getScriptProperties().setProperty(FORM_ASISTENCIA_KEY, String(form.getId()));
  return form.getId();
}

function obtenerOCrearFormularioAsistencia() {
  const id = PropertiesService.getScriptProperties().getProperty(FORM_ASISTENCIA_KEY);
  if (id) {
    try {
      FormApp.openById(id);
      return id;
    } catch (e) {
      PropertiesService.getScriptProperties().deleteProperty(FORM_ASISTENCIA_KEY);
    }
  }
  return crearFormularioAsistencia();
}

function procesarFormularioRegistro(e) {
  const responses = e.response.getItemResponses();
  const data = {};

  responses.forEach(r => {
    const title = r.getItem().getTitle();
    const value = r.getResponse();
    data[title] = value;
  });

  let fechaNac = '';
  const nac = data['Fecha de Nacimiento'];
  if (nac instanceof Date) {
    fechaNac = Utilities.formatDate(nac, 'America/Bogota', 'yyyy-MM-dd');
  } else if (nac) {
    fechaNac = String(nac);
  }

  const result = registrarAfiliado({
    tipoDoc: data['Tipo de Documento'] || 'CC',
    numDoc: String(data['Número de Documento'] || ''),
    nombres: (data['Nombres Completos'] || '').toUpperCase(),
    apellidos: (data['Apellidos Completos'] || '').toUpperCase(),
    fechaNac: fechaNac,
    telefono: data['Teléfono / Celular'] || '',
    email: data['Correo Electrónico'] || '',
    direccion: data['Dirección de Residencia'] || '',
    acudiente: data['Nombre del Acudiente'] || '',
    telAcudiente: data['Teléfono del Acudiente'] || '',
    categoria: data['Categoría'] || 'General',
    notas: data['Notas adicionales'] || ''
  });

  console.log('Registro de formulario:', JSON.stringify(result));

  if (result.success) {
    const email = data['Correo Electrónico'];
    const urlQR = generarURLPaginaQR(result.id);
    try {
      MailApp.sendEmail({
        to: email,
        subject: '🏐 Club Voleibol River - Tu código QR de afiliado (ID: ' + result.id + ')',
        htmlBody:
          '<p>Hola <b>' + result.nombres + ' ' + result.apellidos + '</b>,</p>' +
          '<p>Tu registro fue exitoso. Tu ID de afiliado es <b>' + result.id + '</b>.</p>' +
          '<p>Guarda tu código QR (descárgalo o imprímelo en este enlace):</p>' +
          '<p><a href="' + urlQR + '">' + urlQR + '</a></p>' +
          '<p>Preséntalo en el escenario para registrar tu asistencia. 🏐</p>'
      });
    } catch (err) {
      console.log('No se pudo enviar email con QR: ' + err);
    }
  }

  return result;
}

function procesarFormularioAsistencia(e) {
  const responses = e.response.getItemResponses();
  const data = {};

  responses.forEach(r => {
    const title = r.getItem().getTitle();
    data[title] = r.getResponse();
  });

  const afiliado = getAfiliadoPorId(data['ID de Afiliado']);

  if (afiliado) {
    const result = registrarAsistencia(data['ID de Afiliado'], data['Tipo de Actividad']);
    console.log('Asistencia por formulario:', JSON.stringify(result));
  } else {
    console.log('Asistencia no registrada: ID no encontrado - ' + data['ID de Afiliado']);
  }
}