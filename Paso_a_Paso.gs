// ============================================================================
// SISTEMA PASO A PASO - FITO v7.0
// Con Dashboard + Email + Alertas + Derivaciones + Historial
// ============================================================================

const CONFIG = {
  SPREADSHEET_ID:          SpreadsheetApp.getActive().getId(),
  FOLDER_PARTICIPANTES_ID: "1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU",
  KOBO_API_KEY:            "64cc018b88067397addd36b09288be8b6539cf39",
  KOBO_ASSET_ID:           "abHRWdRnPhKwzPQBajc7RZ",
  KOBO_URL:                "https://kf.kobotoolbox.org/api/v2",
  HOJA:                    "Maestro",
  ADMIN_EMAIL:             "adrian@creamosguatemla.org",
  COLORES: {
    "Perfil A": { fondo: "#d9ead3", texto: "#274e13" },
    "Perfil B": { fondo: "#cfe2f3", texto: "#1c4587" },
    "Perfil C": { fondo: "#fff2cc", texto: "#7f6000" },
    "Perfil D": { fondo: "#f4cccc", texto: "#660000" }
  },
  COL: {
    ID: 1, FECHA: 2, NOMBRE: 3, DPI: 4, EDAD: 5, GENERO: 6, TELEFONO: 7,
    ZONA: 8, EMAIL: 9, EDUCACION: 10, LABORAL: 11, FORTALEZAS: 12, OBJETIVO: 13,
    PERFIL: 14, PRIORIDAD: 15, PUNTAJE: 16, DIM1: 17, DIM2: 18, DIM3: 19,
    DIM4: 20, DIM5: 21, DIM6: 22, ESTADO: 23, CARPETA_ID: 24, DOC_ID: 25, DOC_URL: 26
  }
};

// ============================================================================
// MENÚ
// ============================================================================

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📊 PASO A PASO')
      .addItem('📥 Instalar Sistema', 'instalar')
      .addSeparator()
      .addItem('🔄 Sincronizar Kobo', 'sincronizar')
      .addItem('🎨 Colorear Tabla', 'colorearTabla')
      .addItem('📋 Ver Ficha', 'verFicha')
      .addSeparator()
      .addItem('📊 Dashboard', 'abrirDashboard')
      .addItem('📈 Analytics', 'abrirAnalytics')
      .addItem('➡️ Derivaciones', 'verDerivaciones')
      .addItem('🔁 Restaurar desde Drive', 'restaurarDesdeDrive')
      .addSeparator()
      .addItem('🧪 Probar Kobo', 'probarKobo')
      .addItem('⚙️ Configuración', 'abrirConfiguracion')
      .addSeparator()
      .addItem('🗑️ Desinstalar & Limpiar', 'desinstalarYLimpiar')
      .addToUi();
  } catch(e) {
    // Sin contexto UI
  }
}

// ============================================================================
// INSTALAR
// ============================================================================

function instalar() {
  try {
    const ss = SpreadsheetApp.getActive();

    // Crear hoja Maestro
    let hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) {
      hoja = ss.insertSheet(CONFIG.HOJA);
    }
    if (hoja.getLastRow() === 0) {
      const headers = [
        'ID_Creamos','Fecha_Registro','Nombre_Completo','DPI','Edad','Género',
        'Teléfono','Zona','Email','Nivel_Educativo','Situación_Laboral','Fortalezas',
        'Objetivo_Laboral','Perfil_Asignado','Prioridad','Puntaje_Total',
        'Dim_Educativo','Dim_Laboral','Dim_Digital','Dim_Vocacional',
        'Dim_Barreras','Dim_Apoyo','Estado','Carpeta_Drive_ID','Doc_Perfil_ID','Doc_Perfil_URL'
      ];
      hoja.appendRow(headers);
      hoja.getRange(1, 1, 1, headers.length)
        .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
      hoja.setFrozenRows(1);
    }

    // Crear hoja Derivaciones
    if (!ss.getSheetByName('Derivaciones')) {
      const deriv = ss.insertSheet('Derivaciones');
      deriv.appendRow(['Fecha','ID_Participante','Nombre','Tipo','Destino','Motivo','Estado','Responsable','Fecha_Seguimiento','Notas']);
      deriv.getRange(1, 1, 1, 10).setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');
      deriv.setFrozenRows(1);
      deriv.setColumnWidth(3, 180);
      deriv.setColumnWidth(5, 220);
      deriv.setColumnWidth(6, 250);
    }

    // Crear hoja Analytics
    if (!ss.getSheetByName('Analytics')) {
      const an = ss.insertSheet('Analytics');
      an.appendRow(['ANALYTICS - PASO A PASO']);
      an.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#6a1b9a');
      an.setFrozenRows(1);
    }

    // Crear hoja Log
    if (!ss.getSheetByName('Log')) {
      const log = ss.insertSheet('Log');
      log.appendRow(['Fecha_Hora','Hoja','Fila','Columna','Valor_Anterior','Valor_Nuevo','Usuario']);
      log.getRange(1, 1, 1, 7).setBackground('#9c27b0').setFontColor('#ffffff').setFontWeight('bold');
      log.setFrozenRows(1);
    }

    // Crear hoja Dashboard
    if (!ss.getSheetByName('Dashboard')) {
      const dash = ss.insertSheet('Dashboard');
      dash.appendRow(['DASHBOARD - PASO A PASO']);
      dash.getRange(1, 1).setFontSize(16).setFontWeight('bold');
    }

    configurarTriggers();

    SpreadsheetApp.getUi().alert(
      '✅ Sistema v7.0 instalado\n\n' +
      '✨ Nuevas características:\n' +
      '📧 Email semanal los lunes\n' +
      '🚨 Alertas de casos dormidos\n' +
      '➡️ Registro de derivaciones\n' +
      '📜 Historial de cambios\n' +
      '📊 Dashboard automático\n\n' +
      'Próximo: 🔄 Sincronizar Kobo'
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

function configurarTriggers() {
  const TRIGGERS = ['sincronizarAutomatico', 'enviarReporteSemanal', 'verificarCasosDormidos', 'manejarEdicion'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (TRIGGERS.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  // Sincronización diaria 8 AM
  ScriptApp.newTrigger('sincronizarAutomatico').timeBased().atHour(8).everyDays(1).create();

  // Email semanal lunes 9 AM
  ScriptApp.newTrigger('enviarReporteSemanal').timeBased().atHour(9).onWeekDay(ScriptApp.WeekDay.MONDAY).create();

  // Verificar casos dormidos diariamente 10 AM
  ScriptApp.newTrigger('verificarCasosDormidos').timeBased().atHour(10).everyDays(1).create();

  // onEdit INSTALABLE — necesario para poder abrir Google Docs con permisos
  ScriptApp.newTrigger('manejarEdicion').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}

function sincronizarAutomatico() {
  sincronizar(true);
  actualizarDashboard();
}

// ============================================================================
// SINCRONIZAR KOBO
// ============================================================================

function sincronizar(silencioso) {
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ Instala primero');
      return;
    }

    const apiKey = PropertiesService.getUserProperties().getProperty('KOBO_API_KEY') || CONFIG.KOBO_API_KEY;
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json',
      { headers: { Authorization: 'Token ' + apiKey }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ Error Kobo ' + resp.getResponseCode());
      return;
    }

    const registros = JSON.parse(resp.getContentText()).results || [];
    if (!registros.length) {
      if (!silencioso) SpreadsheetApp.getUi().alert('⚠️ Sin respuestas en Kobo');
      return;
    }

    const existentes = new Set();
    if (hoja.getLastRow() > 1) {
      hoja.getRange(2, CONFIG.COL.ID, hoja.getLastRow() - 1, 1)
        .getValues().forEach(r => { if (r[0]) existentes.add(String(r[0])); });
    }

    const folderId = PropertiesService.getUserProperties().getProperty('FOLDER_ID') || CONFIG.FOLDER_PARTICIPANTES_ID;
    const folderBase = DriveApp.getFolderById(folderId);
    let agregados = 0;

    registros.forEach(r => {
      const id = String(r['creamos_id'] || r['_id'] || '');
      if (!id || existentes.has(id)) return;

      const nombre = r['nombre_completo_del_la_participante'] || 'Participante ' + id;
      const expediente = crearExpediente(folderBase, id, nombre, r);

      const fila = [
        id, new Date(), nombre,
        r['numero_de_dpi_opcional'] || '',
        r['edad'] || '',
        r['genero'] || '',
        r['numero_de_telefono'] || '',
        r['lugar_de_residencia'] || '',
        r['correo_electronico_opcional'] || '',
        r['cual_es_el_ultimo_grado_que_completaste'] || '',
        r['cual_es_tu_situacion_laboral_actual'] || '',
        r['que_sabes_hacer_bien'] || '',
        r['que_tipo_de_empleo_estas_buscando_especificamente'] || '',
        normalizarPerfil(r['perfil_asignado']),
        normalizarPrioridad(r['prioridad_caso']),
        Number(r['puntaje_total_60']) || 0,
        Number(r['dimension_1_capital_educativo']) || 0,
        Number(r['dimension_2_capital_laboral']) || 0,
        Number(r['dimension_3_habilidades_digitales']) || 0,
        Number(r['dimension_4_claridad_vocacional']) || 0,
        Number(r['dimension_5_barreras_estructurales']) || 0,
        Number(r['dimension_6_red_apoyo']) || 0,
        'Orientación',
        expediente.carpetaId,
        expediente.docId,
        expediente.docUrl
      ];

      hoja.appendRow(fila);
      colorearFila(hoja, hoja.getLastRow(), fila[CONFIG.COL.PERFIL - 1]);
      registrarDerivacionesAutomaticas(ss, id, nombre, fila, r);
      agregados++;
    });

    actualizarDashboard();
    actualizarAnalytics(ss);

    if (!silencioso) {
      SpreadsheetApp.getUi().alert('✅ Sincronizado\n👤 ' + agregados + ' nuevos\n📊 ' + registros.length + ' en Kobo');
    }
  } catch(e) {
    if (!silencioso) SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================================================
// CREAR EXPEDIENTE
// ============================================================================

function crearExpediente(folderBase, id, nombre, datos) {
  try {
    const carpeta = folderBase.createFolder(id + ' - ' + nombre);
    const doc = DocumentApp.create('Perfil - ' + nombre);
    DriveApp.getFileById(doc.getId()).moveTo(carpeta);

    const body = doc.getBody();
    body.clear();

    const t = body.appendParagraph('PERFIL DEL PARTICIPANTE');
    t.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph('ID: ' + id + '   |   Perfil: ' + (datos['perfil_asignado'] || '—') + '   |   Prioridad: ' + (datos['prioridad_caso'] || '—'));
    body.appendParagraph('');

    const s1 = body.appendParagraph('DATOS PERSONALES');
    s1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Nombre:    ' + nombre);
    body.appendParagraph('DPI:       ' + (datos['numero_de_dpi_opcional'] || '—'));
    body.appendParagraph('Edad:      ' + (datos['edad'] || '—'));
    body.appendParagraph('Género:    ' + (datos['genero'] || '—'));
    body.appendParagraph('Zona:      ' + (datos['lugar_de_residencia'] || '—'));
    body.appendParagraph('Teléfono:  ' + (datos['numero_de_telefono'] || '—'));
    body.appendParagraph('');

    const s2 = body.appendParagraph('DIAGNÓSTICO');
    s2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Educativo:   ' + (datos['dimension_1_capital_educativo'] || 0) + '/10');
    body.appendParagraph('Laboral:     ' + (datos['dimension_2_capital_laboral'] || 0) + '/10');
    body.appendParagraph('Digital:     ' + (datos['dimension_3_habilidades_digitales'] || 0) + '/10');
    body.appendParagraph('Vocacional:  ' + (datos['dimension_4_claridad_vocacional'] || 0) + '/10');
    body.appendParagraph('Barreras:    ' + (datos['dimension_5_barreras_estructurales'] || 0) + '/10');
    body.appendParagraph('Red Apoyo:   ' + (datos['dimension_6_red_apoyo'] || 0) + '/10');
    body.appendParagraph('');

    const s3 = body.appendParagraph('PERFIL PROFESIONAL');
    s3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Educación:         ' + (datos['cual_es_el_ultimo_grado_que_completaste'] || '—'));
    body.appendParagraph('Situación laboral: ' + (datos['cual_es_tu_situacion_laboral_actual'] || '—'));
    body.appendParagraph('Fortalezas:        ' + (datos['que_sabes_hacer_bien'] || '—'));
    body.appendParagraph('Objetivo laboral:  ' + (datos['que_tipo_de_empleo_estas_buscando_especificamente'] || '—'));
    body.appendParagraph('');

    const s4 = body.appendParagraph('PLAN DE ACCIÓN');
    s4.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('[Definir acciones para este participante]');
    body.appendParagraph('');

    const s5 = body.appendParagraph('HISTORIAL DE CAMBIOS');
    s5.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Fecha             | Campo             | Cambio                          | Usuario');
    body.appendParagraph('─────────────────────────────────────────────────────────────────────────');

    doc.saveAndClose();

    return { carpetaId: carpeta.getId(), docId: doc.getId(), docUrl: 'https://docs.google.com/document/d/' + doc.getId() + '/edit' };
  } catch(e) {
    return { carpetaId: '', docId: '', docUrl: '' };
  }
}

// ============================================================================
// COLOREAR + ALERTAS DE CASOS DORMIDOS
// ============================================================================

function colorearTabla() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Sin datos');
      return;
    }
    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 26).getValues();
    datos.forEach((fila, i) => colorearFila(hoja, i + 2, fila[CONFIG.COL.PERFIL - 1]));
    SpreadsheetApp.getUi().alert('✅ Coloreado\n🟢 A  🔵 B  🟡 C  🔴 D');
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

function colorearFila(hoja, numFila, perfil) {
  const c = CONFIG.COLORES[perfil];
  const rango = hoja.getRange(numFila, 1, 1, 26);
  if (c) {
    rango.setBackground(c.fondo).setFontColor(c.texto);
  } else {
    rango.setBackground('#ffffff').setFontColor('#333333');
  }
}

function verificarCasosDormidos() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) return;

    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 26).getValues();
    const hoy = new Date();
    const casosDormidos = [];

    datos.forEach((fila, i) => {
      if (!fila[0]) return;
      const fechaUltimoCambio = new Date(fila[CONFIG.COL.FECHA - 1]);
      const diasSinContacto = Math.floor((hoy - fechaUltimoCambio) / (1000 * 60 * 60 * 24));

      if (diasSinContacto > 30) {
        const numFila = i + 2;
        hoja.getRange(numFila, 1, 1, 26).setBackground('#8b0000').setFontColor('#ffffff');
        casosDormidos.push({
          nombre: fila[CONFIG.COL.NOMBRE - 1],
          dias: diasSinContacto,
          estado: fila[CONFIG.COL.ESTADO - 1]
        });
      }
    });

    if (casosDormidos.length > 0) {
      const adminEmail = PropertiesService.getUserProperties().getProperty('ADMIN_EMAIL') || CONFIG.ADMIN_EMAIL;
      const asunto = '🚨 Alerta: ' + casosDormidos.length + ' casos sin seguimiento';
      let body = 'Los siguientes participantes llevan más de 30 días sin cambio de estado:\n\n';
      casosDormidos.forEach(c => {
        body += '👤 ' + c.nombre + ' (' + c.dias + ' días)\n   Estado: ' + c.estado + '\n\n';
      });
      body += 'Abre el sheet y revisa los casos marcados en rojo oscuro.\n\nPaso a Paso Sistema';
      MailApp.sendEmail(adminEmail, asunto, body);
    }
  } catch(e) {
    // Error silencioso en trigger
  }
}

function enviarReporteSemanal() {
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) return;

    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 26).getValues().filter(r => r[0]);
    const hoy = new Date();
    const hace7dias = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);

    const nuevos = datos.filter(r => {
      const fecha = new Date(r[CONFIG.COL.FECHA - 1]);
      return fecha >= hace7dias;
    });

    const criticos = datos.filter(r => r[CONFIG.COL.PRIORIDAD - 1] === 'CRÍTICO');
    const dormidos = datos.filter(r => {
      const diasSinContacto = Math.floor((hoy - new Date(r[CONFIG.COL.FECHA - 1])) / (1000 * 60 * 60 * 24));
      return diasSinContacto > 30;
    });

    const adminEmail = PropertiesService.getUserProperties().getProperty('ADMIN_EMAIL') || CONFIG.ADMIN_EMAIL;
    const asunto = '📊 Reporte semanal - Paso a Paso';
    let body = '═════════════════════════════════════\n';
    body += '📊 REPORTE SEMANAL - PASO A PASO\n';
    body += '═════════════════════════════════════\n\n';
    body += 'Semana: ' + Utilities.formatDate(hace7dias, Session.getScriptTimeZone(), 'dd/MM') + ' - ' + Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '\n\n';

    body += '✨ NUEVOS PARTICIPANTES: ' + nuevos.length + '\n';
    nuevos.forEach(r => {
      body += '  • ' + r[CONFIG.COL.NOMBRE - 1] + ' (' + r[CONFIG.COL.PERFIL - 1] + ')\n';
    });

    body += '\n🔴 CASOS CRÍTICOS: ' + criticos.length + '\n';
    criticos.forEach(r => {
      body += '  • ' + r[CONFIG.COL.NOMBRE - 1] + ' - ' + r[CONFIG.COL.ESTADO - 1] + '\n';
    });

    body += '\n⏰ PENDIENTES (30+ días): ' + dormidos.length + '\n';
    dormidos.forEach(r => {
      const dias = Math.floor((hoy - new Date(r[CONFIG.COL.FECHA - 1])) / (1000 * 60 * 60 * 24));
      body += '  • ' + r[CONFIG.COL.NOMBRE - 1] + ' (' + dias + ' días sin cambio)\n';
    });

    body += '\n═════════════════════════════════════\n';
    body += 'Total participantes: ' + datos.length + '\n';
    body += 'Abre el sheet para ver detalles.\n\n';
    body += 'Sistema Paso a Paso';

    MailApp.sendEmail(adminEmail, asunto, body);
  } catch(e) {
    // Error silencioso
  }
}

// ============================================================================
// HISTORIAL DE CAMBIOS
// onEdit simple NO puede abrir Docs (sin permisos).
// manejarEdicion() es un trigger INSTALABLE con permisos completos.
// Se instala automáticamente al ejecutar 📥 Instalar Sistema.
// ============================================================================

function onEdit(e) {
  // Solo registra en Log — no abre Docs (sin permisos en trigger simple)
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== CONFIG.HOJA) return;
    const fila = e.range.getRow();
    const columna = e.range.getColumn();
    if (fila < 2) return;
    const logSheet = e.source.getSheetByName('Log');
    if (logSheet && (e.value || e.oldValue)) {
      logSheet.appendRow([new Date(), CONFIG.HOJA, fila, columna, e.oldValue || '', e.value || '', Session.getEffectiveUser().getEmail()]);
    }
  } catch(err) {}
}

function manejarEdicion(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== CONFIG.HOJA) return;

    const range = e.range;
    const fila = range.getRow();
    const columna = range.getColumn();
    const valorNuevo = e.value || '';
    const valorAnterior = e.oldValue || '';

    if (fila < 2 || !valorNuevo || valorNuevo === valorAnterior) return;

    // Registrar en log
    const logSheet = e.source.getSheetByName('Log');
    if (logSheet) {
      const ahora = new Date();
      const usuario = Session.getEffectiveUser().getEmail();
      logSheet.appendRow([ahora, CONFIG.HOJA, fila, columna, valorAnterior, valorNuevo, usuario]);
    }

    // Sincronizar con Google Doc si la columna es un campo de datos
    const C = CONFIG.COL;
    const COLS_SYNC = [
      C.NOMBRE, C.DPI, C.EDAD, C.GENERO, C.TELEFONO, C.ZONA,
      C.EDUCACION, C.LABORAL, C.FORTALEZAS, C.OBJETIVO,
      C.PERFIL, C.PRIORIDAD, C.ESTADO
    ];
    if (COLS_SYNC.includes(columna)) {
      const datos = sheet.getRange(fila, 1, 1, 26).getValues()[0];
      const docId = datos[C.DOC_ID - 1];
      if (docId) {
        sincronizarConDocumento(docId, datos, columna, valorNuevo, valorAnterior);
      }

      // Si cambió ESTADO → Calendar + Email al participante
      if (columna === C.ESTADO) {
        const nombre   = datos[C.NOMBRE - 1] || '';
        const email    = datos[C.EMAIL - 1] || '';
        const docUrl   = datos[C.DOC_URL - 1] || '';
        crearEventoCalendario(nombre, valorNuevo, docUrl);
        if (email) notificarParticipante(nombre, email, valorNuevo, datos);
      }
    }
  } catch(e) {
    // Fallo silencioso
  }
}

function sincronizarConDocumento(docId, datos, columna, valorNuevo, valorAnterior) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const C = CONFIG.COL;

    // Mapa: número de columna → {patrón regex, prefijo del campo en el doc}
    const mapa = {};
    mapa[C.NOMBRE]     = { p: 'Nombre:.*',            pre: 'Nombre:    ' };
    mapa[C.DPI]        = { p: 'DPI:.*',               pre: 'DPI:       ' };
    mapa[C.EDAD]       = { p: 'Edad:.*',              pre: 'Edad:      ' };
    mapa[C.GENERO]     = { p: 'Género:.*',            pre: 'Género:    ' };
    mapa[C.TELEFONO]   = { p: 'Teléfono:.*',          pre: 'Teléfono:  ' };
    mapa[C.ZONA]       = { p: 'Zona:.*',              pre: 'Zona:      ' };
    mapa[C.EDUCACION]  = { p: 'Educación:.*',         pre: 'Educación:         ' };
    mapa[C.LABORAL]    = { p: 'Situación laboral:.*', pre: 'Situación laboral: ' };
    mapa[C.FORTALEZAS] = { p: 'Fortalezas:.*',        pre: 'Fortalezas:        ' };
    mapa[C.OBJETIVO]   = { p: 'Objetivo laboral:.*',  pre: 'Objetivo laboral:  ' };

    // Reemplazar el campo correspondiente
    const campo = mapa[columna];
    if (campo) {
      body.replaceText(campo.p, campo.pre + (valorNuevo || '—'));
    }

    // Para PERFIL o PRIORIDAD, actualizar la línea de encabezado ID | Perfil | Prioridad
    if (columna === C.PERFIL || columna === C.PRIORIDAD) {
      const id = datos[C.ID - 1] || '';
      const perfil = datos[C.PERFIL - 1] || '—';
      const prioridad = datos[C.PRIORIDAD - 1] || '—';
      body.replaceText('ID: .*\\|.*Prioridad:.*', 'ID: ' + id + '   |   Perfil: ' + perfil + '   |   Prioridad: ' + prioridad);
    }

    // Si cambió NOMBRE también actualizar título del documento
    if (columna === C.NOMBRE && valorNuevo) {
      doc.setName('Perfil - ' + valorNuevo);
    }

    // Registrar en sección HISTORIAL del documento
    const fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const usuario = Session.getEffectiveUser().getEmail().split('@')[0];
    const nombreCampo = getNombreCampo(columna);
    const lineaLog = fechaHora + ' | ' + nombreCampo.padEnd(17) + ' | ' + (valorAnterior || '—').toString().substring(0, 30).padEnd(31) + ' → ' + (valorNuevo || '—').toString().substring(0, 20) + ' | ' + usuario;

    if (body.getText().includes('HISTORIAL DE CAMBIOS')) {
      body.appendParagraph(lineaLog);
    }

    doc.saveAndClose();
  } catch(e) {
    // Error silencioso - doc puede haber sido movido o eliminado
  }
}

function getNombreCampo(columna) {
  const C = CONFIG.COL;
  const n = {};
  n[C.NOMBRE] = 'Nombre'; n[C.DPI] = 'DPI'; n[C.EDAD] = 'Edad';
  n[C.GENERO] = 'Género'; n[C.TELEFONO] = 'Teléfono'; n[C.ZONA] = 'Zona';
  n[C.EDUCACION] = 'Educación'; n[C.LABORAL] = 'Situación Laboral';
  n[C.FORTALEZAS] = 'Fortalezas'; n[C.OBJETIVO] = 'Objetivo';
  n[C.PERFIL] = 'Perfil'; n[C.PRIORIDAD] = 'Prioridad'; n[C.ESTADO] = 'Estado';
  return n[columna] || 'Campo ' + columna;
}

// ============================================================================
// DERIVACIONES
// ============================================================================

function abrirFormDerivacion() {
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona un participante primero');
      return;
    }

    const range = ss.getActiveRange();
    if (range.getRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila de participante');
      return;
    }

    const datos = hoja.getRange(range.getRow(), 1, 1, 26).getValues()[0];
    const participanteID = datos[0] || '';
    const participanteNombre = datos[CONFIG.COL.NOMBRE - 1] || '';

    if (!participanteID) {
      SpreadsheetApp.getUi().alert('⚠️ Participante sin ID');
      return;
    }

    const html = HtmlService.createHtmlOutput(
      '<style>body{font-family:Arial;padding:16px;background:#f5f5f5;font-size:13px}' +
      'label{display:block;font-weight:bold;margin:12px 0 4px;color:#333}' +
      'input,textarea,select{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;font-size:12px}' +
      'textarea{height:80px;resize:vertical}' +
      'button{margin-top:14px;background:#1f73e6;color:#fff;padding:10px;border:none;border-radius:4px;cursor:pointer;width:100%;font-weight:bold}' +
      '.info{background:#e3f2fd;color:#1565c0;padding:10px;border-radius:4px;margin-bottom:12px;font-size:12px}' +
      '</style>' +
      '<div class="info">📋 Derivación de: <strong>' + participanteNombre + '</strong></div>' +
      '<label>Organización destino:</label>' +
      '<select id="org"><option>Selecciona...</option><option>Otro programa Creamos</option><option>Entidad pública</option><option>ONG externa</option><option>Sector privado</option><option>Otra</option></select>' +
      '<label>Motivo de la derivación:</label>' +
      '<textarea id="motivo" placeholder="Razón por la que se derivó..."></textarea>' +
      '<label>Resultado esperado:</label>' +
      '<textarea id="resultado" placeholder="Qué se espera lograr..."></textarea>' +
      '<button onclick="guardarDer()">➡️ Registrar Derivación</button>' +
      '<script>function guardarDer(){var o=document.getElementById(\'org\').value,m=document.getElementById(\'motivo\').value,r=document.getElementById(\'resultado\').value;if(o==="Selecciona..."||!m||!r){alert("Completa todos los campos");return;}google.script.run.guardarDerivacion("' + participanteID + '","' + participanteNombre + '",o,m,r);alert("✅ Derivación registrada");google.script.host.close();}</script>'
    );

    SpreadsheetApp.getUi().showSidebar(html);
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

function guardarDerivacion(id, nombre, organizacion, motivo, resultado) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('Derivaciones');
    if (!sheet) return;

    sheet.appendRow([
      new Date(),
      id,
      nombre,
      organizacion,
      motivo,
      resultado,
      '',
      ''
    ]);
  } catch(e) {
    // Error silencioso
  }
}

// ============================================================================
// DASHBOARD AUTOMÁTICO
// ============================================================================

function abrirDashboard() {
  try {
    actualizarDashboard();
    const ss = SpreadsheetApp.getActive();
    ss.setActiveSheet(ss.getSheetByName('Dashboard'));
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

function actualizarDashboard() {
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    const dash = ss.getSheetByName('Dashboard');
    if (!hoja || !dash) return;

    dash.clear();

    // Título
    dash.appendRow(['DASHBOARD - PASO A PASO']);
    dash.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1f73e6');

    // Fecha de actualización
    const ahora = new Date();
    dash.appendRow(['Actualizado: ' + Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')]);
    dash.appendRow(['']);

    // Datos
    const datos = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 0), 26).getValues().filter(r => r[0]);

    // Resumen
    dash.appendRow(['RESUMEN GENERAL']);
    dash.getRange(dash.getLastRow(), 1).setFontWeight('bold').setFontColor('#1f73e6');
    dash.appendRow(['Total de participantes', datos.length]);
    dash.appendRow(['']);

    // Por perfil — normaliza valores de Kobo antes de contar
    dash.appendRow(['POR PERFIL']);
    dash.getRange(dash.getLastRow(), 1).setFontWeight('bold').setFontColor('#1f73e6');
    const perfiles = { 'Perfil A': 0, 'Perfil B': 0, 'Perfil C': 0, 'Perfil D': 0, 'Sin perfil': 0 };
    datos.forEach(r => {
      const p = normalizarPerfil(r[CONFIG.COL.PERFIL - 1]) || 'Sin perfil';
      perfiles[p] = (perfiles[p] || 0) + 1;
    });
    Object.entries(perfiles).forEach(([k, v]) => {
      dash.appendRow([k, v]);
      colorearFilaDash(dash, dash.getLastRow(), k);
    });
    dash.appendRow(['']);

    // Por estado
    dash.appendRow(['POR ESTADO']);
    dash.getRange(dash.getLastRow(), 1).setFontWeight('bold').setFontColor('#1f73e6');
    const estados = {};
    datos.forEach(r => {
      const e = r[CONFIG.COL.ESTADO - 1] || 'Sin estado';
      estados[e] = (estados[e] || 0) + 1;
    });
    Object.entries(estados).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => {
      dash.appendRow([k, v]);
    });
    dash.appendRow(['']);

    // Por prioridad — normaliza valores de Kobo antes de contar
    dash.appendRow(['POR PRIORIDAD']);
    dash.getRange(dash.getLastRow(), 1).setFontWeight('bold').setFontColor('#1f73e6');
    const prioridades = { 'CRÍTICO': 0, 'ALTO': 0, 'MEDIO': 0, 'BAJO': 0 };
    datos.forEach(r => {
      const p = normalizarPrioridad(r[CONFIG.COL.PRIORIDAD - 1]) || 'Sin prioridad';
      prioridades[p] = (prioridades[p] || 0) + 1;
    });
    [['CRÍTICO','#ffcdd2'], ['ALTO','#ffe0b2'], ['MEDIO','#fff9c4'], ['BAJO','#c8e6c9']].forEach(([k, color]) => {
      dash.appendRow([k, prioridades[k] || 0]);
      dash.getRange(dash.getLastRow(), 1, 1, 2).setBackground(color);
    });
    dash.appendRow(['']);

    // Puntaje promedio
    dash.appendRow(['DIAGNÓSTICO']);
    dash.getRange(dash.getLastRow(), 1).setFontWeight('bold').setFontColor('#1f73e6');
    const puntajes = datos.map(r => Number(r[CONFIG.COL.PUNTAJE - 1])).filter(n => n > 0);
    const prom = puntajes.length ? Math.round(puntajes.reduce((a, b) => a + b, 0) / puntajes.length) : 0;
    dash.appendRow(['Puntaje promedio', prom + ' / 60']);
    dash.appendRow(['Máximo', Math.max(...puntajes) || 0]);
    dash.appendRow(['Mínimo', Math.min(...puntajes) || 0]);

    // Ancho de columnas
    dash.setColumnWidth(1, 250);
    dash.setColumnWidth(2, 100);
  } catch(e) {
    // Error silencioso
  }
}

function colorearFilaDash(sheet, fila, perfil) {
  const c = CONFIG.COLORES[perfil];
  if (c) {
    sheet.getRange(fila, 1, 1, 2).setBackground(c.fondo).setFontColor(c.texto);
  }
}

// ============================================================================
// VER FICHA
// ============================================================================

function verFicha() {
  try {
    const ss = SpreadsheetApp.getActive();
    const rango = ss.getActiveRange();
    const hoja = rango.getSheet();

    if (hoja.getName() !== CONFIG.HOJA) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila en Maestro');
      return;
    }
    if (rango.getRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila de datos');
      return;
    }

    const datos = hoja.getRange(rango.getRow(), 1, 1, 26).getValues()[0];
    if (!datos[0]) {
      SpreadsheetApp.getUi().alert('⚠️ Fila vacía');
      return;
    }

    mostrarFicha(datos);
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

function mostrarFicha(datos) {
  const C = CONFIG.COL;
  const nombre     = datos[C.NOMBRE - 1]     || '—';
  const dpi        = datos[C.DPI - 1]        || '—';
  const edad       = datos[C.EDAD - 1]       || '—';
  const genero     = datos[C.GENERO - 1]     || '—';
  const telefono   = datos[C.TELEFONO - 1]   || '—';
  const zona       = datos[C.ZONA - 1]       || '—';
  const perfil     = datos[C.PERFIL - 1]     || '—';
  const prioridad  = datos[C.PRIORIDAD - 1]  || '—';
  const puntaje    = datos[C.PUNTAJE - 1]    || 0;
  const estado     = datos[C.ESTADO - 1]     || 'Orientación';
  const objetivo   = datos[C.OBJETIVO - 1]   || '—';
  const fortalezas = datos[C.FORTALEZAS - 1] || '—';
  const docUrl     = datos[C.DOC_URL - 1]    || '';
  const d = [C.DIM1,C.DIM2,C.DIM3,C.DIM4,C.DIM5,C.DIM6].map(k => Number(datos[k-1]) || 0);

  const c    = CONFIG.COLORES[perfil] || { fondo: '#f5f5f5', texto: '#333' };
  const cP   = prioridad === 'CRÍTICO' ? '#c62828' : prioridad === 'ALTO' ? '#e65100' : prioridad === 'MEDIO' ? '#f9a825' : '#388e3c';
  const ini  = nombre.split(' ').slice(0,2).map(p => p[0]||'').join('').toUpperCase();
  const docBtn = docUrl ? '<a href="' + docUrl + '" target="_blank" style="display:block;text-align:center;background:#1f73e6;color:#fff;padding:8px;border-radius:4px;margin:10px 0 0;text-decoration:none;font-size:12px;font-weight:bold">📄 Abrir Expediente en Drive</a>' : '';

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    'body{font-family:Arial;font-size:13px;background:#f8f9fa;color:#333;overflow-y:auto}' +
    '.hdr{background:' + c.fondo + ';padding:16px;border-bottom:3px solid ' + c.texto + ';display:flex;gap:12px;align-items:center}' +
    '.av{width:60px;height:60px;border-radius:50%;background:' + c.texto + ';color:#fff;font-size:20px;font-weight:bold;display:flex;align-items:center;justify-content:center}' +
    '.nom{font-size:15px;font-weight:bold;color:' + c.texto + '}' +
    '.bdg{display:flex;gap:6px;margin-top:6px;flex-wrap:wrap}' +
    '.b{padding:3px 10px;border-radius:12px;font-size:11px;font-weight:bold;border:1.5px solid ' + c.texto + ';color:' + c.texto + '}' +
    '.bp{background:' + cP + ';border-color:' + cP + ';color:#fff}' +
    '.be{background:#fff;color:#777}' +
    '.pts{text-align:center;padding:12px;background:#fff;border-bottom:1px solid #eee}' +
    '.pn{font-size:32px;font-weight:bold;color:' + c.texto + '}' +
    '.sec{padding:12px 14px;background:#fff;border-bottom:1px solid #eee}' +
    '.st{font-size:10px;font-weight:bold;color:#999;text-transform:uppercase;margin-bottom:8px}' +
    '.row{display:flex;justify-content:space-between;padding:4px 0}' +
    '.desc{color:#555;line-height:1.5}' +
    '</style></head><body>' +
    '<div class="hdr"><div class="av">' + ini + '</div>' +
    '<div><div class="nom">' + nombre + '</div>' +
    '<div class="bdg"><span class="b">' + perfil + '</span>' +
    '<span class="b bp">' + prioridad + '</span>' +
    '<span class="b be">' + estado + '</span></div></div></div>' +
    '<div class="pts"><div class="pn">' + puntaje + '<span style="font-size:14px;color:#aaa">/60</span></div></div>' +
    '<div class="sec"><div class="st">Datos Personales</div>' +
    '<div class="row"><span>DPI</span><span>' + dpi + '</span></div>' +
    '<div class="row"><span>Edad</span><span>' + edad + '</span></div>' +
    '<div class="row"><span>Género</span><span>' + genero + '</span></div>' +
    '<div class="row"><span>Teléfono</span><span>' + telefono + '</span></div>' +
    '<div class="row"><span>Zona</span><span>' + zona + '</span></div></div>' +
    '<div class="sec"><div class="st">Objetivo</div><div class="desc">' + objetivo + '</div></div>' +
    '<div class="sec"><div class="st">Fortalezas</div><div class="desc">' + fortalezas + '</div></div>' +
    '<div class="sec"><div class="st">Dimensiones</div>' +
    '<div style="display:flex;justify-content:center;padding:8px 0">' + svgRadar(d, c.texto) + '</div>' +
    svgBarras(d, c.texto) + docBtn + '</div></body></html>'
  );

  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================================
// ESTADÍSTICAS
// ============================================================================

function verEstadisticas() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Sin datos');
      return;
    }

    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 26).getValues().filter(r => r[0]);
    const total = datos.length;
    const pA = datos.filter(r => r[CONFIG.COL.PERFIL-1] === 'Perfil A').length;
    const pB = datos.filter(r => r[CONFIG.COL.PERFIL-1] === 'Perfil B').length;
    const pC = datos.filter(r => r[CONFIG.COL.PERFIL-1] === 'Perfil C').length;
    const pD = datos.filter(r => r[CONFIG.COL.PERFIL-1] === 'Perfil D').length;
    const criticos = datos.filter(r => r[CONFIG.COL.PRIORIDAD-1] === 'CRÍTICO').length;
    const puntajes = datos.map(r => Number(r[CONFIG.COL.PUNTAJE-1])).filter(n => n > 0);
    const prom = puntajes.length ? Math.round(puntajes.reduce((a,b) => a+b, 0) / puntajes.length) : 0;

    const html = HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
      'body{font-family:Arial;padding:16px;background:#f8f9fa;font-size:13px}' +
      'h2{color:#1f73e6;margin-bottom:14px;font-size:15px}' +
      '.card{background:#fff;border-radius:8px;padding:14px;margin-bottom:10px}' +
      '.big{font-size:40px;font-weight:bold;color:#1f73e6;text-align:center}' +
      '.pill{display:inline-block;padding:4px 10px;border-radius:12px;font-size:12px;font-weight:bold;margin:2px}' +
      '.pA{background:#d9ead3;color:#274e13}.pB{background:#cfe2f3;color:#1c4587}' +
      '.pC{background:#fff2cc;color:#7f6000}.pD{background:#f4cccc;color:#660000}' +
      '</style></head><body>' +
      '<h2>📊 Estadísticas</h2>' +
      '<div class="card"><div class="big">' + total + '</div><div style="text-align:center;color:#999;font-size:11px;margin-top:4px">Participantes</div></div>' +
      '<div class="card"><div class="big" style="font-size:20px;text-align:left">Por Perfil</div>' +
      '<span class="pill pA">A: ' + pA + '</span>' +
      '<span class="pill pB">B: ' + pB + '</span>' +
      '<span class="pill pC">C: ' + pC + '</span>' +
      '<span class="pill pD">D: ' + pD + '</span></div>' +
      '<div class="card"><div class="big" style="font-size:20px;text-align:left;margin-bottom:8px">Prioridad Alta</div>' +
      '<div style="color:#c62828;font-size:28px;font-weight:bold">🔴 ' + criticos + ' Críticos</div></div>' +
      '<div class="card"><div class="big" style="font-size:20px;text-align:left;margin-bottom:4px">Puntaje Promedio</div>' +
      '<div style="font-size:28px;font-weight:bold;color:#555">' + prom + ' <span style="font-size:14px;color:#aaa">/60</span></div></div>' +
      '</body></html>'
    );

    SpreadsheetApp.getUi().showSidebar(html);
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================================================
// PROBAR KOBO
// ============================================================================

function probarKobo() {
  try {
    const apiKey = PropertiesService.getUserProperties().getProperty('KOBO_API_KEY') || CONFIG.KOBO_API_KEY;
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/',
      { headers: { Authorization: 'Token ' + apiKey }, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() === 200) {
      const n = JSON.parse(resp.getContentText()).deployment__submission_count || 0;
      SpreadsheetApp.getUi().alert('✅ Conexión OK\n\n📊 ' + n + ' respuestas en Kobo');
    } else {
      SpreadsheetApp.getUi().alert('❌ Error ' + resp.getResponseCode());
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================================================
// CONFIGURACIÓN
// ============================================================================

function abrirConfiguracion() {
  const apiKey = PropertiesService.getUserProperties().getProperty('KOBO_API_KEY') || CONFIG.KOBO_API_KEY;
  const folderId = PropertiesService.getUserProperties().getProperty('FOLDER_ID') || CONFIG.FOLDER_PARTICIPANTES_ID;
  const email = PropertiesService.getUserProperties().getProperty('ADMIN_EMAIL') || CONFIG.ADMIN_EMAIL;

  const html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;padding:16px;background:#f5f5f5;font-size:13px}' +
    'label{display:block;font-weight:bold;margin:12px 0 4px;color:#333}' +
    'input,textarea{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;font-size:11px}' +
    'textarea{height:80px;resize:vertical}' +
    'button{margin-top:14px;background:#1f73e6;color:#fff;padding:10px;border:none;border-radius:4px;cursor:pointer;width:100%;font-weight:bold}' +
    '.info{background:#e8f5e9;color:#2e7d32;padding:10px;border-radius:4px;margin-bottom:12px;font-size:12px}' +
    '</style>' +
    '<div class="info">⚙️ Configuración del sistema</div>' +
    '<label>Email para alertas:</label>' +
    '<input id="email" value="' + email + '">' +
    '<label>ID Carpeta Drive:</label>' +
    '<input id="cid" value="' + folderId + '">' +
    '<label>API Key Kobo:</label>' +
    '<textarea id="key">' + apiKey + '</textarea>' +
    '<button onclick="var e=document.getElementById(\'email\').value.trim(),c=document.getElementById(\'cid\').value.trim(),k=document.getElementById(\'key\').value.trim();if(!e||!c||!k){alert(\'Completa todos\');return;}google.script.run.guardarConfig(e,c,k);alert(\'✅ Guardado\')">💾 Guardar</button>'
  );
  SpreadsheetApp.getUi().showSidebar(html);
}

function guardarConfig(email, carpeta, apiKey) {
  PropertiesService.getUserProperties()
    .setProperty('ADMIN_EMAIL', email)
    .setProperty('FOLDER_ID', carpeta)
    .setProperty('KOBO_API_KEY', apiKey);
}

// ============================================================================
// SVG
// ============================================================================

function svgRadar(dims, color) {
  const cx = 85, cy = 85, r = 65;
  const ang = [-90, -30, 30, 90, 150, 210];
  const lbs = ['Educ', 'Labor', 'Digit', 'Vocal', 'Barr', 'Apoyo'];

  let fnd = '';
  [0.33, 0.66, 1].forEach(n => {
    const pts = ang.map(a => {
      const rad = a * Math.PI / 180;
      return (cx + r*n*Math.cos(rad)).toFixed(1) + ',' + (cy + r*n*Math.sin(rad)).toFixed(1);
    }).join(' ');
    fnd += '<polygon points="' + pts + '" fill="none" stroke="#e0e0e0" stroke-width="1"/>';
  });

  const ejes = ang.map(a => {
    const rad = a * Math.PI / 180;
    return '<line x1="' + cx + '" y1="' + cy + '" x2="' + (cx+r*Math.cos(rad)).toFixed(1) + '" y2="' + (cy+r*Math.sin(rad)).toFixed(1) + '" stroke="#e0e0e0" stroke-width="1"/>';
  }).join('');

  const dpts = ang.map((a,i) => {
    const rad = a * Math.PI / 180, esc = dims[i]/10;
    return (cx+r*esc*Math.cos(rad)).toFixed(1) + ',' + (cy+r*esc*Math.sin(rad)).toFixed(1);
  }).join(' ');

  const lbls = ang.map((a,i) => {
    const rad = a * Math.PI / 180;
    return '<text x="' + (cx+(r+14)*Math.cos(rad)).toFixed(1) + '" y="' + (cy+(r+14)*Math.sin(rad)).toFixed(1) + '" text-anchor="middle" font-size="8" fill="#777">' + lbs[i] + '</text>';
  }).join('');

  return '<svg width="170" height="170" viewBox="0 0 170 170">' + fnd + ejes + '<polygon points="' + dpts + '" fill="' + color + '33" stroke="' + color + '" stroke-width="2"/>' + lbls + '</svg>';
}

function svgBarras(dims, color) {
  const lbs = ['Educativo','Laboral','Digital','Vocacional','Barreras','Red Apoyo'];
  return dims.map((v,i) => {
    const pct = Math.round((v/10)*100);
    return '<div style="margin:5px 0"><div style="display:flex;justify-content:space-between;font-size:11px;color:#666;margin-bottom:2px"><span>' + lbs[i] + '</span><span>' + v + '/10</span></div><div style="background:#eee;border-radius:3px;height:6px"><div style="width:' + pct + '%;background:' + color + ';height:6px"></div></div></div>';
  }).join('');
}

// ============================================================================
// DERIVACIONES AUTOMÁTICAS
// ============================================================================

function registrarDerivacionesAutomaticas(ss, id, nombre, fila, datosKobo) {
  try {
    const deriv = ss.getSheetByName('Derivaciones');
    if (!deriv) return;

    const perfil    = fila[CONFIG.COL.PERFIL - 1] || '';
    const prioridad = fila[CONFIG.COL.PRIORIDAD - 1] || '';
    const dimBarreras = Number(fila[CONFIG.COL.DIM5 - 1]) || 0;
    const dimDigital  = Number(fila[CONFIG.COL.DIM3 - 1]) || 0;
    const dimEducativo = Number(fila[CONFIG.COL.DIM1 - 1]) || 0;
    const dimApoyo    = Number(fila[CONFIG.COL.DIM6 - 1]) || 0;
    const fecha = new Date();

    const agregarDerivacion = (tipo, destino, motivo) => {
      deriv.appendRow([fecha, id, nombre, tipo, destino, motivo, 'Pendiente', '', '', '']);
      const ultima = deriv.getLastRow();
      const color = tipo === '🚨 URGENTE' ? '#ffcdd2' : '#fff9c4';
      deriv.getRange(ultima, 1, 1, 10).setBackground(color);
    };

    // PERFIL D → Derivaciones urgentes
    if (perfil === 'Perfil D' || prioridad === 'CRÍTICO') {
      agregarDerivacion('🚨 URGENTE', 'Creamos Voces (Apoyo Emocional)', 'Barreras críticas detectadas - Perfil ' + perfil);
      if (dimBarreras <= 4) {
        agregarDerivacion('🚨 URGENTE', 'Servicios Profesionales (Legal/Salud)', 'Dimensión barreras estructurales crítica: ' + dimBarreras + '/10');
      }
      enviarAlertaBarrerasCriticas(id, nombre, perfil, dimBarreras);
    }

    // PERFIL B → Mentoría vocacional
    if (perfil === 'Perfil B') {
      agregarDerivacion('💡 SUGERIDA', 'Mentoría Vocacional (Fito)', 'No reconoce su valor/habilidades - Requiere 3-6 sesiones');
    }

    // PERFIL C → Formación y educación
    if (perfil === 'Perfil C') {
      if (dimEducativo <= 4) {
        agregarDerivacion('💡 SUGERIDA', 'Educación de Adultos', 'Capital educativo bajo: ' + dimEducativo + '/10');
      }
      if (dimDigital <= 4) {
        agregarDerivacion('💡 SUGERIDA', 'Alfabetización Digital', 'Habilidades digitales bajas: ' + dimDigital + '/10');
      }
      agregarDerivacion('💡 SUGERIDA', 'Formación Técnica', 'Desarrollo de capital humano requerido');
    }

    // Barreras críticas sin importar perfil
    if (dimBarreras <= 3 && perfil !== 'Perfil D') {
      agregarDerivacion('⚠️ ALERTA', 'Creamos Voces (Apoyo Emocional)', 'Barreras estructurales muy bajas: ' + dimBarreras + '/10');
    }

    // Red de apoyo débil
    if (dimApoyo <= 3) {
      agregarDerivacion('💡 SUGERIDA', 'Grupos de Apoyo Comunitario', 'Red de apoyo débil: ' + dimApoyo + '/10');
    }

    // PERFIL A → Lista para empleo directo
    if (perfil === 'Perfil A') {
      agregarDerivacion('✅ OPORTUNIDAD', 'Intermediación Laboral (Bolsa de empleo)', 'Lista para empleo directo - ' + (Number(fila[CONFIG.COL.PUNTAJE - 1]) || 0) + '/60 pts');
    }
  } catch(e) {
    // Error silencioso
  }
}

function enviarAlertaBarrerasCriticas(id, nombre, perfil, dimBarreras) {
  try {
    const adminEmail = PropertiesService.getUserProperties().getProperty('ADMIN_EMAIL') || CONFIG.ADMIN_EMAIL;
    const asunto = '🚨 ALERTA URGENTE: Barreras críticas — ' + nombre;
    const body =
      '🚨 CASO URGENTE DETECTADO EN SINCRONIZACIÓN\n' +
      '═══════════════════════════════════════════\n\n' +
      'Participante: ' + nombre + '\n' +
      'ID: ' + id + '\n' +
      'Perfil: ' + perfil + '\n' +
      'Barreras estructurales: ' + dimBarreras + '/10\n\n' +
      'ACCIONES REQUERIDAS:\n' +
      '• Derivar URGENTE a Creamos Voces (Apoyo Emocional)\n' +
      '• Derivar a Servicios Profesionales (Legal/Salud)\n' +
      '• NO iniciar proceso de empleo hasta resolver barreras\n\n' +
      'Abre el Sheet para ver el expediente completo.\n\n' +
      'Sistema Paso a Paso — Alerta Automática';
    MailApp.sendEmail(adminEmail, asunto, body);
  } catch(e) {}
}

function verDerivaciones() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('Derivaciones');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('❌ Instala el sistema primero.');
      return;
    }
    ss.setActiveSheet(sheet);

    const pendientes = sheet.getLastRow() > 1
      ? sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues().filter(r => r[0] === 'Pendiente').length
      : 0;

    if (pendientes > 0) {
      SpreadsheetApp.getUi().alert('📋 Hoja Derivaciones abierta\n\n⚠️ Tienes ' + pendientes + ' derivación(es) PENDIENTE(S).\n\nCambia el Estado a "Completado" cuando las hayas gestionado.');
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================================================
// ANALYTICS
// ============================================================================

function abrirAnalytics() {
  try {
    const ss = SpreadsheetApp.getActive();
    actualizarAnalytics(ss);
    ss.setActiveSheet(ss.getSheetByName('Analytics'));
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

function actualizarAnalytics(ss) {
  try {
    if (!ss) ss = SpreadsheetApp.getActive();
    const hoja  = ss.getSheetByName(CONFIG.HOJA);
    const an    = ss.getSheetByName('Analytics');
    const deriv = ss.getSheetByName('Derivaciones');
    if (!hoja || !an) return;

    an.clear();
    const ahora = new Date();
    const tz = Session.getScriptTimeZone();

    const datos = hoja.getLastRow() > 1
      ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 26).getValues().filter(r => r[0])
      : [];
    const total = datos.length;

    // ─── Encabezado ───
    an.appendRow(['ANALYTICS — PASO A PASO']);
    an.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#6a1b9a');
    an.appendRow(['Actualizado: ' + Utilities.formatDate(ahora, tz, 'dd/MM/yyyy HH:mm') + '   |   Total: ' + total + ' participantes']);
    an.getRange(2, 1).setFontColor('#888');
    an.appendRow([]);

    const titulo = (txt, color) => {
      an.appendRow([txt]);
      an.getRange(an.getLastRow(), 1).setFontWeight('bold').setFontColor(color || '#6a1b9a').setFontSize(12);
    };
    const fila2 = (a, b, bg) => {
      an.appendRow([a, b]);
      if (bg) an.getRange(an.getLastRow(), 1, 1, 2).setBackground(bg);
    };

    // ─── Perfiles ───
    titulo('POR PERFIL');
    const perfilesCnt = { 'Perfil A': 0, 'Perfil B': 0, 'Perfil C': 0, 'Perfil D': 0, 'Sin perfil': 0 };
    datos.forEach(r => { const p = normalizarPerfil(r[CONFIG.COL.PERFIL-1]); perfilesCnt[p] = (perfilesCnt[p]||0)+1; });
    fila2('🟢 Perfil A — Listo para empleo', perfilesCnt['Perfil A'], '#d9ead3');
    fila2('🔵 Perfil B — Orientación vocacional', perfilesCnt['Perfil B'], '#cfe2f3');
    fila2('🟡 Perfil C — Desarrollo de capacidades', perfilesCnt['Perfil C'], '#fff2cc');
    fila2('🔴 Perfil D — Barreras críticas URGENTE', perfilesCnt['Perfil D'], '#f4cccc');
    fila2('⬜ Sin perfil asignado', perfilesCnt['Sin perfil'] || 0);
    an.appendRow([]);

    // ─── Prioridades ───
    titulo('POR PRIORIDAD');
    const prioCnt = { 'CRÍTICO': 0, 'ALTO': 0, 'MEDIO': 0, 'BAJO': 0 };
    datos.forEach(r => { const p = normalizarPrioridad(r[CONFIG.COL.PRIORIDAD-1]); prioCnt[p] = (prioCnt[p]||0)+1; });
    fila2('🔴 CRÍTICO', prioCnt['CRÍTICO'], '#ffcdd2');
    fila2('🟠 ALTO',    prioCnt['ALTO'],    '#ffe0b2');
    fila2('🟡 MEDIO',   prioCnt['MEDIO'],   '#fff9c4');
    fila2('🟢 BAJO',    prioCnt['BAJO'],    '#c8e6c9');
    an.appendRow([]);

    // ─── Promedios por dimensión ───
    titulo('PROMEDIO POR DIMENSIÓN (escala 0-10)');
    const dims = [
      [CONFIG.COL.DIM1, '📚 Capital Educativo'],
      [CONFIG.COL.DIM2, '💼 Capital Laboral'],
      [CONFIG.COL.DIM3, '💻 Habilidades Digitales'],
      [CONFIG.COL.DIM4, '🎯 Claridad Vocacional'],
      [CONFIG.COL.DIM5, '🚧 Barreras Estructurales'],
      [CONFIG.COL.DIM6, '🤝 Red de Apoyo']
    ];
    dims.forEach(([col, nombre]) => {
      const vals = datos.map(r => Number(r[col-1])).filter(v => v > 0);
      const prom = vals.length ? (vals.reduce((a,b) => a+b, 0) / vals.length).toFixed(1) : '—';
      const bg = prom !== '—' && prom < 4 ? '#ffcdd2' : prom < 7 ? '#fff9c4' : '#c8e6c9';
      fila2(nombre, prom + ' / 10', prom !== '—' ? bg : null);
    });
    an.appendRow([]);

    // ─── Por zona ───
    titulo('POR ZONA / LUGAR DE RESIDENCIA');
    const zonas = {};
    datos.forEach(r => { const z = r[CONFIG.COL.ZONA-1] || 'Sin zona'; zonas[z] = (zonas[z]||0)+1; });
    Object.entries(zonas).sort((a,b) => b[1]-a[1]).slice(0, 10).forEach(([z, n]) => fila2(z, n));
    an.appendRow([]);

    // ─── Derivaciones ───
    titulo('DERIVACIONES');
    if (deriv && deriv.getLastRow() > 1) {
      const derivDatos = deriv.getRange(2, 1, deriv.getLastRow()-1, 10).getValues().filter(r => r[0]);
      const pendientes = derivDatos.filter(r => r[6] === 'Pendiente').length;
      const completadas = derivDatos.filter(r => r[6] === 'Completado').length;
      const urgentes = derivDatos.filter(r => r[3] === '🚨 URGENTE').length;
      fila2('Total derivaciones',  derivDatos.length);
      fila2('🚨 URGENTES',         urgentes, urgentes > 0 ? '#ffcdd2' : null);
      fila2('⏳ Pendientes',        pendientes, pendientes > 0 ? '#fff9c4' : null);
      fila2('✅ Completadas',       completadas, '#c8e6c9');
    } else {
      an.appendRow(['Sin derivaciones registradas aún']);
    }
    an.appendRow([]);

    // ─── Puntajes ───
    titulo('DIAGNÓSTICO — PUNTAJE TOTAL');
    const pts = datos.map(r => Number(r[CONFIG.COL.PUNTAJE-1])).filter(v => v > 0);
    const prom = pts.length ? Math.round(pts.reduce((a,b) => a+b,0) / pts.length) : 0;
    fila2('Promedio', prom + ' / 60');
    fila2('Máximo',   pts.length ? Math.max(...pts) + ' / 60' : '—');
    fila2('Mínimo',   pts.length ? Math.min(...pts) + ' / 60' : '—');
    const bajo30 = pts.filter(p => p < 30).length;
    if (bajo30 > 0) fila2('⚠️ Con puntaje < 30 (necesitan más apoyo)', bajo30, '#ffcdd2');

    // Formato
    an.setColumnWidth(1, 280);
    an.setColumnWidth(2, 120);
  } catch(e) {}
}

// ============================================================================
// NORMALIZACIÓN DE VALORES KOBO
// ============================================================================

function normalizarPerfil(valor) {
  if (!valor) return 'Sin perfil';
  const v = valor.toString().toLowerCase();
  if (v.includes('perfil a') || v.startsWith('a ') || v === 'a') return 'Perfil A';
  if (v.includes('perfil b') || v.startsWith('b ') || v === 'b') return 'Perfil B';
  if (v.includes('perfil c') || v.startsWith('c ') || v === 'c') return 'Perfil C';
  if (v.includes('perfil d') || v.startsWith('d ') || v === 'd') return 'Perfil D';
  return valor;
}

function normalizarPrioridad(valor) {
  if (!valor) return 'Sin prioridad';
  const v = valor.toString().toLowerCase();
  if (v.includes('crít') || v.includes('urgente') || v.includes('crítico')) return 'CRÍTICO';
  if (v.includes('alto') || v.includes('alta') || v.includes('high')) return 'ALTO';
  if (v.includes('medio') || v.includes('media') || v.includes('normal') || v.includes('medium')) return 'MEDIO';
  if (v.includes('bajo') || v.includes('baja') || v.includes('low')) return 'BAJO';
  return valor;
}

// ============================================================================
// DESINSTALAR & LIMPIAR
// ============================================================================

function desinstalarYLimpiar() {
  const ui = SpreadsheetApp.getUi();
  const resp1 = ui.alert(
    '⚠️ DESINSTALAR SISTEMA',
    '¿Confirmas que quieres eliminar?\n\n' +
    '• Todas las carpetas y documentos de participantes en Drive\n' +
    '• Hojas: Derivaciones, Log, Dashboard\n' +
    '• Todos los triggers automáticos\n\n' +
    'La hoja Maestro se conserva con sus datos.',
    ui.ButtonSet.YES_NO
  );
  if (resp1 !== ui.Button.YES) return;

  const resp2 = ui.alert('⚠️ CONFIRMAR', '¿Estás 100% seguro? Esta acción NO se puede deshacer.', ui.ButtonSet.YES_NO);
  if (resp2 !== ui.Button.YES) return;

  try {
    let eliminados = 0;

    // Eliminar carpetas de participantes en Drive
    try {
      const folderId = PropertiesService.getUserProperties().getProperty('FOLDER_ID') || CONFIG.FOLDER_PARTICIPANTES_ID;
      const folderBase = DriveApp.getFolderById(folderId);
      const subcarpetas = folderBase.getFolders();
      while (subcarpetas.hasNext()) {
        subcarpetas.next().setTrashed(true);
        eliminados++;
      }
    } catch(e) {}

    // Eliminar hojas secundarias
    const ss = SpreadsheetApp.getActive();
    ['Derivaciones', 'Log', 'Dashboard'].forEach(nombre => {
      const h = ss.getSheetByName(nombre);
      if (h) ss.deleteSheet(h);
    });

    // Eliminar todos los triggers
    ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

    // Limpiar PropertiesService
    PropertiesService.getUserProperties().deleteAllProperties();

    ui.alert('✅ Sistema limpiado\n\n' +
      '• ' + eliminados + ' carpetas eliminadas de Drive\n' +
      '• Hojas secundarias eliminadas\n' +
      '• Triggers eliminados\n\n' +
      'Puedes volver a instalar con 📥 Instalar Sistema.'
    );
  } catch(e) {
    ui.alert('❌ Error al desinstalar: ' + e);
  }
}

// ============================================================================
// RESTAURAR DESDE DRIVE
// ============================================================================

function restaurarDesdeDrive() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) {
      ui.alert('❌ No existe la hoja Maestro. Instala primero.');
      return;
    }

    // IDs existentes en el Sheet
    const existentes = new Set();
    if (hoja.getLastRow() > 1) {
      hoja.getRange(2, CONFIG.COL.ID, hoja.getLastRow() - 1, 1)
        .getValues().forEach(r => { if (r[0]) existentes.add(String(r[0])); });
    }

    const folderId = PropertiesService.getUserProperties().getProperty('FOLDER_ID') || CONFIG.FOLDER_PARTICIPANTES_ID;
    const folderBase = DriveApp.getFolderById(folderId);
    const subcarpetas = folderBase.getFolders();
    let restaurados = 0;

    while (subcarpetas.hasNext()) {
      const carpeta = subcarpetas.next();
      const nombreCarpeta = carpeta.getName(); // Formato: "ID - Nombre"
      const partes = nombreCarpeta.split(' - ');
      if (partes.length < 2) continue;

      const id = partes[0].trim();
      const nombre = partes.slice(1).join(' - ').trim();

      if (existentes.has(id)) continue; // Ya existe en Sheet

      // Buscar el doc de perfil dentro de la carpeta
      const archivos = carpeta.getFilesByName('Perfil - ' + nombre);
      const docId = archivos.hasNext() ? archivos.next().getId() : '';
      const docUrl = docId ? 'https://docs.google.com/document/d/' + docId + '/edit' : '';

      // Crear fila mínima con lo que podemos recuperar
      const fila = [
        id, new Date(), nombre,
        '','','','','','','','','','',
        '','','0','0','0','0','0','0','0',
        'Recuperado',
        carpeta.getId(), docId, docUrl
      ];

      hoja.appendRow(fila);
      restaurados++;
    }

    ui.alert('✅ Restauración completada\n\n' +
      restaurados + ' participantes recuperados desde Drive\n\n' +
      'Nota: Solo se recuperan ID, Nombre y links.\n' +
      'Sincroniza con Kobo para completar los datos.'
    );
  } catch(e) {
    ui.alert('❌ Error al restaurar: ' + e);
  }
}

// ============================================================================
// GOOGLE CALENDAR — Eventos automáticos por cambio de estado
// ============================================================================

function crearEventoCalendario(nombre, nuevoEstado, docUrl) {
  try {
    const cal = CalendarApp.getDefaultCalendar();
    const hoy = new Date();
    let diasAdelante = 0;
    let titulo = '';
    let descripcion = '';

    if (nuevoEstado === 'Mentoría') {
      diasAdelante = 7;
      titulo = '📋 Seguimiento: ' + nombre;
      descripcion = 'Primera sesión de seguimiento en mentoría.\n\nParticipante: ' + nombre + '\nEstado: Mentoría\nExpediente: ' + (docUrl || 'Sin link');
    } else if (nuevoEstado === 'Formación') {
      diasAdelante = 14;
      titulo = '📚 Revisión formación: ' + nombre;
      descripcion = 'Revisión de avance en formación técnica.\n\nParticipante: ' + nombre + '\nExpediente: ' + (docUrl || 'Sin link');
    } else if (nuevoEstado === 'Completado' || nuevoEstado === 'Cierre') {
      diasAdelante = 30;
      titulo = '🏁 Revisión de resultados: ' + nombre;
      descripcion = 'Revisión de resultados post-cierre del caso.\n\nParticipante: ' + nombre + '\nExpediente: ' + (docUrl || 'Sin link');
    }

    if (!titulo) return;

    const fechaEvento = new Date(hoy.getTime() + diasAdelante * 24 * 60 * 60 * 1000);
    fechaEvento.setHours(9, 0, 0, 0);
    const fechaFin = new Date(fechaEvento.getTime() + 30 * 60 * 1000);

    cal.createEvent(titulo, fechaEvento, fechaFin, { description: descripcion });
  } catch(e) {
    // Sin permisos de calendar o error silencioso
  }
}

// ============================================================================
// EMAIL AL PARTICIPANTE — Notificación de avance de estado
// ============================================================================

function notificarParticipante(nombre, emailParticipante, nuevoEstado, datos) {
  try {
    const C = CONFIG.COL;
    const estados = ['Mentoría', 'Formación', 'Completado'];
    if (!estados.includes(nuevoEstado)) return;

    const nombreCorto = nombre.split(' ')[0];
    const responsable = datos[14] || 'Tu equipo Creamos'; // columna 15 = Responsable si existe
    const docUrl = datos[C.DOC_URL - 1] || '';

    let cuerpo = '';
    let asunto = '';

    if (nuevoEstado === 'Mentoría') {
      asunto = '✅ ¡Tu proceso avanzó a Mentoría! — Paso a Paso';
      cuerpo =
        'Hola ' + nombreCorto + ',\n\n' +
        'Tienes una actualización en tu proceso con Creamos. 🌱\n\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        '✅ Tu estado ha avanzado a: MENTORÍA\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
        'En esta etapa vas a recibir acompañamiento personalizado para\n' +
        'definir tu objetivo laboral y fortalecer tus habilidades.\n\n' +
        'Tu equipo de acompañamiento está contigo. 💪\n\n' +
        'Si tienes preguntas, responde a este correo.\n\n' +
        '¡Seguimos adelante! 🚀\n' +
        'Equipo Creamos — Programa Paso a Paso';
    } else if (nuevoEstado === 'Formación') {
      asunto = '📚 ¡Iniciaste tu proceso de Formación! — Paso a Paso';
      cuerpo =
        'Hola ' + nombreCorto + ',\n\n' +
        '¡Excelentes noticias! 🎉\n\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        '📚 Tu estado ha avanzado a: FORMACIÓN\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
        'Estás iniciando tu proceso de desarrollo de habilidades técnicas.\n' +
        'Este es un paso muy importante en tu camino al empleo. 🌟\n\n' +
        '¡Seguimos contigo!\n' +
        'Equipo Creamos — Programa Paso a Paso';
    } else if (nuevoEstado === 'Completado') {
      asunto = '🏆 ¡Completaste el programa Paso a Paso! — Creamos';
      cuerpo =
        'Hola ' + nombreCorto + ',\n\n' +
        '¡FELICITACIONES! 🎊🎉\n\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        '🏆 Has COMPLETADO el programa Paso a Paso\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
        'Has recorrido un camino increíble y estamos muy orgullosos\n' +
        'de tu esfuerzo y dedicación. 💪\n\n' +
        'Recuerda que siempre puedes contar con Creamos.\n\n' +
        '¡Mucho éxito en tu nueva etapa! 🌟\n' +
        'Equipo Creamos — Programa Paso a Paso';
    }

    if (cuerpo) {
      MailApp.sendEmail({
        to: emailParticipante,
        subject: asunto,
        body: cuerpo,
        replyTo: CONFIG.ADMIN_EMAIL
      });
    }
  } catch(e) {
    // Sin email o error silencioso
  }
}

// ============================================================================
// FIN - v8.0
// ============================================================================
