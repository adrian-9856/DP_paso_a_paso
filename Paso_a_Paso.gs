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
      .addItem('➡️ Registrar Derivación', 'abrirFormDerivacion')
      .addSeparator()
      .addItem('📈 Estadísticas', 'verEstadisticas')
      .addItem('🧪 Probar Kobo', 'probarKobo')
      .addItem('⚙️ Configuración', 'abrirConfiguracion')
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
      deriv.appendRow(['Fecha','Participante_ID','Participante_Nombre','Organización_Destino','Motivo','Resultado','Fecha_Seguimiento','Notas']);
      deriv.getRange(1, 1, 1, 8).setBackground('#1f73e6').setFontColor('#ffffff').setFontWeight('bold');
      deriv.setFrozenRows(1);
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
  // Limpiar triggers anteriores
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'sincronizarAutomatico' || fn === 'enviarReporteSemanal' || fn === 'verificarCasosDormidos') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Sincronización diaria 8 AM
  ScriptApp.newTrigger('sincronizarAutomatico').timeBased().atHour(8).everyDays(1).create();

  // Email semanal lunes 9 AM
  ScriptApp.newTrigger('enviarReporteSemanal').timeBased().atHour(9).onWeeksMonday().create();

  // Verificar casos dormidos diariamente 10 AM
  ScriptApp.newTrigger('verificarCasosDormidos').timeBased().atHour(10).everyDays(1).create();
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
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
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
        r['perfil_asignado'] || '',
        r['prioridad_caso'] || '',
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
      agregados++;
    });

    actualizarDashboard();

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

    const s3 = body.appendParagraph('PLAN DE ACCIÓN');
    s3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('[Definir acciones para este participante]');

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

// ============================================================================
// HISTORIAL DE CAMBIOS (onEdit)
// ============================================================================

function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== CONFIG.HOJA) return;

    const range = e.range;
    const fila = range.getRow();
    const columna = range.getColumn();
    const valorNuevo = e.value || '';
    const valorAnterior = e.oldValue || '';

    if (fila < 2 || !valorNuevo || valorNuevo === valorAnterior) return;

    const logSheet = e.source.getSheetByName('Log');
    if (!logSheet) return;

    const ahora = new Date();
    const usuario = Session.getEffectiveUser().getEmail();

    logSheet.appendRow([
      ahora,
      CONFIG.HOJA,
      fila,
      columna,
      valorAnterior,
      valorNuevo,
      usuario
    ]);
  } catch(e) {
    // Fallo silencioso
  }
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

    // Por perfil
    dash.appendRow(['POR PERFIL']);
    dash.getRange(dash.getLastRow(), 1).setFontWeight('bold').setFontColor('#1f73e6');
    const perfiles = { 'Perfil A': 0, 'Perfil B': 0, 'Perfil C': 0, 'Perfil D': 0, 'Sin perfil': 0 };
    datos.forEach(r => {
      const p = r[CONFIG.COL.PERFIL - 1] || 'Sin perfil';
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

    // Por prioridad
    dash.appendRow(['POR PRIORIDAD']);
    dash.getRange(dash.getLastRow(), 1).setFontWeight('bold').setFontColor('#1f73e6');
    const prioridades = { 'CRÍTICO': 0, 'ALTO': 0, 'MEDIO': 0, 'BAJO': 0, 'Sin prioridad': 0 };
    datos.forEach(r => {
      const p = r[CONFIG.COL.PRIORIDAD - 1] || 'Sin prioridad';
      prioridades[p] = (prioridades[p] || 0) + 1;
    });
    Object.entries(prioridades).forEach(([k, v]) => {
      if (k === 'CRÍTICO') dash.getRange(dash.getLastRow() + 1, 1, 1, 2).setBackground('#ffcdd2');
      dash.appendRow([k, v]);
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
// FIN - v7.0
// ============================================================================
