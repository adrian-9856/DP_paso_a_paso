// ============================================================================
// SISTEMA PASO A PASO - FITO v6.0
// Automatizado + Fácil + Sin errores
// ============================================================================

const CONFIG = {
  SPREADSHEET_ID:          SpreadsheetApp.getActive().getId(),
  FOLDER_PARTICIPANTES_ID: "1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU",
  KOBO_API_KEY:            "64cc018b88067397addd36b09288be8b6539cf39",
  KOBO_ASSET_ID:           "abHRWdRnPhKwzPQBajc7RZ",
  KOBO_URL:                "https://kf.kobotoolbox.org/api/v2",
  HOJA:                    "Maestro",
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
      .addItem('📊 Estadísticas', 'verEstadisticas')
      .addItem('🧪 Probar Kobo', 'probarKobo')
      .addItem('⚙️ Configuración', 'abrirConfiguracion')
      .addToUi();
  } catch(e) {
    // Sin contexto UI — disparado desde trigger o API
  }
}

// ============================================================================
// INSTALAR
// ============================================================================

function instalar() {
  try {
    const ss = SpreadsheetApp.getActive();
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

    configurarTrigger();

    SpreadsheetApp.getUi().alert(
      '✅ Sistema instalado correctamente\n\n' +
      'Próximos pasos:\n' +
      '1. 🧪 Probar Kobo — verifica conexión\n' +
      '2. 🔄 Sincronizar Kobo — importa participantes\n' +
      '3. 📁 Se crean carpetas automáticamente en Drive\n' +
      '4. 📋 Ver Ficha — click en cualquier fila'
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error al instalar: ' + e);
  }
}

function configurarTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sincronizarAutomatico') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sincronizarAutomatico').timeBased().atHour(8).everyDays(1).create();
}

function sincronizarAutomatico() {
  sincronizar(true);
}

// ============================================================================
// SINCRONIZAR KOBO
// ============================================================================

function sincronizar(silencioso) {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ Instala primero: 📥 Instalar Sistema');
      return;
    }

    const apiKey = PropertiesService.getUserProperties().getProperty('KOBO_API_KEY') || CONFIG.KOBO_API_KEY;
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json',
      { headers: { Authorization: 'Token ' + apiKey }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ Error Kobo ' + resp.getResponseCode() + '\n\nVerifica API Key en ⚙️ Configuración');
      return;
    }

    const registros = JSON.parse(resp.getContentText()).results || [];
    if (!registros.length) {
      if (!silencioso) SpreadsheetApp.getUi().alert('⚠️ Sin respuestas en Kobo todavía');
      return;
    }

    // IDs ya existentes
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
        id,
        new Date(),
        nombre,
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

    if (!silencioso) {
      SpreadsheetApp.getUi().alert(
        '✅ Sincronización completada\n\n' +
        '👤 ' + agregados + ' participantes nuevos\n' +
        '📊 ' + registros.length + ' total en Kobo\n' +
        '📁 Carpetas y documentos creados en Drive'
      );
    }
  } catch(e) {
    if (!silencioso) SpreadsheetApp.getUi().alert('❌ Error en sincronización: ' + e);
  }
}

// ============================================================================
// CREAR EXPEDIENTE EN DRIVE (CARPETA + GOOGLE DOC)
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

    body.appendParagraph(
      'ID: ' + id + '   |   Perfil: ' + (datos['perfil_asignado'] || '—') +
      '   |   Prioridad: ' + (datos['prioridad_caso'] || '—') +
      '   |   Puntaje: ' + (datos['puntaje_total_60'] || 0) + '/60'
    );
    body.appendParagraph('');

    const s1 = body.appendParagraph('DATOS PERSONALES');
    s1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Nombre:    ' + nombre);
    body.appendParagraph('DPI:       ' + (datos['numero_de_dpi_opcional'] || '—'));
    body.appendParagraph('Edad:      ' + (datos['edad'] || '—') + ' años');
    body.appendParagraph('Género:    ' + (datos['genero'] || '—'));
    body.appendParagraph('Zona:      ' + (datos['lugar_de_residencia'] || '—'));
    body.appendParagraph('Teléfono:  ' + (datos['numero_de_telefono'] || '—'));
    body.appendParagraph('Email:     ' + (datos['correo_electronico_opcional'] || '—'));
    body.appendParagraph('');

    const s2 = body.appendParagraph('PERFIL PROFESIONAL');
    s2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Educación:         ' + (datos['cual_es_el_ultimo_grado_que_completaste'] || '—'));
    body.appendParagraph('Situación laboral: ' + (datos['cual_es_tu_situacion_laboral_actual'] || '—'));
    body.appendParagraph('Fortalezas:        ' + (datos['que_sabes_hacer_bien'] || '—'));
    body.appendParagraph('Objetivo laboral:  ' + (datos['que_tipo_de_empleo_estas_buscando_especificamente'] || '—'));
    body.appendParagraph('');

    const s3 = body.appendParagraph('DIAGNÓSTICO');
    s3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Dim. Educativo:   ' + (datos['dimension_1_capital_educativo'] || 0) + '/10');
    body.appendParagraph('Dim. Laboral:     ' + (datos['dimension_2_capital_laboral'] || 0) + '/10');
    body.appendParagraph('Dim. Digital:     ' + (datos['dimension_3_habilidades_digitales'] || 0) + '/10');
    body.appendParagraph('Dim. Vocacional:  ' + (datos['dimension_4_claridad_vocacional'] || 0) + '/10');
    body.appendParagraph('Dim. Barreras:    ' + (datos['dimension_5_barreras_estructurales'] || 0) + '/10');
    body.appendParagraph('Dim. Red Apoyo:   ' + (datos['dimension_6_red_apoyo'] || 0) + '/10');
    body.appendParagraph('');

    const s4 = body.appendParagraph('DISCUSIÓN DE CASO');
    s4.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('[Agregar notas de cada sesión aquí]');
    body.appendParagraph('');

    const s5 = body.appendParagraph('PLAN DE ACCIÓN');
    s5.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('[Definir acciones, metas y responsables]');

    doc.saveAndClose();

    return {
      carpetaId: carpeta.getId(),
      docId:     doc.getId(),
      docUrl:    'https://docs.google.com/document/d/' + doc.getId() + '/edit'
    };
  } catch(e) {
    return { carpetaId: '', docId: '', docUrl: '' };
  }
}

// ============================================================================
// COLOREAR TABLA
// ============================================================================

function colorearTabla() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Sin datos para colorear');
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

// ============================================================================
// VER FICHA
// ============================================================================

function verFicha() {
  try {
    const ss    = SpreadsheetApp.getActive();
    const rango = ss.getActiveRange();
    const hoja  = rango.getSheet();

    if (hoja.getName() !== CONFIG.HOJA) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila en la hoja Maestro');
      return;
    }
    if (rango.getRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila de datos (no el encabezado)');
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
  const nombre     = datos[C.NOMBRE - 1]    || '—';
  const dpi        = datos[C.DPI - 1]       || '—';
  const edad       = datos[C.EDAD - 1]      || '—';
  const genero     = datos[C.GENERO - 1]    || '—';
  const telefono   = datos[C.TELEFONO - 1]  || '—';
  const zona       = datos[C.ZONA - 1]      || '—';
  const perfil     = datos[C.PERFIL - 1]    || '—';
  const prioridad  = datos[C.PRIORIDAD - 1] || '—';
  const puntaje    = datos[C.PUNTAJE - 1]   || 0;
  const estado     = datos[C.ESTADO - 1]    || 'Orientación';
  const objetivo   = datos[C.OBJETIVO - 1]  || '—';
  const fortalezas = datos[C.FORTALEZAS - 1]|| '—';
  const docUrl     = datos[C.DOC_URL - 1]   || '';
  const d = [C.DIM1,C.DIM2,C.DIM3,C.DIM4,C.DIM5,C.DIM6].map(k => Number(datos[k-1]) || 0);

  const c    = CONFIG.COLORES[perfil] || { fondo: '#f5f5f5', texto: '#333' };
  const cP   = prioridad === 'CRÍTICO' ? '#c62828' : prioridad === 'ALTO' ? '#e65100' : prioridad === 'MEDIO' ? '#f9a825' : '#388e3c';
  const ini  = nombre.split(' ').slice(0,2).map(p => p[0]||'').join('').toUpperCase();
  const docBtn = docUrl
    ? '<a href="' + docUrl + '" target="_blank" style="display:block;text-align:center;background:#1f73e6;color:#fff;padding:8px;border-radius:4px;margin:10px 0 0;text-decoration:none;font-size:12px;font-weight:bold">📄 Abrir Expediente en Drive</a>'
    : '';

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    'body{font-family:Arial;font-size:13px;background:#f8f9fa;color:#333;overflow-y:auto}' +
    '.hdr{background:' + c.fondo + ';padding:16px;border-bottom:3px solid ' + c.texto + ';display:flex;gap:12px;align-items:center}' +
    '.av{width:60px;height:60px;border-radius:50%;background:' + c.texto + ';color:#fff;font-size:20px;font-weight:bold;display:flex;align-items:center;justify-content:center;flex-shrink:0}' +
    '.nom{font-size:15px;font-weight:bold;color:' + c.texto + '}' +
    '.bdg{display:flex;gap:6px;margin-top:6px;flex-wrap:wrap}' +
    '.b{padding:3px 10px;border-radius:12px;font-size:11px;font-weight:bold;border:1.5px solid ' + c.texto + ';color:' + c.texto + '}' +
    '.bp{background:' + cP + ';border-color:' + cP + ';color:#fff}' +
    '.be{background:#fff;color:#777;border-color:#ddd}' +
    '.pts{text-align:center;padding:12px;background:#fff;border-bottom:1px solid #eee}' +
    '.pn{font-size:32px;font-weight:bold;color:' + c.texto + '}' +
    '.sec{padding:12px 14px;background:#fff;border-bottom:1px solid #eee}' +
    '.st{font-size:10px;font-weight:bold;color:#999;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px}' +
    '.row{display:flex;justify-content:space-between;padding:4px 0}' +
    '.rk{color:#777}.rv{font-weight:bold}' +
    '.desc{color:#555;line-height:1.5}' +
    '.gap{height:6px}' +
    '</style></head><body>' +
    '<div class="hdr"><div class="av">' + ini + '</div>' +
    '<div><div class="nom">' + nombre + '</div>' +
    '<div class="bdg">' +
    '<span class="b">' + perfil + '</span>' +
    '<span class="b bp">' + prioridad + '</span>' +
    '<span class="b be">' + estado + '</span>' +
    '</div></div></div>' +
    '<div class="pts"><div class="pn">' + puntaje + '<span style="font-size:14px;color:#aaa">/60</span></div>' +
    '<div style="font-size:10px;color:#999;margin-top:2px">Puntaje de Diagnóstico</div></div>' +
    '<div class="gap"></div>' +
    '<div class="sec"><div class="st">Datos Personales</div>' +
    '<div class="row"><span class="rk">DPI</span><span class="rv">' + dpi + '</span></div>' +
    '<div class="row"><span class="rk">Edad</span><span class="rv">' + edad + '</span></div>' +
    '<div class="row"><span class="rk">Género</span><span class="rv">' + genero + '</span></div>' +
    '<div class="row"><span class="rk">Teléfono</span><span class="rv">' + telefono + '</span></div>' +
    '<div class="row"><span class="rk">Zona</span><span class="rv">' + zona + '</span></div>' +
    '</div>' +
    '<div class="sec"><div class="st">Objetivo Laboral</div><div class="desc">' + objetivo + '</div></div>' +
    '<div class="sec"><div class="st">Fortalezas</div><div class="desc">' + fortalezas + '</div></div>' +
    '<div class="sec"><div class="st">Perfil por Dimensiones</div>' +
    '<div style="display:flex;justify-content:center;padding:8px 0">' + svgRadar(d, c.texto) + '</div>' +
    svgBarras(d, c.texto) + docBtn + '</div>' +
    '</body></html>'
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
    const total  = datos.length;
    const pA = datos.filter(r => r[CONFIG.COL.PERFIL-1] === 'Perfil A').length;
    const pB = datos.filter(r => r[CONFIG.COL.PERFIL-1] === 'Perfil B').length;
    const pC = datos.filter(r => r[CONFIG.COL.PERFIL-1] === 'Perfil C').length;
    const pD = datos.filter(r => r[CONFIG.COL.PERFIL-1] === 'Perfil D').length;
    const criticos = datos.filter(r => r[CONFIG.COL.PRIORIDAD-1] === 'CRÍTICO').length;
    const altos    = datos.filter(r => r[CONFIG.COL.PRIORIDAD-1] === 'ALTO').length;
    const puntajes = datos.map(r => Number(r[CONFIG.COL.PUNTAJE-1])).filter(n => n > 0);
    const prom     = puntajes.length ? Math.round(puntajes.reduce((a,b) => a+b, 0) / puntajes.length) : 0;

    const porEstado = {};
    datos.forEach(r => {
      const e = r[CONFIG.COL.ESTADO-1] || 'Sin estado';
      porEstado[e] = (porEstado[e]||0) + 1;
    });

    const estadoRows = Object.entries(porEstado).map(([k,v]) =>
      '<div class="row"><span>' + k + '</span><strong>' + v + '</strong></div>'
    ).join('');

    const html = HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
      'body{font-family:Arial;padding:16px;background:#f8f9fa;font-size:13px}' +
      'h2{color:#333;margin-bottom:14px;font-size:15px}' +
      '.card{background:#fff;border-radius:8px;padding:14px;margin-bottom:10px;box-shadow:0 1px 3px rgba(0,0,0,.1)}' +
      '.big{font-size:40px;font-weight:bold;color:#1f73e6;text-align:center}' +
      '.lbl{text-align:center;color:#999;font-size:11px;margin-top:4px}' +
      '.st{font-size:10px;font-weight:bold;color:#999;text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px}' +
      '.row{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #f0f0f0}' +
      '.pill{display:inline-block;padding:4px 10px;border-radius:12px;font-size:12px;font-weight:bold;margin:2px}' +
      '.pA{background:#d9ead3;color:#274e13}.pB{background:#cfe2f3;color:#1c4587}' +
      '.pC{background:#fff2cc;color:#7f6000}.pD{background:#f4cccc;color:#660000}' +
      '</style></head><body>' +
      '<h2>📊 Estadísticas del Programa</h2>' +
      '<div class="card"><div class="big">' + total + '</div><div class="lbl">Participantes totales</div></div>' +
      '<div class="card"><div class="st">Por Perfil</div>' +
      '<span class="pill pA">A: ' + pA + '</span>' +
      '<span class="pill pB">B: ' + pB + '</span>' +
      '<span class="pill pC">C: ' + pC + '</span>' +
      '<span class="pill pD">D: ' + pD + '</span></div>' +
      '<div class="card"><div class="st">Por Estado</div>' + estadoRows + '</div>' +
      '<div class="card"><div class="st">Prioridad Alta</div>' +
      '<div class="row"><span>🔴 CRÍTICO</span><strong>' + criticos + '</strong></div>' +
      '<div class="row"><span>🟠 ALTO</span><strong>' + altos + '</strong></div></div>' +
      '<div class="card"><div class="st">Puntaje Promedio</div>' +
      '<div style="font-size:28px;font-weight:bold;color:#555;text-align:center">' + prom + '<span style="font-size:14px;color:#aaa">/60</span></div></div>' +
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
      SpreadsheetApp.getUi().alert('✅ Conexión exitosa\n\n📊 ' + n + ' respuestas en Kobo\n\nPuedes sincronizar ahora.');
    } else {
      SpreadsheetApp.getUi().alert('❌ Error ' + resp.getResponseCode() + '\n\nVerifica API Key en ⚙️ Configuración');
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================================================
// CONFIGURACIÓN
// ============================================================================

function abrirConfiguracion() {
  const apiKey   = PropertiesService.getUserProperties().getProperty('KOBO_API_KEY') || CONFIG.KOBO_API_KEY;
  const folderId = PropertiesService.getUserProperties().getProperty('FOLDER_ID')    || CONFIG.FOLDER_PARTICIPANTES_ID;

  const html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial;padding:16px;background:#f5f5f5;font-size:13px}' +
    'label{display:block;font-weight:bold;margin:12px 0 4px;color:#333}' +
    'input,textarea{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;font-family:monospace;font-size:11px;box-sizing:border-box}' +
    'textarea{height:80px;resize:vertical}' +
    'button{margin-top:14px;background:#1f73e6;color:#fff;padding:10px;border:none;border-radius:4px;cursor:pointer;width:100%;font-weight:bold}' +
    '.info{background:#e8f5e9;color:#2e7d32;padding:10px;border-radius:4px;margin-bottom:12px;font-size:12px;line-height:1.5}' +
    '</style>' +
    '<div class="info">⚙️ Configuración del sistema<br>Los cambios se aplican en la próxima sincronización.</div>' +
    '<label>ID Carpeta Drive (participantes):</label>' +
    '<input id="cid" value="' + folderId + '">' +
    '<label>API Key de Kobo:</label>' +
    '<textarea id="key">' + apiKey + '</textarea>' +
    '<button onclick="var c=document.getElementById(\'cid\').value.trim(),k=document.getElementById(\'key\').value.trim();if(!c||!k){alert(\'Completa todos los campos\');return;}google.script.run.withSuccessHandler(function(){alert(\'✅ Guardado correctamente\')}).guardarConfig(c,k)">💾 Guardar Configuración</button>'
  );
  SpreadsheetApp.getUi().showSidebar(html);
}

function guardarConfig(carpeta, apiKey) {
  PropertiesService.getUserProperties()
    .setProperty('FOLDER_ID', carpeta)
    .setProperty('KOBO_API_KEY', apiKey);
}

// ============================================================================
// SVG RADAR Y BARRAS
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

  return '<svg width="170" height="170" viewBox="0 0 170 170">' + fnd + ejes
    + '<polygon points="' + dpts + '" fill="' + color + '33" stroke="' + color + '" stroke-width="2"/>'
    + lbls + '</svg>';
}

function svgBarras(dims, color) {
  const lbs = ['Educativo','Laboral','Digital','Vocacional','Barreras','Red Apoyo'];
  return dims.map((v,i) => {
    const pct = Math.round((v/10)*100);
    return '<div style="margin:5px 0">'
      + '<div style="display:flex;justify-content:space-between;font-size:11px;color:#666;margin-bottom:2px"><span>' + lbs[i] + '</span><span>' + v + '/10</span></div>'
      + '<div style="background:#eee;border-radius:3px;height:6px;overflow:hidden"><div style="width:' + pct + '%;background:' + color + ';height:6px"></div></div></div>';
  }).join('');
}

// ============================================================================
// FIN - SISTEMA PASO A PASO v6.0
// ============================================================================
