// ============================================================================
// SISTEMA PASO A PASO — v8.7
// Gestión de participantes: DP_Empleabilidad + Kobo + Drive + Email + Calendar
// ============================================================================

const CONFIG = {
  SPREADSHEET_ID:          SpreadsheetApp.getActive().getId(),
  FOLDER_PARTICIPANTES_ID: "1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU",
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
  // Columnas 1-based del Maestro (26 columnas en total)
  COL: {
    ID:1, FECHA:2, NOMBRE:3, DPI:4, EDAD:5, GENERO:6, TELEFONO:7,
    ZONA:8, EMAIL:9, EDUCACION:10, LABORAL:11, FORTALEZAS:12, OBJETIVO:13,
    PERFIL:14, PRIORIDAD:15, PUNTAJE:16, DIM1:17, DIM2:18, DIM3:19,
    DIM4:20, DIM5:21, DIM6:22, ESTADO:23, CARPETA_ID:24, DOC_ID:25, DOC_URL:26
  }
};

// DP_Empleabilidad — fuente externa
// Hoja "Paso a paso": A=Fecha | B=ID | C=Nombre | D=Teléfono | E=Género | F=Edad
//                     G=Educación | H=DPI | I=Formación | J=Cohorte | K=Nota | L=Activo
const DP_EMPLEABILIDAD = {
  SPREADSHEET_ID: '1_596FX6yr8tX93UyIks4emSeE2_vxLJMDyw9Zncsnzs',
  HOJA: 'Paso a paso',
  NCOLS: 12,
  C: { FECHA:0, ID:1, NOMBRE:2, TELEFONO:3, GENERO:4, EDAD:5,
       EDUCACION:6, DPI:7, FORMACION:8, COHORTE:9, NOTA:10, ACTIVO:11 }
};

const VERSION_SISTEMA = 'v8.7-2026-05-19';

const FLUJO_ESTADOS = {
  'Orientación': ['Mentoría', 'Derivación', 'Inactivo'],
  'Mentoría':    ['Formación', 'Cierre', 'Derivación', 'Inactivo'],
  'Formación':   ['Cierre', 'Derivación', 'Mentoría', 'Inactivo'],
  'Derivación':  ['Mentoría', 'Formación', 'Cierre', 'Inactivo'],
  'Inactivo':    ['Orientación', 'Mentoría'],
  'Cierre':      [],
  'Completado':  []
};

// ============================================================================
// HELPERS
// ============================================================================

function getFolderId() {
  const p = PropertiesService.getUserProperties();
  let id = p.getProperty('FOLDER_ID');
  if (!id) { id = CONFIG.FOLDER_PARTICIPANTES_ID; p.setProperty('FOLDER_ID', id); }
  return id;
}

function getAdminEmail() {
  const p = PropertiesService.getUserProperties();
  let em = p.getProperty('ADMIN_EMAIL');
  if (!em) { em = CONFIG.ADMIN_EMAIL; p.setProperty('ADMIN_EMAIL', em); }
  return em;
}

function getKoboKey() {
  const v = PropertiesService.getUserProperties().getProperty('KOBO_API_KEY');
  if (!v) throw new Error('API Key de Kobo no configurada. Abre ⚙️ Configuración.');
  return v;
}

function logError(fn, err) {
  try {
    const log = SpreadsheetApp.getActive().getSheetByName('Log');
    if (log) log.appendRow([new Date(), '⚠️ ERROR', fn, '', '', String(err), Session.getEffectiveUser().getEmail()]);
  } catch(e) {}
}

function escaparHtml(str) {
  return String(str || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#039;');
}

function normalizarPerfil(v) {
  if (!v) return 'Sin perfil';
  const s = String(v).toLowerCase();
  if (s.includes('perfil a') || s === 'a') return 'Perfil A';
  if (s.includes('perfil b') || s === 'b') return 'Perfil B';
  if (s.includes('perfil c') || s === 'c') return 'Perfil C';
  if (s.includes('perfil d') || s === 'd') return 'Perfil D';
  return String(v);
}

function normalizarPrioridad(v) {
  if (!v) return 'Sin prioridad';
  const s = String(v).toLowerCase();
  if (s.includes('crít') || s.includes('urgente')) return 'CRÍTICO';
  if (s.includes('alto') || s.includes('alta'))    return 'ALTO';
  if (s.includes('medio') || s.includes('media'))  return 'MEDIO';
  if (s.includes('bajo') || s.includes('baja'))    return 'BAJO';
  return String(v);
}

function validarTransicionEstado(ant, nuevo) {
  if (!ant || ant === nuevo) return null;
  const perm = FLUJO_ESTADOS[ant];
  if (!perm) return null;
  if (!perm.includes(nuevo)) {
    return '⚠️ Transición no permitida: ' + ant + ' → ' + nuevo +
      '\n\nDesde "' + ant + '" puedes ir a:\n• ' +
      (perm.length ? perm.join('\n• ') : '(estado final, no permite cambios)');
  }
  return null;
}

// ============================================================================
// MENÚ
// ============================================================================

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📊 PASO A PASO')
      .addItem('🔍 Diagnóstico del Sistema', 'diagnostico')
      .addItem('📥 Instalar Sistema', 'instalar')
      .addSeparator()
      .addItem('➕ Agregar Participante', 'agregarParticipanteManual')
      .addItem('✏️ Editar Participante', 'editarParticipante')
      .addSeparator()
      .addItem('🏢 Importar DP_Empleabilidad', 'sincronizarDesdeSheet')
      .addItem('📁 Crear Expedientes Drive', 'crearExpedientesPendientes')
      .addItem('🔄 Sincronizar Kobo', 'sincronizar')
      .addItem('🎨 Colorear Tabla', 'colorearTabla')
      .addSeparator()
      .addItem('📊 Dashboard', 'abrirDashboard')
      .addItem('📈 Analytics', 'abrirAnalytics')
      .addItem('📋 Ver Ficha del Participante', 'verFicha')
      .addItem('➡️ Derivar Participante', 'abrirFormDerivacion')
      .addItem('📊 Ver Derivaciones', 'verDerivaciones')
      .addSeparator()
      .addItem('🔁 Restaurar desde Drive', 'restaurarDesdeDrive')
      .addItem('🧪 Probar Kobo', 'probarKobo')
      .addItem('⚙️ Configuración', 'abrirConfiguracion')
      .addSeparator()
      .addItem('🗑️ Desinstalar & Limpiar', 'desinstalarYLimpiar')
      .addToUi();
  } catch(e) {}
}

// ============================================================================
// DIAGNÓSTICO
// ============================================================================

function diagnostico() {
  const ui = SpreadsheetApp.getUi();
  let r = '🔍 DIAGNÓSTICO DEL SISTEMA\n══════════════════════════════\n\n';
  r += '📌 VERSIÓN: ' + VERSION_SISTEMA + '\n\n';

  try {
    const ts = ScriptApp.getProjectTriggers();
    r += '⏱️ TRIGGERS (' + ts.length + '):\n';
    ts.forEach(t => r += '   • ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')\n');
    r += '\n';
  } catch(e) { r += '⏱️ Error: ' + e + '\n\n'; }

  try {
    const ss = SpreadsheetApp.getActive();
    r += '📋 HOJAS: ' + ss.getSheets().map(h => h.getName()).join(', ') + '\n';
    const m = ss.getSheetByName(CONFIG.HOJA);
    r += '   Maestro: ' + (m ? (m.getLastRow()-1) + ' participantes' : '❌ NO EXISTE') + '\n\n';
  } catch(e) { r += '📋 Error: ' + e + '\n\n'; }

  const props = PropertiesService.getUserProperties();
  r += '⚙️ CONFIGURACIÓN:\n';
  r += '   API Key Kobo: ' + (props.getProperty('KOBO_API_KEY') ? '✅ OK' : '❌ No configurada') + '\n';
  r += '   Folder Drive: ' + (props.getProperty('FOLDER_ID') || '(usando default)') + '\n\n';

  r += '🏢 DP_EMPLEABILIDAD:\n';
  try {
    const ext = SpreadsheetApp.openById(DP_EMPLEABILIDAD.SPREADSHEET_ID);
    const hf  = ext.getSheetByName(DP_EMPLEABILIDAD.HOJA);
    r += hf ? '   ✅ Accesible — ' + (hf.getLastRow()-1) + ' filas en "' + DP_EMPLEABILIDAD.HOJA + '"' :
              '   ⚠️ Hoja "' + DP_EMPLEABILIDAD.HOJA + '" no encontrada';
  } catch(e) { r += '   ❌ No accesible: ' + e; }

  ui.alert(r);
}

// ============================================================================
// INSTALAR
// ============================================================================

function instalar() {
  try {
    const ss = SpreadsheetApp.getActive();
    const HEADERS = [
      'ID_Creamos','Fecha_Registro','Nombre_Completo','DPI','Edad','Género',
      'Teléfono','Zona','Email','Nivel_Educativo','Situación_Laboral','Fortalezas',
      'Objetivo_Laboral','Perfil_Asignado','Prioridad','Puntaje_Total',
      'Dim_Educativo','Dim_Laboral','Dim_Digital','Dim_Vocacional',
      'Dim_Barreras','Dim_Apoyo','Estado','Carpeta_Drive_ID','Doc_Perfil_ID','Doc_Perfil_URL'
    ];

    // Maestro
    let maestro = ss.getSheetByName(CONFIG.HOJA);
    if (!maestro) maestro = ss.insertSheet(CONFIG.HOJA);
    if (maestro.getLastRow() === 0) {
      maestro.appendRow(HEADERS);
      maestro.getRange(1,1,1,26)
        .setBackground('#1a237e').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
      maestro.setFrozenRows(1);
      maestro.setColumnWidth(1,120).setColumnWidth(3,220).setColumnWidth(13,200)
             .setColumnWidth(24,160).setColumnWidth(25,160).setColumnWidth(26,280);
    }

    // Derivaciones
    if (!ss.getSheetByName('Derivaciones')) {
      const d = ss.insertSheet('Derivaciones');
      d.appendRow(['Fecha','ID_Participante','Nombre','Tipo','Destino','Motivo','Estado','Responsable','Fecha_Seguimiento','Notas']);
      d.getRange(1,1,1,10).setBackground('#bf360c').setFontColor('#fff').setFontWeight('bold');
      d.setFrozenRows(1);
      d.setColumnWidth(3,180); d.setColumnWidth(5,200); d.setColumnWidth(6,250);
    }

    // Log
    if (!ss.getSheetByName('Log')) {
      const l = ss.insertSheet('Log');
      l.appendRow(['Fecha_Hora','Hoja','Fila','Columna','Valor_Anterior','Valor_Nuevo','Usuario']);
      l.getRange(1,1,1,7).setBackground('#4a148c').setFontColor('#fff').setFontWeight('bold');
      l.setFrozenRows(1);
    }

    // Dashboard
    if (!ss.getSheetByName('Dashboard')) {
      const db = ss.insertSheet('Dashboard');
      db.appendRow(['DASHBOARD — PASO A PASO']);
      db.getRange(1,1).setFontSize(16).setFontWeight('bold').setFontColor('#1a237e');
    }

    // Analytics
    if (!ss.getSheetByName('Analytics')) {
      const an = ss.insertSheet('Analytics');
      an.appendRow(['ANALYTICS — PASO A PASO']);
      an.getRange(1,1).setFontSize(16).setFontWeight('bold').setFontColor('#6a1b9a');
      an.setFrozenRows(1);
    }

    configurarTriggers();

    SpreadsheetApp.getUi().alert(
      '✅ Sistema v8.7 instalado\n\n' +
      '📋 Hojas creadas: Maestro, Derivaciones, Log, Dashboard, Analytics\n' +
      '⏱️ Triggers configurados\n\n' +
      '⚙️ SIGUIENTE PASO:\n' +
      'Abre ⚙️ Configuración e ingresa tu email\n' +
      'y el ID de tu carpeta Drive.\n\n' +
      'Luego: 🏢 Importar DP_Empleabilidad'
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error al instalar: ' + e);
  }
}

function configurarTriggers() {
  const FNS = ['sincronizarAutomatico','enviarReporteSemanal','verificarCasosDormidos','manejarEdicion'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (FNS.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sincronizarAutomatico').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('enviarReporteSemanal').timeBased().atHour(9).onWeekDay(ScriptApp.WeekDay.MONDAY).create();
  ScriptApp.newTrigger('verificarCasosDormidos').timeBased().atHour(10).everyDays(1).create();
  ScriptApp.newTrigger('manejarEdicion').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}

function sincronizarAutomatico() {
  sincronizarDesdeSheet(true);
}

// ============================================================================
// IMPORTAR DESDE DP_EMPLEABILIDAD
// ============================================================================

function sincronizarDesdeSheet(silencioso) {
  try {
    const ss   = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ Instala el sistema primero (📥 Instalar Sistema).');
      return;
    }

    let ssExt;
    try { ssExt = SpreadsheetApp.openById(DP_EMPLEABILIDAD.SPREADSHEET_ID); }
    catch(e) {
      if (!silencioso) SpreadsheetApp.getUi().alert(
        '❌ No se puede abrir DP_Empleabilidad.\n\n' +
        'Verifica que tienes acceso al Google Sheet:\n' + DP_EMPLEABILIDAD.SPREADSHEET_ID
      );
      return;
    }

    const hojaSrc = ssExt.getSheetByName(DP_EMPLEABILIDAD.HOJA);
    if (!hojaSrc) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ No existe la hoja "' + DP_EMPLEABILIDAD.HOJA + '" en DP_Empleabilidad.');
      return;
    }
    if (hojaSrc.getLastRow() < 2) {
      if (!silencioso) SpreadsheetApp.getUi().alert('⚠️ La hoja "' + DP_EMPLEABILIDAD.HOJA + '" no tiene datos.');
      return;
    }

    // Leer toda la hoja fuente de una sola vez
    const datos = hojaSrc.getRange(2, 1, hojaSrc.getLastRow()-1, DP_EMPLEABILIDAD.NCOLS).getValues();

    // IDs ya existentes en Maestro → para evitar duplicados
    const existentes = new Set();
    if (hoja.getLastRow() > 1) {
      hoja.getRange(2, CONFIG.COL.ID, hoja.getLastRow()-1, 1)
        .getValues().forEach(r => { if (r[0]) existentes.add(String(r[0]).trim()); });
    }

    const C   = DP_EMPLEABILIDAD.C;
    const M   = CONFIG.COL;
    const hoy = new Date();
    const filasMaestro      = [];
    const filasDerivaciones = [];
    let omitidos = 0;

    datos.forEach(fila => {
      const nombre = String(fila[C.NOMBRE] || '').trim();
      if (!nombre) return; // fila vacía

      let id = String(fila[C.ID] || '').trim();
      if (!id) {
        // Generar ID automático si la fuente no tiene uno
        id = nombre.replace(/\s+/g,'').substring(0,4).toUpperCase() +
             Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'ddMMyyyy');
      }

      if (existentes.has(id)) { omitidos++; return; }
      existentes.add(id);

      const formacion = String(fila[C.FORMACION] || '').trim();
      const cohorte   = String(fila[C.COHORTE]   || '').trim();
      const nota      = String(fila[C.NOTA]       || '').trim();
      const activoStr = String(fila[C.ACTIVO]     || '').toLowerCase().trim();
      const activo    = ['true','si','sí','1','activo','yes'].includes(activoStr);
      const estado    = activo ? 'Orientación' : 'Inactivo';
      const fecha     = fila[C.FECHA] instanceof Date ? fila[C.FECHA] : hoy;

      // Construir fila completa (26 columnas) para el Maestro
      const row = new Array(26).fill('');
      row[M.ID         - 1] = id;
      row[M.FECHA      - 1] = fecha;
      row[M.NOMBRE     - 1] = nombre;
      row[M.DPI        - 1] = String(fila[C.DPI]       || '').trim();
      row[M.EDAD       - 1] = fila[C.EDAD] || '';
      row[M.GENERO     - 1] = String(fila[C.GENERO]    || '').trim();
      row[M.TELEFONO   - 1] = String(fila[C.TELEFONO]  || '').trim();
      row[M.EDUCACION  - 1] = String(fila[C.EDUCACION] || '').trim();
      row[M.LABORAL    - 1] = formacion;
      row[M.FORTALEZAS - 1] = cohorte ? 'Cohorte: ' + cohorte : '';
      row[M.OBJETIVO   - 1] = nota;
      row[M.ESTADO     - 1] = estado;
      // CARPETA_ID, DOC_ID, DOC_URL → se llenan con "📁 Crear Expedientes Drive"

      filasMaestro.push(row);

      if (formacion) {
        filasDerivaciones.push([
          hoy, id, nombre, '💡 SUGERIDA',
          'Formación Técnica — ' + formacion,
          'Importado de DP_Empleabilidad. Formación: ' + formacion + (cohorte ? ' | Cohorte: ' + cohorte : ''),
          'Pendiente', '', '', nota
        ]);
      }
    });

    // Escribir Maestro — una sola llamada de API (eficiente)
    if (filasMaestro.length > 0) {
      hoja.getRange(hoja.getLastRow()+1, 1, filasMaestro.length, 26).setValues(filasMaestro);
      SpreadsheetApp.flush();
    }

    // Escribir Derivaciones — una sola llamada de API
    if (filasDerivaciones.length > 0) {
      const deriv = ss.getSheetByName('Derivaciones');
      if (deriv) {
        const dr = deriv.getLastRow()+1;
        deriv.getRange(dr, 1, filasDerivaciones.length, 10).setValues(filasDerivaciones);
        deriv.getRange(dr, 1, filasDerivaciones.length, 10).setBackground('#e8f5e9');
        SpreadsheetApp.flush();
      }
    }

    if (!silencioso) {
      SpreadsheetApp.getUi().alert(
        '✅ Importación completada\n\n' +
        '👤 ' + filasMaestro.length + ' participante(s) nuevo(s) en Maestro\n' +
        '➡️ ' + filasDerivaciones.length + ' derivación(es) automática(s)\n' +
        '⏭️ ' + omitidos + ' ya existían (omitidos)\n\n' +
        (filasMaestro.length > 0
          ? '📁 SIGUIENTE PASO:\nEjecuta "📁 Crear Expedientes Drive"\npara generar carpetas en Drive.'
          : '✅ No hay participantes nuevos para importar.')
      );
    }
  } catch(e) {
    logError('sincronizarDesdeSheet', e);
    if (!silencioso) SpreadsheetApp.getUi().alert('❌ Error al importar: ' + e.message);
  }
}

// ============================================================================
// CREAR EXPEDIENTES EN DRIVE (de a 10 para no agotar los 6 minutos)
// ============================================================================

function crearExpedientesPendientes() {
  try {
    const ss   = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ No hay participantes en el Maestro todavía.');
      return;
    }

    const M          = CONFIG.COL;
    const datos      = hoja.getRange(2, 1, hoja.getLastRow()-1, 26).getValues();
    const folderBase = DriveApp.getFolderById(getFolderId());
    let creados      = 0;
    let limite       = 10;

    datos.forEach((fila, i) => {
      if (limite <= 0) return;
      if (fila[M.CARPETA_ID - 1]) return; // ya tiene expediente

      const id     = String(fila[M.ID     - 1] || '').trim();
      const nombre = String(fila[M.NOMBRE - 1] || '').trim();
      if (!id || !nombre) return;

      const exp = crearExpediente(folderBase, id, nombre, fila, M);
      if (exp.carpetaId) {
        hoja.getRange(i+2, M.CARPETA_ID, 1, 3).setValues([[exp.carpetaId, exp.docId, exp.docUrl]]);
        creados++;
        limite--;
      }
    });

    SpreadsheetApp.flush();
    const pendientes = datos.filter(f => !f[M.CARPETA_ID-1] && f[M.NOMBRE-1]).length - creados;

    SpreadsheetApp.getUi().alert(
      '✅ Expedientes creados: ' + creados + '\n\n' +
      (pendientes > 0
        ? '⏳ Faltan ' + pendientes + ' más.\nVuelve a ejecutar "📁 Crear Expedientes Drive".'
        : '🎉 ¡Todos los expedientes están listos!')
    );
  } catch(e) {
    logError('crearExpedientesPendientes', e);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

function crearExpediente(folderBase, id, nombre, fila, M) {
  try {
    const carpeta = folderBase.createFolder(id + ' — ' + nombre);
    const doc     = DocumentApp.create('Perfil — ' + nombre);
    DriveApp.getFileById(doc.getId()).moveTo(carpeta);
    const body = doc.getBody();
    body.clear();

    const h1 = body.appendParagraph('PERFIL DEL PARTICIPANTE');
    h1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    h1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('ID: ' + id);
    body.appendParagraph('');

    const s1 = body.appendParagraph('DATOS PERSONALES');
    s1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Nombre:            ' + nombre);
    body.appendParagraph('DPI:               ' + (fila[M.DPI-1]       || '—'));
    body.appendParagraph('Edad:              ' + (fila[M.EDAD-1]      || '—'));
    body.appendParagraph('Género:            ' + (fila[M.GENERO-1]    || '—'));
    body.appendParagraph('Teléfono:          ' + (fila[M.TELEFONO-1]  || '—'));
    body.appendParagraph('Zona:              ' + (fila[M.ZONA-1]      || '—'));
    body.appendParagraph('Email:             ' + (fila[M.EMAIL-1]     || '—'));
    body.appendParagraph('');

    const s2 = body.appendParagraph('PERFIL PROFESIONAL');
    s2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Educación:         ' + (fila[M.EDUCACION-1]  || '—'));
    body.appendParagraph('Situación laboral: ' + (fila[M.LABORAL-1]    || '—'));
    body.appendParagraph('Fortalezas:        ' + (fila[M.FORTALEZAS-1] || '—'));
    body.appendParagraph('Objetivo laboral:  ' + (fila[M.OBJETIVO-1]   || '—'));
    body.appendParagraph('');

    const s3 = body.appendParagraph('DIAGNÓSTICO');
    s3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Perfil:            ' + (fila[M.PERFIL-1]    || '—'));
    body.appendParagraph('Prioridad:         ' + (fila[M.PRIORIDAD-1] || '—'));
    body.appendParagraph('Puntaje total:     ' + (fila[M.PUNTAJE-1]   || 0) + '/60');
    body.appendParagraph('Educativo:         ' + (fila[M.DIM1-1] || 0) + '/10');
    body.appendParagraph('Laboral:           ' + (fila[M.DIM2-1] || 0) + '/10');
    body.appendParagraph('Digital:           ' + (fila[M.DIM3-1] || 0) + '/10');
    body.appendParagraph('Vocacional:        ' + (fila[M.DIM4-1] || 0) + '/10');
    body.appendParagraph('Barreras:          ' + (fila[M.DIM5-1] || 0) + '/10');
    body.appendParagraph('Red de apoyo:      ' + (fila[M.DIM6-1] || 0) + '/10');
    body.appendParagraph('');

    const s4 = body.appendParagraph('PLAN DE ACCIÓN');
    s4.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('[Definir acciones concretas para este participante]');
    body.appendParagraph('');

    const s5 = body.appendParagraph('HISTORIAL DE CAMBIOS');
    s5.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Fecha             | Campo            | Cambio                         | Usuario');
    body.appendParagraph('─────────────────────────────────────────────────────────────────────────');

    doc.saveAndClose();
    const docId = doc.getId();
    return {
      carpetaId: carpeta.getId(),
      docId:     docId,
      docUrl:    'https://docs.google.com/document/d/' + docId + '/edit'
    };
  } catch(e) {
    logError('crearExpediente', e);
    return { carpetaId:'', docId:'', docUrl:'' };
  }
}

// ============================================================================
// SINCRONIZAR KOBO (opcional — si se usa KoboToolbox)
// ============================================================================

function sincronizar(silencioso) {
  try {
    const ss   = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) { if (!silencioso) SpreadsheetApp.getUi().alert('❌ Instala primero.'); return; }

    let apiKey;
    try { apiKey = getKoboKey(); }
    catch(e) { if (!silencioso) SpreadsheetApp.getUi().alert('❌ ' + e.message); return; }

    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json',
      { headers: { Authorization: 'Token ' + apiKey }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ Error Kobo ' + resp.getResponseCode() + '\n\nVerifica tu API Key en ⚙️ Configuración.');
      return;
    }

    const registros = JSON.parse(resp.getContentText()).results || [];
    if (!registros.length) {
      if (!silencioso) SpreadsheetApp.getUi().alert('⚠️ No hay respuestas en Kobo todavía.');
      return;
    }

    const existentes = new Set();
    if (hoja.getLastRow() > 1) {
      hoja.getRange(2, CONFIG.COL.ID, hoja.getLastRow()-1, 1)
        .getValues().forEach(r => { if (r[0]) existentes.add(String(r[0])); });
    }

    const folderBase = DriveApp.getFolderById(getFolderId());
    const M = CONFIG.COL;
    let agregados = 0;

    registros.forEach(r => {
      const id = String(r['creamos_id'] || r['_id'] || '');
      if (!id || existentes.has(id)) return;

      const nombre = r['nombre_completo_del_la_participante'] || 'Participante ' + id;
      const fila   = new Array(26).fill('');
      fila[M.ID-1]         = id;
      fila[M.FECHA-1]      = new Date();
      fila[M.NOMBRE-1]     = nombre;
      fila[M.DPI-1]        = r['numero_de_dpi_opcional']                             || '';
      fila[M.EDAD-1]       = r['edad']                                               || '';
      fila[M.GENERO-1]     = r['genero']                                             || '';
      fila[M.TELEFONO-1]   = r['numero_de_telefono']                                 || '';
      fila[M.ZONA-1]       = r['lugar_de_residencia']                                || '';
      fila[M.EMAIL-1]      = r['correo_electronico_opcional']                        || '';
      fila[M.EDUCACION-1]  = r['cual_es_el_ultimo_grado_que_completaste']            || '';
      fila[M.LABORAL-1]    = r['cual_es_tu_situacion_laboral_actual']                || '';
      fila[M.FORTALEZAS-1] = r['que_sabes_hacer_bien']                               || '';
      fila[M.OBJETIVO-1]   = r['que_tipo_de_empleo_estas_buscando_especificamente']  || '';
      fila[M.PERFIL-1]     = normalizarPerfil(r['perfil_asignado']);
      fila[M.PRIORIDAD-1]  = normalizarPrioridad(r['prioridad_caso']);
      fila[M.PUNTAJE-1]    = Number(r['puntaje_total_60'])                           || 0;
      fila[M.DIM1-1]       = Number(r['dimension_1_capital_educativo'])              || 0;
      fila[M.DIM2-1]       = Number(r['dimension_2_capital_laboral'])                || 0;
      fila[M.DIM3-1]       = Number(r['dimension_3_habilidades_digitales'])          || 0;
      fila[M.DIM4-1]       = Number(r['dimension_4_claridad_vocacional'])            || 0;
      fila[M.DIM5-1]       = Number(r['dimension_5_barreras_estructurales'])         || 0;
      fila[M.DIM6-1]       = Number(r['dimension_6_red_apoyo'])                      || 0;
      fila[M.ESTADO-1]     = 'Orientación';

      const exp = crearExpediente(folderBase, id, nombre, fila, M);
      fila[M.CARPETA_ID-1] = exp.carpetaId;
      fila[M.DOC_ID-1]     = exp.docId;
      fila[M.DOC_URL-1]    = exp.docUrl;

      hoja.appendRow(fila);
      colorearFila(hoja, hoja.getLastRow(), fila[M.PERFIL-1]);
      registrarDerivacionesAutomaticas(ss, id, nombre, fila, M);
      existentes.add(id);
      agregados++;
    });

    if (!silencioso) {
      SpreadsheetApp.getUi().alert('✅ Kobo sincronizado\n👤 ' + agregados + ' nuevos\n📊 ' + registros.length + ' total en Kobo');
    }
  } catch(e) {
    logError('sincronizar', e);
    if (!silencioso) SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

// ============================================================================
// COLOREAR TABLA
// ============================================================================

function colorearTabla() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) { SpreadsheetApp.getUi().alert('⚠️ Sin datos.'); return; }
    const datos = hoja.getRange(2, 1, hoja.getLastRow()-1, CONFIG.COL.PERFIL).getValues();
    datos.forEach((r, i) => { if (r[0]) colorearFila(hoja, i+2, r[CONFIG.COL.PERFIL-1]); });
    SpreadsheetApp.getUi().alert('✅ Tabla coloreada\n🟢 A  🔵 B  🟡 C  🔴 D');
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e); }
}

function colorearFila(hoja, numFila, perfil) {
  const c = CONFIG.COLORES[perfil];
  const rango = hoja.getRange(numFila, 1, 1, 26);
  if (c) rango.setBackground(c.fondo).setFontColor(c.texto);
  else   rango.setBackground('#ffffff').setFontColor('#333333');
}

// ============================================================================
// TRIGGER DE EDICIÓN (instalable — tiene permisos para abrir Docs)
// ============================================================================

function onEdit(e) { return; } // Vacío — manejarEdicion() es el trigger activo

function manejarEdicion(e) {
  const cache = CacheService.getScriptCache();
  if (cache.get('revert_lock')) return;

  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== CONFIG.HOJA) return;

    const fila = e.range.getRow();
    const col  = e.range.getColumn();
    const valN = e.value    || '';
    const valA = e.oldValue || '';
    if (fila < 2 || !valN || valN === valA) return;

    // Validar transición de estado
    if (col === CONFIG.COL.ESTADO) {
      const err = validarTransicionEstado(valA, valN);
      if (err) {
        cache.put('revert_lock', '1', 15);
        sheet.getRange(fila, col).setValue(valA);
        cache.remove('revert_lock');
        SpreadsheetApp.getUi().alert(err);
        return;
      }
    }

    // Registrar en Log
    const logSheet = e.source.getSheetByName('Log');
    if (logSheet) {
      logSheet.appendRow([new Date(), CONFIG.HOJA, fila, col, valA, valN, Session.getEffectiveUser().getEmail()]);
    }

    // Sincronizar con Google Doc
    const C = CONFIG.COL;
    const syncCols = [C.NOMBRE,C.DPI,C.EDAD,C.GENERO,C.TELEFONO,C.ZONA,C.EMAIL,
                      C.EDUCACION,C.LABORAL,C.FORTALEZAS,C.OBJETIVO,C.PERFIL,C.PRIORIDAD,C.ESTADO];
    if (syncCols.includes(col)) {
      const datos = sheet.getRange(fila, 1, 1, 26).getValues()[0];
      const docId = datos[C.DOC_ID - 1];
      if (docId) sincronizarConDoc(docId, datos, col, valN, valA);

      if (col === C.ESTADO) {
        crearEventoCalendario(datos[C.NOMBRE-1], valN, datos[C.DOC_URL-1]);
        if (datos[C.EMAIL-1]) notificarParticipante(datos[C.NOMBRE-1], datos[C.EMAIL-1], valN);
      }
    }
  } catch(err) {
    cache.remove('revert_lock');
    logError('manejarEdicion', err);
  }
}

function sincronizarConDoc(docId, datos, col, valN, valA) {
  try {
    const doc  = DocumentApp.openById(docId);
    const body = doc.getBody();
    const C    = CONFIG.COL;

    const mapa = {};
    mapa[C.NOMBRE]     = { p:'Nombre:.*',            pre:'Nombre:            ' };
    mapa[C.DPI]        = { p:'DPI:.*',               pre:'DPI:               ' };
    mapa[C.EDAD]       = { p:'Edad:.*',              pre:'Edad:              ' };
    mapa[C.GENERO]     = { p:'Género:.*',            pre:'Género:            ' };
    mapa[C.TELEFONO]   = { p:'Teléfono:.*',          pre:'Teléfono:          ' };
    mapa[C.ZONA]       = { p:'Zona:.*',              pre:'Zona:              ' };
    mapa[C.EMAIL]      = { p:'Email:.*',             pre:'Email:             ' };
    mapa[C.EDUCACION]  = { p:'Educación:.*',         pre:'Educación:         ' };
    mapa[C.LABORAL]    = { p:'Situación laboral:.*', pre:'Situación laboral: ' };
    mapa[C.FORTALEZAS] = { p:'Fortalezas:.*',        pre:'Fortalezas:        ' };
    mapa[C.OBJETIVO]   = { p:'Objetivo laboral:.*',  pre:'Objetivo laboral:  ' };
    mapa[C.PERFIL]     = { p:'Perfil:.*',            pre:'Perfil:            ' };
    mapa[C.PRIORIDAD]  = { p:'Prioridad:.*',         pre:'Prioridad:         ' };

    const campo = mapa[col];
    if (campo) body.replaceText(campo.p, campo.pre + (valN || '—'));
    if (col === C.NOMBRE && valN) doc.setName('Perfil — ' + valN);

    const ts     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const user   = Session.getEffectiveUser().getEmail().split('@')[0];
    const nombres = {[C.NOMBRE]:'Nombre',[C.DPI]:'DPI',[C.EDAD]:'Edad',[C.GENERO]:'Género',
                     [C.TELEFONO]:'Teléfono',[C.ZONA]:'Zona',[C.EMAIL]:'Email',
                     [C.EDUCACION]:'Educación',[C.LABORAL]:'Situación Laboral',
                     [C.FORTALEZAS]:'Fortalezas',[C.OBJETIVO]:'Objetivo',
                     [C.PERFIL]:'Perfil',[C.PRIORIDAD]:'Prioridad',[C.ESTADO]:'Estado'};
    const linea  = ts + ' | ' + String(nombres[col]||'Campo').padEnd(18) + ' | ' +
                   String(valA||'—').substring(0,28).padEnd(28) + ' → ' +
                   String(valN||'—').substring(0,18) + ' | ' + user;

    if (body.getText().includes('HISTORIAL DE CAMBIOS')) body.appendParagraph(linea);
    doc.saveAndClose();
  } catch(e) {}
}

// ============================================================================
// DASHBOARD
// ============================================================================

function abrirDashboard() {
  try {
    actualizarDashboard();
    SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName('Dashboard'));
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e); }
}

function actualizarDashboard() {
  try {
    const ss   = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    const dash = ss.getSheetByName('Dashboard');
    if (!hoja || !dash) return;

    dash.clear();
    const M    = CONFIG.COL;
    const tz   = Session.getScriptTimeZone();
    const datos = hoja.getLastRow() > 1
      ? hoja.getRange(2, 1, hoja.getLastRow()-1, 26).getValues().filter(r => r[0])
      : [];

    const addFila = (a, b, bg, fg) => {
      dash.appendRow([a, b === undefined ? '' : b]);
      if (bg) dash.getRange(dash.getLastRow(),1,1,2).setBackground(bg);
      if (fg) dash.getRange(dash.getLastRow(),1,1,2).setFontColor(fg);
    };
    const addTit = (txt) => {
      dash.appendRow([txt]);
      dash.getRange(dash.getLastRow(),1).setFontWeight('bold').setFontColor('#1a237e').setFontSize(12);
    };

    // Encabezado
    dash.appendRow(['📊 DASHBOARD — PASO A PASO', '', '', 'Actualizado: ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm')]);
    dash.getRange(1,1).setFontSize(16).setFontWeight('bold').setFontColor('#1a237e');
    dash.getRange(1,4).setFontColor('#9e9e9e').setFontSize(10);
    dash.appendRow([]);

    // Resumen
    addTit('📋 RESUMEN');
    addFila('Total participantes', datos.length);
    dash.appendRow([]);

    // Por perfil
    addTit('🎯 POR PERFIL');
    const pc = {'Perfil A':0,'Perfil B':0,'Perfil C':0,'Perfil D':0,'Sin perfil':0};
    datos.forEach(r => { const p = normalizarPerfil(r[M.PERFIL-1]); pc[p] = (pc[p]||0)+1; });
    addFila('🟢 Perfil A — Listo para empleo',      pc['Perfil A'], '#d9ead3','#274e13');
    addFila('🔵 Perfil B — Orientación vocacional',  pc['Perfil B'], '#cfe2f3','#1c4587');
    addFila('🟡 Perfil C — Desarrollo de capacidades',pc['Perfil C'],'#fff2cc','#7f6000');
    addFila('🔴 Perfil D — Barreras críticas',       pc['Perfil D'], '#f4cccc','#660000');
    addFila('⬜ Sin perfil asignado',                pc['Sin perfil']||0);
    dash.appendRow([]);

    // Por estado
    addTit('📌 POR ESTADO');
    const est = {};
    datos.forEach(r => { const e = r[M.ESTADO-1]||'Sin estado'; est[e]=(est[e]||0)+1; });
    Object.entries(est).sort((a,b)=>b[1]-a[1]).forEach(([k,v]) => addFila(k, v));
    dash.appendRow([]);

    // Puntajes
    addTit('📈 PUNTAJE DIAGNÓSTICO');
    const pts = datos.map(r=>Number(r[M.PUNTAJE-1])).filter(v=>v>0);
    const prom = pts.length ? Math.round(pts.reduce((a,b)=>a+b,0)/pts.length) : 0;
    const maxP = pts.reduce((a,b)=>b>a?b:a,0);
    const minP = pts.length ? pts.reduce((a,b)=>b<a?b:a,pts[0]) : 0;
    addFila('Promedio', prom + ' / 60');
    addFila('Máximo',   pts.length ? maxP : '—');
    addFila('Mínimo',   pts.length ? minP : '—');

    dash.setColumnWidth(1,280); dash.setColumnWidth(2,120);
  } catch(e) {}
}

// ============================================================================
// ANALYTICS
// ============================================================================

function abrirAnalytics() {
  try {
    const ss = SpreadsheetApp.getActive();
    actualizarAnalytics(ss);
    ss.setActiveSheet(ss.getSheetByName('Analytics'));
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e); }
}

function actualizarAnalytics(ss) {
  try {
    if (!ss) ss = SpreadsheetApp.getActive();
    const hoja  = ss.getSheetByName(CONFIG.HOJA);
    const an    = ss.getSheetByName('Analytics');
    const deriv = ss.getSheetByName('Derivaciones');
    if (!hoja || !an) return;

    an.clear();
    const M    = CONFIG.COL;
    const tz   = Session.getScriptTimeZone();
    const datos = hoja.getLastRow() > 1
      ? hoja.getRange(2, 1, hoja.getLastRow()-1, 26).getValues().filter(r => r[0])
      : [];

    const tit = (txt) => {
      an.appendRow([txt]);
      an.getRange(an.getLastRow(),1).setFontWeight('bold').setFontColor('#6a1b9a').setFontSize(12);
    };
    const row2 = (a, b, bg, fg) => {
      an.appendRow([a, b]);
      if (bg) an.getRange(an.getLastRow(),1,1,2).setBackground(bg);
      if (fg) an.getRange(an.getLastRow(),1,1,2).setFontColor(fg);
    };

    // Encabezado
    an.appendRow(['ANALYTICS — PASO A PASO']);
    an.getRange(1,1).setFontSize(16).setFontWeight('bold').setFontColor('#6a1b9a');
    an.appendRow(['Actualizado: ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm') +
                  '  |  Total: ' + datos.length + ' participantes']);
    an.getRange(2,1).setFontColor('#9e9e9e');
    an.appendRow([]);

    // Perfiles
    tit('POR PERFIL');
    const pc = {'Perfil A':0,'Perfil B':0,'Perfil C':0,'Perfil D':0,'Sin perfil':0};
    datos.forEach(r => { const p = normalizarPerfil(r[M.PERFIL-1]); pc[p]=(pc[p]||0)+1; });
    row2('🟢 Perfil A — Listo para empleo',         pc['Perfil A'], '#d9ead3','#274e13');
    row2('🔵 Perfil B — Orientación vocacional',     pc['Perfil B'], '#cfe2f3','#1c4587');
    row2('🟡 Perfil C — Desarrollo de capacidades',  pc['Perfil C'], '#fff2cc','#7f6000');
    row2('🔴 Perfil D — Barreras críticas (URGENTE)',pc['Perfil D'], '#f4cccc','#660000');
    row2('⬜ Sin perfil', pc['Sin perfil']||0);
    an.appendRow([]);

    // Prioridades
    tit('POR PRIORIDAD');
    const prc = {'CRÍTICO':0,'ALTO':0,'MEDIO':0,'BAJO':0};
    datos.forEach(r => { const p = normalizarPrioridad(r[M.PRIORIDAD-1]); prc[p]=(prc[p]||0)+1; });
    row2('🔴 CRÍTICO', prc['CRÍTICO'], '#ffcdd2','#c62828');
    row2('🟠 ALTO',    prc['ALTO'],    '#ffe0b2','#e65100');
    row2('🟡 MEDIO',   prc['MEDIO'],   '#fff9c4','#f57f17');
    row2('🟢 BAJO',    prc['BAJO'],    '#c8e6c9','#2e7d32');
    an.appendRow([]);

    // Por estado
    tit('POR ESTADO');
    const est = {};
    datos.forEach(r => { const e = r[M.ESTADO-1]||'Sin estado'; est[e]=(est[e]||0)+1; });
    Object.entries(est).sort((a,b)=>b[1]-a[1]).forEach(([k,v]) => row2(k, v));
    an.appendRow([]);

    // Dimensiones promedio
    tit('PROMEDIO POR DIMENSIÓN (0-10)');
    [[M.DIM1,'📚 Capital Educativo'],[M.DIM2,'💼 Capital Laboral'],[M.DIM3,'💻 Habilidades Digitales'],
     [M.DIM4,'🎯 Claridad Vocacional'],[M.DIM5,'🚧 Barreras Estructurales'],[M.DIM6,'🤝 Red de Apoyo']
    ].forEach(([col, nombre]) => {
      const vals = datos.map(r => Number(r[col-1])).filter(v => v > 0);
      const prom = vals.length ? (vals.reduce((a,b)=>a+b,0)/vals.length).toFixed(1) : '—';
      const bg   = prom !== '—' ? (Number(prom)<4?'#ffcdd2':Number(prom)<7?'#fff9c4':'#c8e6c9') : null;
      row2(nombre, prom === '—' ? '—' : prom + ' / 10', bg);
    });
    an.appendRow([]);

    // Por zona
    tit('POR ZONA');
    const zonas = {};
    datos.forEach(r => { const z = r[M.ZONA-1]||'Sin zona'; zonas[z]=(zonas[z]||0)+1; });
    Object.entries(zonas).sort((a,b)=>b[1]-a[1]).slice(0,10).forEach(([z,n]) => row2(z, n));
    an.appendRow([]);

    // Derivaciones
    tit('DERIVACIONES');
    if (deriv && deriv.getLastRow() > 1) {
      const dd = deriv.getRange(2, 1, deriv.getLastRow()-1, 10).getValues().filter(r => r[0]);
      const urgentes    = dd.filter(r => r[3] === '🚨 URGENTE').length;
      const pendientes  = dd.filter(r => r[6] === 'Pendiente').length;
      const completadas = dd.filter(r => r[6] === 'Completado').length;
      row2('Total derivaciones',   dd.length);
      row2('🚨 URGENTES',          urgentes,   urgentes  > 0 ? '#ffcdd2' : null, urgentes  > 0 ? '#c62828' : null);
      row2('⏳ Pendientes',         pendientes, pendientes> 0 ? '#fff9c4' : null, pendientes> 0 ? '#f57f17' : null);
      row2('✅ Completadas',        completadas,'#c8e6c9','#2e7d32');
    } else {
      an.appendRow(['Sin derivaciones registradas todavía']);
    }
    an.appendRow([]);

    // Puntajes
    tit('DIAGNÓSTICO — PUNTAJE TOTAL (máx. 60)');
    const pts = datos.map(r => Number(r[M.PUNTAJE-1])).filter(v => v > 0);
    const prom = pts.length ? Math.round(pts.reduce((a,b)=>a+b,0)/pts.length) : 0;
    const maxP = pts.reduce((a,b)=>b>a?b:a,0);
    const minP = pts.length ? pts.reduce((a,b)=>b<a?b:a,pts[0]) : 0;
    row2('Promedio', prom + ' / 60');
    row2('Máximo',   pts.length ? maxP + ' / 60' : '—');
    row2('Mínimo',   pts.length ? minP + ' / 60' : '—');
    const bajo30 = pts.filter(p => p < 30).length;
    if (bajo30 > 0) row2('⚠️ Con puntaje < 30 (necesitan más apoyo)', bajo30, '#ffcdd2','#c62828');

    an.setColumnWidth(1,300); an.setColumnWidth(2,130);
  } catch(e) {}
}

// ============================================================================
// VER FICHA DEL PARTICIPANTE
// ============================================================================

function verFicha() {
  try {
    const ss    = SpreadsheetApp.getActive();
    const rango = ss.getActiveRange();
    if (rango.getSheet().getName() !== CONFIG.HOJA) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila en la hoja Maestro.');
      return;
    }
    if (rango.getRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila de datos, no el encabezado.');
      return;
    }
    const datos = rango.getSheet().getRange(rango.getRow(), 1, 1, 26).getValues()[0];
    if (!datos[0]) { SpreadsheetApp.getUi().alert('⚠️ Fila vacía.'); return; }
    mostrarFicha(datos);
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e); }
}

function mostrarFicha(datos) {
  const C        = CONFIG.COL;
  const nombre   = datos[C.NOMBRE-1]     || '—';
  const perfil   = datos[C.PERFIL-1]     || 'Sin perfil';
  const priorid  = datos[C.PRIORIDAD-1]  || 'Sin prioridad';
  const puntaje  = datos[C.PUNTAJE-1]    || 0;
  const estado   = datos[C.ESTADO-1]     || 'Orientación';
  const docUrl   = datos[C.DOC_URL-1]    || '';
  const dims     = [C.DIM1,C.DIM2,C.DIM3,C.DIM4,C.DIM5,C.DIM6].map(k => Number(datos[k-1])||0);

  const col  = CONFIG.COLORES[perfil] || {fondo:'#f5f5f5',texto:'#333'};
  const cP   = priorid==='CRÍTICO'?'#c62828':priorid==='ALTO'?'#e65100':priorid==='MEDIO'?'#f9a825':'#388e3c';
  const ini  = nombre.split(' ').slice(0,2).map(p=>p[0]||'').join('').toUpperCase();
  const pct  = Math.min(100, Math.round((puntaje/60)*100));
  const docB = docUrl
    ? '<a href="'+docUrl+'" target="_blank" style="display:block;text-align:center;background:#1a237e;color:#fff;padding:9px 12px;border-radius:5px;margin:10px 0 0;text-decoration:none;font-size:12px;font-weight:bold">📄 Abrir Expediente en Drive</a>'
    : '';

  const lbs  = ['Educativo','Laboral','Digital','Vocacional','Barreras','Red Apoyo'];
  const barras = lbs.map((lb,i) =>
    '<div style="margin:4px 0"><div style="display:flex;justify-content:space-between;font-size:10px;color:#777;margin-bottom:2px"><span>'+lb+'</span><span>'+dims[i]+'/10</span></div><div style="background:#e8eaf6;border-radius:3px;height:5px"><div style="width:'+Math.round(dims[i]*10)+'%;background:'+col.texto+';height:5px;border-radius:3px"></div></div></div>'
  ).join('');

  // SVG radar
  const cx=85,cy=85,r=65,ang=[-90,-30,30,90,150,210],lbsR=['Educ','Labor','Digit','Vocal','Barr','Apoyo'];
  let svg='<svg width="170" height="170" viewBox="0 0 170 170">';
  [0.33,0.66,1].forEach(n=>{
    const pts=ang.map(a=>{const rad=a*Math.PI/180;return(cx+r*n*Math.cos(rad)).toFixed(1)+','+(cy+r*n*Math.sin(rad)).toFixed(1);}).join(' ');
    svg+='<polygon points="'+pts+'" fill="none" stroke="#e0e0e0" stroke-width="1"/>';
  });
  ang.forEach(a=>{const rad=a*Math.PI/180;svg+='<line x1="'+cx+'" y1="'+cy+'" x2="'+(cx+r*Math.cos(rad)).toFixed(1)+'" y2="'+(cy+r*Math.sin(rad)).toFixed(1)+'" stroke="#e0e0e0" stroke-width="1"/>';});
  const dpts=ang.map((a,i)=>{const rad=a*Math.PI/180,esc=dims[i]/10;return(cx+r*esc*Math.cos(rad)).toFixed(1)+','+(cy+r*esc*Math.sin(rad)).toFixed(1);}).join(' ');
  ang.forEach((a,i)=>{const rad=a*Math.PI/180;svg+='<text x="'+(cx+(r+14)*Math.cos(rad)).toFixed(1)+'" y="'+(cy+(r+14)*Math.sin(rad)).toFixed(1)+'" text-anchor="middle" font-size="8" fill="#777">'+lbsR[i]+'</text>';});
  svg+='<polygon points="'+dpts+'" fill="'+col.texto+'33" stroke="'+col.texto+'" stroke-width="2"/></svg>';

  const campo = (lbl, val) =>
    '<div class="row"><span class="rl">'+lbl+'</span><span class="rv">'+escaparHtml(String(val||'—'))+'</span></div>';

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    'body{font-family:Arial,sans-serif;font-size:13px;background:#f5f7ff;color:#333}' +
    '.hdr{background:'+col.fondo+';padding:14px 16px;border-left:5px solid '+col.texto+';display:flex;gap:12px;align-items:center}' +
    '.av{width:50px;height:50px;border-radius:50%;background:'+col.texto+';color:#fff;font-size:17px;font-weight:bold;display:flex;align-items:center;justify-content:center;flex-shrink:0}' +
    '.nom{font-size:14px;font-weight:bold;color:'+col.texto+';line-height:1.3}' +
    '.tags{display:flex;gap:4px;flex-wrap:wrap;margin-top:5px}' +
    '.tag{padding:2px 8px;border-radius:10px;font-size:10px;font-weight:bold}' +
    '.tp{background:'+col.texto+';color:'+col.fondo+'}.tpr{background:'+cP+';color:#fff}.te{background:#e0e0e0;color:#555}' +
    '.prog{padding:10px 16px;background:#fff;border-bottom:1px solid #e8eaf6}' +
    '.pl{display:flex;justify-content:space-between;font-size:11px;color:#9e9e9e;margin-bottom:4px}' +
    '.pb{height:8px;background:#e8eaf6;border-radius:4px;overflow:hidden}' +
    '.pf{height:8px;background:'+col.texto+';width:'+pct+'%;border-radius:4px}' +
    '.sec{padding:10px 16px;background:#fff;border-bottom:1px solid #f0f0f0}' +
    '.sh{font-size:9px;font-weight:bold;color:#9e9e9e;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}' +
    '.row{display:flex;justify-content:space-between;padding:3px 0;font-size:12px}' +
    '.rl{color:#777;flex-shrink:0;margin-right:8px}.rv{color:#333;font-weight:500;text-align:right}' +
    '.desc{font-size:12px;color:#444;line-height:1.5}' +
    '</style></head><body>' +
    '<div class="hdr"><div class="av">'+ini+'</div><div><div class="nom">'+escaparHtml(nombre)+'</div>' +
    '<div class="tags"><span class="tag tp">'+escaparHtml(perfil)+'</span>' +
    '<span class="tag tpr">'+escaparHtml(priorid)+'</span>' +
    '<span class="tag te">'+escaparHtml(estado)+'</span></div></div></div>' +
    '<div class="prog"><div class="pl"><span>Puntaje diagnóstico</span><span>'+puntaje+' / 60</span></div>' +
    '<div class="pb"><div class="pf"></div></div></div>' +
    '<div class="sec"><div class="sh">Datos Personales</div>' +
    campo('DPI',      datos[C.DPI-1]) +
    campo('Edad',     datos[C.EDAD-1]) +
    campo('Género',   datos[C.GENERO-1]) +
    campo('Teléfono', datos[C.TELEFONO-1]) +
    campo('Zona',     datos[C.ZONA-1]) +
    campo('Email',    datos[C.EMAIL-1]) + '</div>' +
    '<div class="sec"><div class="sh">Perfil Profesional</div>' +
    campo('Educación',         datos[C.EDUCACION-1]) +
    campo('Situación laboral', datos[C.LABORAL-1]) + '</div>' +
    '<div class="sec"><div class="sh">Fortalezas</div><div class="desc">'+escaparHtml(String(datos[C.FORTALEZAS-1]||'—'))+'</div></div>' +
    '<div class="sec"><div class="sh">Objetivo laboral</div><div class="desc">'+escaparHtml(String(datos[C.OBJETIVO-1]||'—'))+'</div></div>' +
    '<div class="sec"><div class="sh">Dimensiones de diagnóstico</div>' +
    '<div style="display:flex;justify-content:center;margin:6px 0">'+svg+'</div>' +
    barras + docB + '</div></body></html>'
  ).setTitle('Perfil — '+nombre).setWidth(360);

  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================================
// AGREGAR PARTICIPANTE MANUALMENTE
// ============================================================================

function agregarParticipanteManual() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    '*{box-sizing:border-box}body{font-family:Arial;padding:16px;background:#f5f7ff;font-size:12px}' +
    'h3{color:#1a237e;margin:0 0 14px;font-size:14px}' +
    'label{display:block;font-weight:bold;margin:10px 0 3px;color:#333}' +
    'input,select,textarea{width:100%;padding:7px;border:1px solid #c5cae9;border-radius:4px;margin-bottom:8px;font-size:12px}' +
    'textarea{height:60px;resize:vertical}' +
    'button{background:#1a237e;color:#fff;padding:10px;border:none;border-radius:4px;cursor:pointer;width:100%;font-weight:bold;margin-top:10px}' +
    'button:hover{background:#283593}' +
    '.row{display:grid;grid-template-columns:1fr 1fr;gap:8px}' +
    '.full{grid-column:1/-1}' +
    '</style></head><body>' +
    '<h3>➕ Agregar Nuevo Participante</h3>' +
    '<div class="row">' +
    '<div><label>ID (opcional)</label><input id="id" placeholder="Auto-generado si está vacío"></div>' +
    '<div><label>Nombre completo *</label><input id="nombre" placeholder="Nombre y apellido" required></div>' +
    '<div><label>DPI</label><input id="dpi" placeholder="13 dígitos"></div>' +
    '<div><label>Edad</label><input id="edad" type="number" min="15" max="120"></div>' +
    '<div><label>Género</label><select id="genero"><option>—</option><option>Masculino</option><option>Femenino</option><option>Otro</option></select></div>' +
    '<div><label>Teléfono</label><input id="tel" placeholder="+502 XXXX XXXX"></div>' +
    '<div><label>Zona/Región</label><input id="zona" placeholder="Ej: Zona 3, Guatemala"></div>' +
    '<div><label>Email</label><input id="email" type="email" placeholder="correo@ejemplo.com"></div>' +
    '<div class="full"><label>Educación</label><input id="educacion" placeholder="Nivel educativo alcanzado"></div>' +
    '<div class="full"><label>Situación laboral actual</label><input id="laboral" placeholder="Empleado, desempleado, etc."></div>' +
    '<div class="full"><label>Fortalezas/Habilidades</label><textarea id="fortalezas" placeholder="¿Qué sabe hacer bien?"></textarea></div>' +
    '<div class="full"><label>Objetivo laboral</label><textarea id="objetivo" placeholder="¿Qué tipo de empleo busca?"></textarea></div>' +
    '<div><label>Estado inicial</label><select id="estado"><option>Orientación</option><option>Mentoría</option><option>Formación</option><option>Cierre</option><option>Inactivo</option></select></div>' +
    '<div><label>Perfil asignado</label><select id="perfil"><option>—</option><option>Perfil A</option><option>Perfil B</option><option>Perfil C</option><option>Perfil D</option></select></div>' +
    '</div>' +
    '<button onclick="guardar()">✅ Guardar Participante</button>' +
    '<script>' +
    'function guardar(){' +
    'var n=document.getElementById("nombre").value.trim();' +
    'if(!n){alert("El nombre es obligatorio");return;}' +
    'var obj={' +
    'id:document.getElementById("id").value.trim(),' +
    'nombre:n,' +
    'dpi:document.getElementById("dpi").value.trim(),' +
    'edad:document.getElementById("edad").value.trim(),' +
    'genero:document.getElementById("genero").value,' +
    'tel:document.getElementById("tel").value.trim(),' +
    'zona:document.getElementById("zona").value.trim(),' +
    'email:document.getElementById("email").value.trim(),' +
    'educacion:document.getElementById("educacion").value.trim(),' +
    'laboral:document.getElementById("laboral").value.trim(),' +
    'fortalezas:document.getElementById("fortalezas").value.trim(),' +
    'objetivo:document.getElementById("objetivo").value.trim(),' +
    'estado:document.getElementById("estado").value,' +
    'perfil:document.getElementById("perfil").value' +
    '};' +
    'google.script.run.withSuccessHandler(function(){alert("✅ Participante agregado");google.script.host.close();}).guardarParticipanteManual(obj);' +
    '}' +
    '</script></body></html>'
  ).setTitle('➕ Nuevo Participante').setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

function guardarParticipanteManual(datos) {
  try {
    const ss   = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) return;

    let id = datos.id;
    if (!id) {
      id = datos.nombre.replace(/\s+/g,'').substring(0,4).toUpperCase() +
           Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMMyyyy');
    }

    const M   = CONFIG.COL;
    const row = new Array(26).fill('');
    row[M.ID-1]         = id;
    row[M.FECHA-1]      = new Date();
    row[M.NOMBRE-1]     = datos.nombre;
    row[M.DPI-1]        = datos.dpi;
    row[M.EDAD-1]       = datos.edad;
    row[M.GENERO-1]     = datos.genero;
    row[M.TELEFONO-1]   = datos.tel;
    row[M.ZONA-1]       = datos.zona;
    row[M.EMAIL-1]      = datos.email;
    row[M.EDUCACION-1]  = datos.educacion;
    row[M.LABORAL-1]    = datos.laboral;
    row[M.FORTALEZAS-1] = datos.fortalezas;
    row[M.OBJETIVO-1]   = datos.objetivo;
    row[M.PERFIL-1]     = datos.perfil === '—' ? '' : datos.perfil;
    row[M.ESTADO-1]     = datos.estado;

    hoja.appendRow(row);
    colorearFila(hoja, hoja.getLastRow(), datos.perfil);
    SpreadsheetApp.flush();
  } catch(e) {
    logError('guardarParticipanteManual', e);
  }
}

// ============================================================================
// EDITAR PARTICIPANTE
// ============================================================================

function editarParticipante() {
  try {
    const ss    = SpreadsheetApp.getActive();
    const rango = ss.getActiveRange();
    if (rango.getSheet().getName() !== CONFIG.HOJA) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila en la hoja Maestro.');
      return;
    }
    if (rango.getRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila de participante (no el encabezado).');
      return;
    }
    const datos = rango.getSheet().getRange(rango.getRow(), 1, 1, 26).getValues()[0];
    if (!datos[0]) { SpreadsheetApp.getUi().alert('⚠️ Fila vacía, no hay participante.'); return; }

    const C = CONFIG.COL;
    const html = HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
      '*{box-sizing:border-box}body{font-family:Arial;padding:16px;background:#f5f7ff;font-size:12px}' +
      'h3{color:#1a237e;margin:0 0 14px;font-size:14px}' +
      'label{display:block;font-weight:bold;margin:10px 0 3px;color:#333}' +
      'input,select,textarea{width:100%;padding:7px;border:1px solid #c5cae9;border-radius:4px;margin-bottom:8px;font-size:12px}' +
      'textarea{height:50px;resize:vertical}' +
      'button{background:#1a237e;color:#fff;padding:10px;border:none;border-radius:4px;cursor:pointer;width:48%;font-weight:bold;margin-right:2%;margin-top:10px}' +
      'button:last-child{margin-right:0}button:hover{background:#283593}' +
      '.row{display:grid;grid-template-columns:1fr 1fr;gap:8px}.full{grid-column:1/-1}' +
      '</style></head><body>' +
      '<h3>✏️ Editar Participante</h3>' +
      '<div class="row">' +
      '<div><label>ID</label><input id="id" value="'+escaparHtml(String(datos[C.ID-1]||''))+'" disabled style="background:#f0f0f0"></div>' +
      '<div><label>Nombre</label><input id="nombre" value="'+escaparHtml(String(datos[C.NOMBRE-1]||''))+'"></div>' +
      '<div><label>DPI</label><input id="dpi" value="'+escaparHtml(String(datos[C.DPI-1]||''))+'"></div>' +
      '<div><label>Edad</label><input id="edad" type="number" value="'+(datos[C.EDAD-1]||'')+'"></div>' +
      '<div><label>Género</label><select id="genero"><option>—</option><option '+(datos[C.GENERO-1]==='Masculino'?'selected':'')+'>Masculino</option><option '+(datos[C.GENERO-1]==='Femenino'?'selected':'')+'>Femenino</option><option '+(datos[C.GENERO-1]==='Otro'?'selected':'')+'>Otro</option></select></div>' +
      '<div><label>Teléfono</label><input id="tel" value="'+escaparHtml(String(datos[C.TELEFONO-1]||''))+'"></div>' +
      '<div><label>Zona</label><input id="zona" value="'+escaparHtml(String(datos[C.ZONA-1]||''))+'"></div>' +
      '<div><label>Email</label><input id="email" type="email" value="'+escaparHtml(String(datos[C.EMAIL-1]||''))+'"></div>' +
      '<div class="full"><label>Educación</label><input id="educacion" value="'+escaparHtml(String(datos[C.EDUCACION-1]||''))+'"></div>' +
      '<div class="full"><label>Situación laboral</label><input id="laboral" value="'+escaparHtml(String(datos[C.LABORAL-1]||''))+'"></div>' +
      '<div class="full"><label>Fortalezas</label><textarea id="fortalezas">'+escaparHtml(String(datos[C.FORTALEZAS-1]||''))+'</textarea></div>' +
      '<div class="full"><label>Objetivo laboral</label><textarea id="objetivo">'+escaparHtml(String(datos[C.OBJETIVO-1]||''))+'</textarea></div>' +
      '<div><label>Perfil</label><select id="perfil"><option>—</option><option '+(datos[C.PERFIL-1]==='Perfil A'?'selected':'')+'>Perfil A</option><option '+(datos[C.PERFIL-1]==='Perfil B'?'selected':'')+'>Perfil B</option><option '+(datos[C.PERFIL-1]==='Perfil C'?'selected':'')+'>Perfil C</option><option '+(datos[C.PERFIL-1]==='Perfil D'?'selected':'')+'>Perfil D</option></select></div>' +
      '<div><label>Estado</label><select id="estado"><option '+(datos[C.ESTADO-1]==='Orientación'?'selected':'')+'>Orientación</option><option '+(datos[C.ESTADO-1]==='Mentoría'?'selected':'')+'>Mentoría</option><option '+(datos[C.ESTADO-1]==='Formación'?'selected':'')+'>Formación</option><option '+(datos[C.ESTADO-1]==='Cierre'?'selected':'')+'>Cierre</option><option '+(datos[C.ESTADO-1]==='Inactivo'?'selected':'')+'>Inactivo</option></select></div>' +
      '<button onclick="google.script.host.close()">✗ Cancelar</button>' +
      '<button onclick="guardar()">✅ Guardar Cambios</button>' +
      '</div>' +
      '<script>' +
      'function guardar(){var obj={nombre:document.getElementById("nombre").value.trim(),dpi:document.getElementById("dpi").value.trim(),edad:document.getElementById("edad").value.trim(),genero:document.getElementById("genero").value,tel:document.getElementById("tel").value.trim(),zona:document.getElementById("zona").value.trim(),email:document.getElementById("email").value.trim(),educacion:document.getElementById("educacion").value.trim(),laboral:document.getElementById("laboral").value.trim(),fortalezas:document.getElementById("fortalezas").value.trim(),objetivo:document.getElementById("objetivo").value.trim(),perfil:document.getElementById("perfil").value,estado:document.getElementById("estado").value};' +
      'google.script.run.withSuccessHandler(function(){alert("✅ Cambios guardados");google.script.host.close();}).guardarEdicionParticipante('+rango.getRow()+',obj);' +
      '}' +
      '</script></body></html>'
    ).setTitle('✏️ Editar Participante').setWidth(380);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e); }
}

function guardarEdicionParticipante(fila, datos) {
  try {
    const ss   = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) return;
    const M = CONFIG.COL;
    hoja.getRange(fila, M.NOMBRE,     1, 1).setValue(datos.nombre);
    hoja.getRange(fila, M.DPI,        1, 1).setValue(datos.dpi);
    hoja.getRange(fila, M.EDAD,       1, 1).setValue(datos.edad);
    hoja.getRange(fila, M.GENERO,     1, 1).setValue(datos.genero);
    hoja.getRange(fila, M.TELEFONO,   1, 1).setValue(datos.tel);
    hoja.getRange(fila, M.ZONA,       1, 1).setValue(datos.zona);
    hoja.getRange(fila, M.EMAIL,      1, 1).setValue(datos.email);
    hoja.getRange(fila, M.EDUCACION,  1, 1).setValue(datos.educacion);
    hoja.getRange(fila, M.LABORAL,    1, 1).setValue(datos.laboral);
    hoja.getRange(fila, M.FORTALEZAS, 1, 1).setValue(datos.fortalezas);
    hoja.getRange(fila, M.OBJETIVO,   1, 1).setValue(datos.objetivo);
    hoja.getRange(fila, M.PERFIL,     1, 1).setValue(datos.perfil === '—' ? '' : datos.perfil);
    hoja.getRange(fila, M.ESTADO,     1, 1).setValue(datos.estado);
    colorearFila(hoja, fila, datos.perfil);
    SpreadsheetApp.flush();
  } catch(e) { logError('guardarEdicionParticipante', e); }
}

// ============================================================================
// FORMULARIO DE DERIVACIÓN MEJORADO
// ============================================================================

function abrirFormDerivacion() {
  try {
    const ss    = SpreadsheetApp.getActive();
    const hoja  = ss.getSheetByName(CONFIG.HOJA);
    const rango = ss.getActiveRange();
    if (!hoja || rango.getRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona un participante en Maestro.');
      return;
    }
    const datos = hoja.getRange(rango.getRow(), 1, 1, 26).getValues()[0];
    const C = CONFIG.COL;
    const id   = datos[C.ID-1]   || '';
    const nom  = datos[C.NOMBRE-1] || '';
    if (!id) { SpreadsheetApp.getUi().alert('⚠️ Participante sin ID.'); return; }

    const html = HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
      '*{box-sizing:border-box}body{font-family:Arial;padding:16px;background:#f5f7ff;font-size:12px}' +
      '.info{background:#e3f2fd;border-left:4px solid #1a237e;padding:12px;margin-bottom:14px;border-radius:4px}' +
      '.info strong{color:#1a237e}' +
      'label{display:block;font-weight:bold;margin:12px 0 4px;color:#333}' +
      'input,select,textarea{width:100%;padding:8px;border:1px solid #c5cae9;border-radius:4px;margin-bottom:8px;font-size:12px}' +
      'textarea{height:80px;resize:vertical}' +
      'button{background:#1a237e;color:#fff;padding:10px;border:none;border-radius:4px;cursor:pointer;width:100%;font-weight:bold;margin-top:10px}' +
      'button:hover{background:#283593}' +
      '</style></head><body>' +
      '<h3>➡️ Registrar Derivación</h3>' +
      '<div class="info"><strong>Participante:</strong> '+escaparHtml(nom)+'<br><strong>ID:</strong> '+escaparHtml(id)+'</div>' +
      '<label>Tipo de derivación *</label>' +
      '<select id="tipo" required>' +
      '<option value="">Selecciona...</option>' +
      '<option value="💡 SUGERIDA">💡 SUGERIDA — Recomendación</option>' +
      '<option value="⚠️ ALERTA">⚠️ ALERTA — Requiere atención</option>' +
      '<option value="🚨 URGENTE">🚨 URGENTE — Caso crítico</option>' +
      '</select>' +
      '<label>Organización/Programa destino *</label>' +
      '<select id="org" required>' +
      '<option value="">Selecciona...</option>' +
      '<option>Creamos Voces (Apoyo Emocional)</option>' +
      '<option>Mentoría Vocacional</option>' +
      '<option>Intermediación Laboral</option>' +
      '<option>Formación Técnica</option>' +
      '<option>Educación de Adultos</option>' +
      '<option>Alfabetización Digital</option>' +
      '<option>Servicios Profesionales (Legal/Salud)</option>' +
      '<option>Grupos de Apoyo Comunitario</option>' +
      '<option>Otro programa Creamos</option>' +
      '<option>Entidad pública</option>' +
      '<option>ONG externa</option>' +
      '<option>Otra</option>' +
      '</select>' +
      '<label>Motivo/Justificación *</label>' +
      '<textarea id="motivo" placeholder="¿Por qué se deriva a este programa?" required></textarea>' +
      '<label>Notas adicionales</label>' +
      '<textarea id="notas" placeholder="Información complementaria..." maxlength="500"></textarea>' +
      '<button onclick="guardar()">✅ Registrar Derivación</button>' +
      '<script>' +
      'function guardar(){' +
      'var t=document.getElementById("tipo").value,' +
      'o=document.getElementById("org").value,' +
      'm=document.getElementById("motivo").value.trim(),' +
      'n=document.getElementById("notas").value.trim();' +
      'if(!t||!o||!m){alert("Completa: Tipo, Organización y Motivo");return;}' +
      'google.script.run.withSuccessHandler(function(){alert("✅ Derivación registrada");google.script.host.close();}).guardarDerivacion("'+escaparHtml(id)+'","'+escaparHtml(nom)+'",t,o,m,n);' +
      '}' +
      '</script></body></html>'
    ).setTitle('➡️ Derivar Participante').setWidth(380);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e); }
}

// ============================================================================
// DERIVACIONES
// ============================================================================

function verDerivaciones() {
  try {
    const ss    = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('Derivaciones');
    if (!sheet) { SpreadsheetApp.getUi().alert('❌ Instala el sistema primero.'); return; }
    ss.setActiveSheet(sheet);
    const pend = sheet.getLastRow() > 1
      ? sheet.getRange(2, 7, sheet.getLastRow()-1, 1).getValues().filter(r => r[0]==='Pendiente').length
      : 0;
    if (pend > 0) SpreadsheetApp.getUi().alert('📋 Derivaciones\n\n⚠️ Tienes ' + pend + ' derivación(es) PENDIENTE(S).\n\nCambia el Estado a "Completado" cuando las gestiones.');
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e); }
}

function guardarDerivacion(id, nombre, tipo, destino, motivo, notas) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Derivaciones');
    if (!sheet) return;
    sheet.appendRow([new Date(), id, nombre, tipo||'Manual', destino, motivo, 'Pendiente',
                     Session.getEffectiveUser().getEmail(), '', notas||'']);
    const n     = sheet.getLastRow();
    const color = tipo && tipo.includes('URGENTE') ? '#ffcdd2' : tipo && tipo.includes('ALERTA') ? '#fff9c4' : '#e8f5e9';
    sheet.getRange(n, 1, 1, 10).setBackground(color);
  } catch(e) { logError('guardarDerivacion', e); }
}

function registrarDerivacionesAutomaticas(ss, id, nombre, fila, M) {
  try {
    const deriv = ss.getSheetByName('Derivaciones');
    if (!deriv) return;
    const perfil    = fila[M.PERFIL-1]    || '';
    const prioridad = fila[M.PRIORIDAD-1] || '';
    const dim5      = Number(fila[M.DIM5-1]) || 0;
    const dim3      = Number(fila[M.DIM3-1]) || 0;
    const dim1      = Number(fila[M.DIM1-1]) || 0;
    const dim6      = Number(fila[M.DIM6-1]) || 0;
    const hoy       = new Date();

    const add = (tipo, dest, motivo) => {
      deriv.appendRow([hoy, id, nombre, tipo, dest, motivo, 'Pendiente','','','']);
      const c = tipo==='🚨 URGENTE' ? '#ffcdd2' : '#fff9c4';
      deriv.getRange(deriv.getLastRow(), 1, 1, 10).setBackground(c);
    };

    if (perfil==='Perfil D' || prioridad==='CRÍTICO') {
      add('🚨 URGENTE','Creamos Voces (Apoyo Emocional)','Barreras críticas — Perfil ' + perfil);
      if (dim5 <= 4) add('🚨 URGENTE','Servicios Profesionales','Barreras estructurales: ' + dim5 + '/10');
      enviarAlertaCritica(id, nombre, perfil, dim5);
    }
    if (perfil === 'Perfil B') add('💡 SUGERIDA','Mentoría Vocacional','Requiere orientación vocacional — 3-6 sesiones');
    if (perfil === 'Perfil C') {
      if (dim1 <= 4) add('💡 SUGERIDA','Educación de Adultos','Capital educativo bajo: ' + dim1 + '/10');
      if (dim3 <= 4) add('💡 SUGERIDA','Alfabetización Digital','Habilidades digitales bajas: ' + dim3 + '/10');
    }
    if (dim5 <= 3 && perfil !== 'Perfil D') add('⚠️ ALERTA','Creamos Voces','Barreras muy bajas: ' + dim5 + '/10');
    if (dim6 <= 3) add('💡 SUGERIDA','Grupos de Apoyo Comunitario','Red de apoyo débil: ' + dim6 + '/10');
    if (perfil === 'Perfil A') add('✅ OPORTUNIDAD','Intermediación Laboral','Lista para empleo — ' + Number(fila[M.PUNTAJE-1]) + '/60 pts');
  } catch(e) {}
}

function enviarAlertaCritica(id, nombre, perfil, dim5) {
  try {
    MailApp.sendEmail(
      getAdminEmail(),
      '🚨 URGENTE: Barreras críticas — ' + nombre,
      '🚨 CASO URGENTE\n\nParticipante: ' + nombre + '\nID: ' + id +
      '\nPerfil: ' + perfil + '\nBarreras estructurales: ' + dim5 + '/10\n\n' +
      'ACCIONES REQUERIDAS:\n• Derivar a Creamos Voces\n• Derivar a Servicios Profesionales\n• No iniciar proceso de empleo hasta resolver barreras\n\nSistema Paso a Paso'
    );
  } catch(e) {}
}

// ============================================================================
// ALERTAS AUTOMÁTICAS
// ============================================================================

function verificarCasosDormidos() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) return;
    const M      = CONFIG.COL;
    const datos  = hoja.getRange(2, 1, hoja.getLastRow()-1, 26).getValues();
    const hoy    = new Date();
    const dormidos = [];

    datos.forEach((fila, i) => {
      if (!fila[0]) return;
      const dias = Math.floor((hoy - new Date(fila[M.FECHA-1])) / 86400000);
      if (dias > 30) {
        hoja.getRange(i+2, 1, 1, 26).setBackground('#b71c1c').setFontColor('#fff');
        dormidos.push({ nombre: fila[M.NOMBRE-1], dias, estado: fila[M.ESTADO-1] });
      }
    });

    if (dormidos.length > 0) {
      let body = dormidos.length + ' participantes sin seguimiento (30+ días):\n\n';
      dormidos.forEach(c => body += '• ' + c.nombre + ' (' + c.dias + ' días) — ' + c.estado + '\n');
      MailApp.sendEmail(getAdminEmail(), '🚨 Alerta — ' + dormidos.length + ' casos sin seguimiento', body);
    }
  } catch(e) {}
}

function enviarReporteSemanal() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) return;
    const M     = CONFIG.COL;
    const tz    = Session.getScriptTimeZone();
    const datos = hoja.getRange(2, 1, hoja.getLastRow()-1, 26).getValues().filter(r => r[0]);
    const hoy   = new Date();
    const hace7 = new Date(hoy.getTime() - 7*24*60*60*1000);

    const nuevos   = datos.filter(r => new Date(r[M.FECHA-1]) >= hace7);
    const criticos = datos.filter(r => normalizarPrioridad(r[M.PRIORIDAD-1]) === 'CRÍTICO');

    let body = '══════════════════════════════\n📊 REPORTE SEMANAL — PASO A PASO\n══════════════════════════════\n\n';
    body += 'Semana: ' + Utilities.formatDate(hace7,tz,'dd/MM') + ' - ' + Utilities.formatDate(hoy,tz,'dd/MM/yyyy') + '\n\n';
    body += '✨ Nuevos: ' + nuevos.length + '\n';
    nuevos.forEach(r => body += '  • ' + r[M.NOMBRE-1] + '\n');
    body += '\n🔴 Críticos: ' + criticos.length + '\n';
    criticos.forEach(r => body += '  • ' + r[M.NOMBRE-1] + ' — ' + r[M.ESTADO-1] + '\n');
    body += '\nTotal participantes: ' + datos.length + '\n\nSistema Paso a Paso';

    MailApp.sendEmail(getAdminEmail(), '📊 Reporte semanal — Paso a Paso', body);
  } catch(e) {}
}

// ============================================================================
// CALENDAR + EMAIL AL PARTICIPANTE
// ============================================================================

function crearEventoCalendario(nombre, estado, docUrl) {
  try {
    const cfg = {
      'Mentoría':   { dias:7,  titulo:'📋 Seguimiento: '   + nombre, desc:'Primera sesión de seguimiento en mentoría.' },
      'Formación':  { dias:14, titulo:'📚 Revisión: '      + nombre, desc:'Revisión de avance en formación técnica.' },
      'Completado': { dias:30, titulo:'🏁 Post-cierre: '   + nombre, desc:'Revisión de resultados post-cierre.' },
      'Cierre':     { dias:30, titulo:'🏁 Post-cierre: '   + nombre, desc:'Revisión de resultados post-cierre.' }
    };
    const conf = cfg[estado];
    if (!conf) return;
    const inicio = new Date(Date.now() + conf.dias * 86400000);
    inicio.setHours(9,0,0,0);
    const fin = new Date(inicio.getTime() + 30*60000);
    CalendarApp.getDefaultCalendar().createEvent(conf.titulo, inicio, fin,
      { description: conf.desc + (docUrl ? '\n\nExpediente: ' + docUrl : '') });
  } catch(e) {}
}

function notificarParticipante(nombre, emailP, estado) {
  try {
    const corto = nombre.split(' ')[0];
    const msgs  = {
      'Mentoría':   { asunto:'✅ Tu proceso avanzó a Mentoría — Paso a Paso',
                      body:'Hola '+corto+',\n\nTu proceso ha avanzado a MENTORÍA. Recibirás acompañamiento personalizado.\n\n💪 ¡Seguimos adelante!\nEquipo Creamos' },
      'Formación':  { asunto:'📚 Iniciaste Formación — Paso a Paso',
                      body:'Hola '+corto+',\n\n🎉 ¡Iniciaste tu proceso de Formación! Un gran paso.\n\n¡Seguimos contigo!\nEquipo Creamos' },
      'Completado': { asunto:'🏆 Completaste el programa — Creamos',
                      body:'Hola '+corto+',\n\n🎊 ¡FELICITACIONES! Completaste el programa Paso a Paso.\n\n🌟 ¡Mucho éxito!\nEquipo Creamos' }
    };
    const m = msgs[estado];
    if (!m) return;
    MailApp.sendEmail({ to: emailP, subject: m.asunto, body: m.body, replyTo: getAdminEmail() });
  } catch(e) {}
}

// ============================================================================
// CONFIGURACIÓN
// ============================================================================

function abrirConfiguracion() {
  const props    = PropertiesService.getUserProperties();
  const tieneKey = !!props.getProperty('KOBO_API_KEY');
  const folderId = escaparHtml(props.getProperty('FOLDER_ID') || CONFIG.FOLDER_PARTICIPANTES_ID);
  const email    = escaparHtml(props.getProperty('ADMIN_EMAIL') || CONFIG.ADMIN_EMAIL);

  const html = HtmlService.createHtmlOutput(
    '<style>*{box-sizing:border-box}' +
    'body{font-family:Arial,sans-serif;padding:16px;background:#f5f7ff;font-size:13px}' +
    'h3{color:#1a237e;margin:0 0 12px;font-size:14px}' +
    'label{display:block;font-weight:bold;margin:12px 0 4px;color:#444;font-size:12px}' +
    'input,textarea{width:100%;padding:8px;border:1px solid #c5cae9;border-radius:5px;font-size:12px;background:#fff;font-family:inherit}' +
    'textarea{height:60px;resize:vertical;font-family:monospace}' +
    'button{margin-top:14px;background:#1a237e;color:#fff;padding:10px;border:none;border-radius:5px;cursor:pointer;width:100%;font-weight:bold;font-size:13px}' +
    'button:hover{background:#283593}' +
    '.chip{display:inline-block;padding:4px 10px;border-radius:12px;font-size:11px;font-weight:bold;margin-bottom:10px}' +
    '.ok{background:#e8f5e9;color:#2e7d32}.no{background:#ffcdd2;color:#c62828}' +
    '.hint{font-size:10px;color:#9e9e9e;margin-top:3px}' +
    '</style>' +
    '<h3>⚙️ Configuración del Sistema</h3>' +
    '<div class="chip ' + (tieneKey?'ok':'no') + '">' + (tieneKey?'✅ Kobo configurado':'⚠️ Kobo no configurado') + '</div>' +
    '<label>Email para alertas</label><input id="em" type="email" value="' + email + '" maxlength="100">' +
    '<label>ID Carpeta Drive</label><input id="ci" value="' + folderId + '" maxlength="70">' +
    '<div class="hint">URL: drive.google.com/drive/folders/<strong>ESTE_ES_EL_ID</strong></div>' +
    '<label>API Key Kobo</label><textarea id="kk" placeholder="' + (tieneKey?'Configurada — deja vacío para no cambiar':'Kobo → Cuenta → Seguridad → Clave API') + '" maxlength="80"></textarea>' +
    '<button onclick="save()">💾 Guardar</button>' +
    '<script>function save(){var e=document.getElementById("em").value.trim(),c=document.getElementById("ci").value.trim(),k=document.getElementById("kk").value.trim();' +
    'if(!e||!c){alert("Email y Carpeta Drive son obligatorios");return;}' +
    'google.script.run.withSuccessHandler(function(){alert("✅ Configuración guardada");google.script.host.close();}).guardarConfig(e,c,k);}</script>'
  ).setTitle('Configuración').setWidth(340);

  SpreadsheetApp.getUi().showSidebar(html);
}

function guardarConfig(email, carpeta, apiKey) {
  const p = PropertiesService.getUserProperties();
  p.setProperty('ADMIN_EMAIL', email);
  p.setProperty('FOLDER_ID',   carpeta);
  if (apiKey && apiKey.trim()) p.setProperty('KOBO_API_KEY', apiKey.trim());
}

// ============================================================================
// PROBAR KOBO
// ============================================================================

function probarKobo() {
  try {
    const key  = getKoboKey();
    const resp = UrlFetchApp.fetch(CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/',
      { headers: { Authorization: 'Token ' + key }, muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) {
      const n = JSON.parse(resp.getContentText()).deployment__submission_count || 0;
      SpreadsheetApp.getUi().alert('✅ Kobo conectado\n\n📊 ' + n + ' respuestas disponibles');
    } else {
      SpreadsheetApp.getUi().alert('❌ Error ' + resp.getResponseCode() + '\n\nVerifica tu API Key en ⚙️ Configuración.');
    }
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e.message); }
}

// ============================================================================
// RESTAURAR DESDE DRIVE
// ============================================================================

function restaurarDesdeDrive() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss   = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) { ui.alert('❌ No existe el Maestro. Instala primero.'); return; }

    const existentes = new Set();
    if (hoja.getLastRow() > 1) {
      hoja.getRange(2, CONFIG.COL.ID, hoja.getLastRow()-1, 1)
        .getValues().forEach(r => { if (r[0]) existentes.add(String(r[0])); });
    }

    const folderBase = DriveApp.getFolderById(getFolderId());
    const subs = folderBase.getFolders();
    let restaurados = 0;

    while (subs.hasNext()) {
      const carp = subs.next();
      const partes = carp.getName().split(' — ');
      if (partes.length < 2) continue;
      const id     = partes[0].trim();
      const nombre = partes.slice(1).join(' — ').trim();
      if (existentes.has(id)) continue;

      const archs = carp.getFilesByName('Perfil — ' + nombre);
      const docId = archs.hasNext() ? archs.next().getId() : '';
      const docUrl = docId ? 'https://docs.google.com/document/d/' + docId + '/edit' : '';

      const row = new Array(26).fill('');
      row[CONFIG.COL.ID-1]         = id;
      row[CONFIG.COL.FECHA-1]      = new Date();
      row[CONFIG.COL.NOMBRE-1]     = nombre;
      row[CONFIG.COL.ESTADO-1]     = 'Recuperado';
      row[CONFIG.COL.CARPETA_ID-1] = carp.getId();
      row[CONFIG.COL.DOC_ID-1]     = docId;
      row[CONFIG.COL.DOC_URL-1]    = docUrl;
      hoja.appendRow(row);
      restaurados++;
    }

    ui.alert('✅ Restaurados: ' + restaurados + ' participantes desde Drive.\n\nNota: solo se recuperan ID, Nombre y links. Sincroniza con Kobo para completar datos.');
  } catch(e) { ui.alert('❌ Error: ' + e); }
}

// ============================================================================
// DESINSTALAR & LIMPIAR
// ============================================================================

function desinstalarYLimpiar() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('⚠️ DESINSTALAR — ¿Eliminar TODO?',
    'Se perderán hojas, expedientes en Drive, triggers y configuración.',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  if (ui.alert('⚠️ ÚLTIMA CONFIRMACIÓN',
    '¿Seguro? Esta acción es IRREVERSIBLE.',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  let eliminados = 0, errores = 0;

  // Eliminar expedientes de Drive
  const maestro = ss.getSheetByName(CONFIG.HOJA);
  if (maestro && maestro.getLastRow() > 1) {
    const M = CONFIG.COL;
    maestro.getRange(2, 1, maestro.getLastRow()-1, 26).getValues().forEach(fila => {
      const cId = String(fila[M.CARPETA_ID-1]||'').trim();
      const dId = String(fila[M.DOC_ID-1]    ||'').trim();
      if (cId) {
        try { DriveApp.getFolderById(cId).setTrashed(true); eliminados++; }
        catch(e) { if (dId) try { DriveApp.getFileById(dId).setTrashed(true); } catch(e2) { errores++; } }
      } else if (dId) {
        try { DriveApp.getFileById(dId).setTrashed(true); eliminados++; } catch(e) { errores++; }
      }
    });
  }

  // Crear hoja temporal (Sheets necesita mínimo 1 hoja)
  let tmp;
  try { tmp = ss.insertSheet('_temp_'); } catch(e) {}

  // Eliminar todas las hojas del sistema
  ['Maestro','Derivaciones','Log','Dashboard','Analytics'].forEach(n => {
    const h = ss.getSheetByName(n);
    if (h) try { ss.deleteSheet(h); } catch(e) {}
  });

  if (tmp) try { tmp.setName('Hoja1'); } catch(e) {}

  // Eliminar triggers y configuración
  ScriptApp.getProjectTriggers().forEach(t => { try { ScriptApp.deleteTrigger(t); } catch(e) {} });
  PropertiesService.getUserProperties().deleteAllProperties();

  ui.alert(
    '✅ Sistema eliminado\n\n' +
    '📁 ' + eliminados + ' expediente(s) a papelera' + (errores > 0 ? ' (' + errores + ' errores)' : '') + '\n' +
    '📋 Hojas eliminadas\n' +
    '⏱️ Triggers eliminados\n\n' +
    'Para empezar de nuevo: 📥 Instalar Sistema'
  );
}

// ============================================================================
// FIN — v8.7
// ============================================================================
