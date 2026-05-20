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
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🚀 Paso a Paso')
      // ── Acceso principal ──────────────────────────────────────────────
      .addItem('📋 Ver Participantes',               'verFicha')
      .addSeparator()
      // ── Datos ─────────────────────────────────────────────────────────
      .addItem('🔄 Importar Todo',                   'sincronizarTodo')
      .addItem('📊 Actualizar Power BI',              'exportarParaPowerBI')
      .addSeparator()
      // ── Automatización ────────────────────────────────────────────────
      .addItem('⏰ Activar Auto-actualización',       'configurarAutoActualizacion')
      .addSeparator()
      // ── Configuración & Setup ─────────────────────────────────────────
      .addItem('📥 Instalar / Reparar Sistema',      'reinstalarCompleto')
      .addItem('⚙️ Configuración',                   'abrirConfiguracion')
      .addSeparator()
      // ── Zona de peligro ───────────────────────────────────────────────
      .addItem('🗑️ Desinstalar & Limpiar',           'desinstalarYLimpiar')
      .addToUi();
  } catch(e) { /* contexto sin UI (trigger automático) — ignorar */ }
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

// ▶ Corre esta función desde el Editor de Apps Script para probar la carga
function testCargarParticipantes() {
  try {
    const lista = obtenerParticipantes();
    SpreadsheetApp.getUi().alert(
      '✅ obtenerParticipantes() funciona\n\n' +
      'Total: ' + lista.length + ' participantes\n' +
      (lista.length > 0
        ? 'Primero: ' + lista[0].nombre + ' — ' + lista[0].estado
        : '⚠️ La hoja Maestro está vacía o sin IDs')
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ ERROR en obtenerParticipantes:\n\n' + e.message + '\n\n' + e.stack);
  }
}

// ============================================================================
// INSTALAR
// ============================================================================

function instalar() {
  reinstalarCompleto();
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
// IMPORTAR DESDE DP_EMPLEABILIDAD → hoja "Derivados" (flujo de aprobación)
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

// ──────────────────────────────────────────────────────────────────────────
// NUEVO FLUJO: DP_Empleabilidad → hoja "Derivados" (pendiente formulario)
// ──────────────────────────────────────────────────────────────────────────

const COL_DER = { ID:1, FECHA:2, NOMBRE:3, DPI:4, EDAD:5, GENERO:6, TELEFONO:7,
                  EDUCACION:8, FORMACION:9, COHORTE:10, NOTA:11,
                  ESTADO:12, ORIGEN:13, FECHA_IMP:14 };
const NCOLS_DER = 14;

/** Importa DP_Empleabilidad → hoja "Derivados" (sin tocar Maestro) */
function importarDPADerivados(silencioso) {
  try {
    const ss = SpreadsheetApp.getActive();
    let hDer = ss.getSheetByName('Derivados');
    if (!hDer) {
      hDer = ss.insertSheet('Derivados');
      const hdrs = ['ID','Fecha Orig.','Nombre','DPI','Edad','Género','Teléfono',
                    'Educación','Formación','Cohorte','Notas','Estado Derivado','Origen','Fecha Importación'];
      hDer.getRange(1,1,1,hdrs.length).setValues([hdrs])
        .setFontWeight('bold').setBackground('#37474f').setFontColor('#fff');
      hDer.setFrozenRows(1);
      [200,110,200,110,60,90,100,130,150,100,200,130,120,130].forEach((w,i) => hDer.setColumnWidth(i+1,w));
    }

    let ssExt;
    try { ssExt = SpreadsheetApp.openById(DP_EMPLEABILIDAD.SPREADSHEET_ID); }
    catch(e) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ No se puede abrir DP_Empleabilidad.');
      return;
    }
    const hSrc = ssExt.getSheetByName(DP_EMPLEABILIDAD.HOJA);
    if (!hSrc || hSrc.getLastRow() < 2) {
      if (!silencioso) SpreadsheetApp.getUi().alert('⚠️ La hoja fuente está vacía.');
      return;
    }

    // IDs ya en Derivados Y ya en Maestro → evitar duplicados
    const idsExist = new Set();
    const maestro  = ss.getSheetByName(CONFIG.HOJA);
    if (maestro && maestro.getLastRow() > 1)
      maestro.getRange(2,CONFIG.COL.ID,maestro.getLastRow()-1,1).getValues()
        .forEach(r => { if(r[0]) idsExist.add(String(r[0]).trim()); });
    if (hDer.getLastRow() > 1)
      hDer.getRange(2,COL_DER.ID,hDer.getLastRow()-1,1).getValues()
        .forEach(r => { if(r[0]) idsExist.add(String(r[0]).trim()); });

    const C   = DP_EMPLEABILIDAD.C;
    const hoy = new Date();
    const datos = hSrc.getRange(2,1,hSrc.getLastRow()-1,DP_EMPLEABILIDAD.NCOLS).getValues();
    const nuevas = [];

    datos.forEach(fila => {
      const nombre = String(fila[C.NOMBRE]||'').trim();
      if (!nombre) return;
      let id = String(fila[C.ID]||'').trim();
      if (!id) id = nombre.replace(/\s+/g,'').substring(0,4).toUpperCase()
                  + Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'ddMMyyyy');
      if (idsExist.has(id)) return;
      idsExist.add(id);
      const activoStr = String(fila[C.ACTIVO]||'').toLowerCase().trim();
      const activo    = ['true','si','sí','1','activo','yes'].includes(activoStr);
      nuevas.push([
        id,
        fila[C.FECHA] instanceof Date ? fila[C.FECHA] : hoy,
        nombre,
        String(fila[C.DPI]     ||'').trim(),
        fila[C.EDAD]   || '',
        String(fila[C.GENERO]  ||'').trim(),
        String(fila[C.TELEFONO]||'').trim(),
        String(fila[C.EDUCACION]||'').trim(),
        String(fila[C.FORMACION]||'').trim(),
        String(fila[C.COHORTE] ||'').trim(),
        String(fila[C.NOTA]    ||'').trim(),
        activo ? 'Pendiente formulario' : 'Inactivo',
        'DP_Empleabilidad',
        hoy
      ]);
    });

    if (nuevas.length) {
      hDer.getRange(hDer.getLastRow()+1, 1, nuevas.length, NCOLS_DER).setValues(nuevas);
      // Colorear pendientes
      const startRow = hDer.getLastRow() - nuevas.length + 1;
      hDer.getRange(startRow, 1, nuevas.length, NCOLS_DER).setBackground('#fff8e1');
      SpreadsheetApp.flush();
    }

    if (!silencioso)
      SpreadsheetApp.getUi().alert(
        '✅ Importación a Derivados completada\n\n'+
        '📋 '+nuevas.length+' personas nuevas en hoja "Derivados"\n'+
        '⏭️ '+(datos.length - nuevas.length)+' ya existían (omitidas)\n\n'+
        '👉 Siguiente paso:\n'+
        'Abre "Gestionar Derivados" para ver el listado\n'+
        'y enviarles el formulario Paso a Paso.'
      );
  } catch(e) {
    logError('importarDPADerivados', e);
    if (!silencioso) SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/** Abre la hoja Derivados directamente */
function gestionarDerivados() {
  const ss  = SpreadsheetApp.getActive();
  let hDer  = ss.getSheetByName('Derivados');
  if (!hDer) {
    importarDPADerivados(false);
    hDer = ss.getSheetByName('Derivados');
    if (!hDer) return;
  }
  ss.setActiveSheet(hDer);
  SpreadsheetApp.getUi().alert(
    '📋 HOJA DERIVADOS\n\n'+
    'Aquí están las personas importadas de DP_Empleabilidad.\n\n'+
    '📌 Columna "Estado Derivado":\n'+
    '  • Pendiente formulario → aún no han llenado el Paso a Paso\n'+
    '  • Contactado → ya se les envió el formulario\n'+
    '  • Promovido → ya están en Maestro\n\n'+
    '📲 Acción recomendada:\n'+
    'Selecciona una persona y usa el menú\n'+
    '"🔄 Datos → Enviar WhatsApp formulario"\n'+
    'para compartirles el link del Kobo.'
  );
}

/** Envía WhatsApp con link del formulario Kobo a un Derivado seleccionado */
function enviarFormularioADerivado() {
  try {
    const ss   = SpreadsheetApp.getActive();
    const hDer = ss.getSheetByName('Derivados');
    if (!hDer) { SpreadsheetApp.getUi().alert('⚠️ No hay hoja Derivados.'); return; }
    const fila = ss.getActiveRange().getRow();
    if (fila < 2) { SpreadsheetApp.getUi().alert('⚠️ Selecciona primero una persona en la hoja Derivados.'); return; }
    const row  = hDer.getRange(fila, 1, 1, NCOLS_DER).getValues()[0];
    const nom  = String(row[COL_DER.NOMBRE-1]||'');
    const tel  = String(row[COL_DER.TELEFONO-1]||'').replace(/\D/g,'');
    if (!tel) { SpreadsheetApp.getUi().alert('⚠️ '+nom+' no tiene teléfono registrado.'); return; }
    const numWA = tel.length===8 ? '502'+tel : tel;
    const koboUrl = 'https://kf.kobotoolbox.org/x/'+CONFIG.KOBO_ASSET_ID;
    const msg = encodeURIComponent(
      'Hola '+nom.split(' ')[0]+', somos el equipo de Paso a Paso de Creamos Guatemala.\n\n'+
      'Te invitamos a completar nuestro formulario de registro para poder iniciar tu proceso:\n'+
      koboUrl+'\n\n'+
      '¿Tienes alguna duda? Estamos para ayudarte. 😊'
    );
    const waUrl = 'https://wa.me/'+numWA+'?text='+msg;
    // Actualizar estado en la hoja
    hDer.getRange(fila, COL_DER.ESTADO).setValue('Contactado');
    hDer.getRange(fila, 1, 1, NCOLS_DER).setBackground('#e8f5e9');
    // Abrir WhatsApp en nueva pestaña via dialog
    const htmlLink = HtmlService.createHtmlOutput(
      '<script>window.open("'+waUrl+'","_blank");google.script.host.close();</script>'
    ).setWidth(1).setHeight(1);
    SpreadsheetApp.getUi().showModalDialog(htmlLink, 'Abriendo WhatsApp…');
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e.message); }
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

      const nombre = r['nombre_completo'] || 'Participante ' + id;
      const fila   = new Array(26).fill('');
      fila[M.ID-1]         = id;
      fila[M.FECHA-1]      = new Date();
      fila[M.NOMBRE-1]     = nombre;
      fila[M.DPI-1]        = r['numero_dpi']                      || '';
      fila[M.EDAD-1]       = r['edad']                            || '';
      fila[M.GENERO-1]     = r['genero']                          || '';
      fila[M.TELEFONO-1]   = r['telefono']                        || '';
      fila[M.ZONA-1]       = r['lugar_residencia']                || '';
      fila[M.EMAIL-1]      = r['email']                           || '';
      fila[M.EDUCACION-1]  = r['ultimo_grado_aprobado']           || '';
      fila[M.LABORAL-1]    = r['situacion_laboral_actual']        || '';
      fila[M.FORTALEZAS-1] = r['habilidades_especificas']         || '';
      fila[M.OBJETIVO-1]   = r['objetivo_laboral_especifico']     || '';
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

    // ── Acciones en hoja Derivados ────────────────────────────────────────
    if (sheet.getName() === 'Derivados') {
      const fila = e.range.getRow();
      const col  = e.range.getColumn();
      const valN = e.value || '';

      // Verificar si es la columna de Acción buscando por header
      const headers = sheet.getRange(1, 1, 1, 100).getValues()[0];
      const accionCol = headers.findIndex(h => String(h).includes('Acción'));
      if (fila < 2 || col !== accionCol + 1 || !valN) return;
      e.range.clearContent();

      const rowData = sheet.getRange(fila, 1, 1, 100).getValues()[0];
      const nombreIdx = headers.findIndex(h => String(h).includes('Nombre'));
      const nombre = String(rowData[nombreIdx] || '');
      const KOBO_FORM = 'https://ee.kobotoolbox.org/x/M9M734If';

      if (valN.includes('Enviar') || valN.includes('recordatorio')) {
        // Dialog para mostrar el link (presencial, no WhatsApp)
        const html = HtmlService.createHtmlOutput(
          '<div style="font-family:\'Segoe UI\',sans-serif;padding:24px;text-align:center;max-width:380px;">' +
          '<div style="font-size:40px;margin-bottom:12px;">📋</div>' +
          '<h2 style="font-size:16px;color:#1a237e;margin:0 0 6px 0;">Formulario Kobo para</h2>' +
          '<p style="font-size:15px;font-weight:700;color:#212121;margin:0 0 16px 0;">' + nombre + '</p>' +
          '<div style="background:#f5f5f5;border-left:4px solid #1a237e;padding:12px;border-radius:4px;margin-bottom:16px;text-align:left;">' +
          '<p style="margin:0 0 8px 0;font-size:11px;color:#666;font-weight:600;">LINK DEL FORMULARIO:</p>' +
          '<p style="margin:0;font-family:monospace;font-size:10px;word-break:break-all;color:#1a237e;padding:8px;background:#fff;border-radius:3px;border:1px solid #ddd;">' + KOBO_FORM + '</p>' +
          '</div>' +
          '<p style="font-size:11px;color:#666;margin:0 0 14px 0;line-height:1.5;"><strong>Instrucciones:</strong><br>Comparte este link con ' + nombre + ' para que llene el formulario. Puedes:<br>• Copiar y enviar por WhatsApp/SMS<br>• Mostrar en tu teléfono presencialmente</p>' +
          '<div style="display:flex;gap:8px;">' +
          '<button onclick="navigator.clipboard.writeText(\'' + KOBO_FORM + '\').then(()=>alert(\'✅ Link copiado\'));" style="flex:1;padding:10px 16px;background:#1a237e;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:700;font-size:12px;">📋 Copiar link</button>' +
          '<a href="' + KOBO_FORM + '" target="_blank" style="flex:1;padding:10px 16px;background:#25d366;color:#fff;border:none;border-radius:6px;text-decoration:none;font-weight:700;font-size:12px;display:flex;align-items:center;justify-content:center;">🌐 Abrir</a>' +
          '</div>' +
          '<p style="font-size:9px;color:#aaa;margin:12px 0 0 0;">Cierra cuando hayas terminado</p>' +
          '</div>'
        ).setWidth(360).setHeight(420).setTitle('Formulario Kobo — ' + nombre);
        SpreadsheetApp.getUi().showModelessDialog(html, 'Formulario Kobo');
        // Marcar como enviado - buscar columna Estado
        const estadoIdx = headers.findIndex(h => String(h).includes('Estado'));
        if (estadoIdx >= 0) sheet.getRange(fila, estadoIdx + 1).setValue('Formulario enviado');

      } else if (valN.includes('completó')) {
        // Marcar como completado
        const estadoIdx = headers.findIndex(h => String(h).includes('Estado'));
        const fechaIdx = headers.findIndex(h => String(h).includes('Fecha Aprobación'));
        if (estadoIdx >= 0) sheet.getRange(fila, estadoIdx + 1).setValue('Completó formulario');
        if (fechaIdx >= 0) sheet.getRange(fila, fechaIdx + 1).setValue(new Date());
        if (estadoIdx >= 0) sheet.getRange(fila, estadoIdx + 1, 1, Math.max(1, Math.abs(fechaIdx - estadoIdx) + 1)).setBackground('#d9ead3');
        SpreadsheetApp.getUi().alert('✅ Marcado como completado.\n\n' + nombre + ' puede proceder con el siguiente paso del programa.');
      }
      return;
    }

    if (sheet.getName() !== CONFIG.HOJA) return;

    const fila = e.range.getRow();
    const col  = e.range.getColumn();
    const valN = e.value    || '';
    const valA = e.oldValue || '';
    if (fila < 2 || !valN || valN === valA) return;

    // ── Acción Rápida (columna 27) ──────────────────────────────────────────
    if (col === 27) {
      e.range.clearContent();
      const rd = sheet.getRange(fila, 1, 1, 26).getValues()[0];
      const nombreP = String(rd[2] || '');
      const telRaw  = String(rd[6] || '').replace(/\D/g,'');
      const tel     = telRaw.length === 8 ? '502'+telRaw : telRaw;
      const docUrl  = String(rd[25]|| '');
      const idP     = String(rd[0] || '');
      if (valN.includes('WhatsApp')) {
        if (!tel) { SpreadsheetApp.getUi().alert('⚠️ Sin teléfono\n\n'+nombreP+' no tiene teléfono registrado.'); return; }
        const msg = encodeURIComponent('Hola '+nombreP+', somos el equipo de Paso a Paso de Creamos Guatemala. ¿Cómo estás?');
        const html = HtmlService.createHtmlOutput(
          '<div style="font-family:\'Segoe UI\',sans-serif;padding:22px;text-align:center">'+
          '<p style="font-size:14px;color:#333;margin-bottom:6px">Enviar mensaje a</p>'+
          '<p style="font-size:16px;font-weight:700;color:#1a237e;margin-bottom:18px">'+nombreP+'</p>'+
          '<a href="https://wa.me/'+tel+'?text='+msg+'" target="_blank"'+
          ' style="display:block;background:#25d366;color:#fff;padding:13px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">💬 Abrir WhatsApp</a>'+
          '</div>'
        ).setWidth(300).setHeight(170);
        SpreadsheetApp.getUi().showModelessDialog(html, '💬 WhatsApp — '+nombreP.split(' ')[0]);
      } else if (valN.includes('Ver Perfil')) {
        PropertiesService.getScriptProperties().setProperty('FICHA_OPEN_ID', idP);
        verFicha();
      } else if (valN.includes('Sesión')) {
        abrirSesionDesdeHoja(idP, nombreP);
      }
      return;
    }
    // ────────────────────────────────────────────────────────────────────────

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
// DASHBOARD (in-sheet con gráficos embebidos)
// ============================================================================

function _OLD_mostrarDashboardGrafico() {
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) { SpreadsheetApp.getUi().alert('❌ Instala primero.'); return; }

    const M = CONFIG.COL;
    const tz = Session.getScriptTimeZone();
    const datos = hoja.getLastRow() > 1
      ? hoja.getRange(2, 1, hoja.getLastRow()-1, 26).getValues().filter(r => r[0])
      : [];

    if (datos.length === 0) {
      SpreadsheetApp.getUi().alert('⚠️ No hay datos para mostrar');
      return;
    }

    // Preparar datos para gráficos
    const perfiles = {'Perfil A':0,'Perfil B':0,'Perfil C':0,'Perfil D':0,'Sin perfil':0};
    const estados = {};
    const dims = [[],[],[],[],[],[]];

    datos.forEach(r => {
      const perfil = normalizarPerfil(r[M.PERFIL-1]);
      perfiles[perfil] = (perfiles[perfil]||0)+1;

      const estado = r[M.ESTADO-1]||'Sin estado';
      estados[estado] = (estados[estado]||0)+1;

      for (let i=0; i<6; i++) {
        const val = Number(r[M.DIM1+i-1])||0;
        dims[i].push(val);
      }
    });

    // Calcular promedios por dimensión
    const dimPromedio = dims.map(d => d.length > 0 ? Math.round(d.reduce((a,b)=>a+b,0)/d.length) : 0);

    // Construir datos JSON
    const datosGrafico = {
      perfiles: perfiles,
      estados: estados,
      dimPromedio: dimPromedio,
      total: datos.length,
      timestamp: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm')
    };

    const html = crearHTMLDashboardGrafico(datosGrafico);
    const modal = HtmlService.createHtmlOutput(html)
      .setWidth(1200)
      .setHeight(900);
    SpreadsheetApp.getUi().showModelessDialog(modal, '📊 Dashboard — Paso a Paso');

  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('mostrarDashboardGrafico', e);
  }
}

function crearHTMLDashboardGrafico(datos) {
  const perfiles = datos.perfiles;
  const estados = datos.estados;
  const dimPromedio = datos.dimPromedio;

  const estdoArray = Object.entries(estados).sort((a,b)=>b[1]-a[1]);

  return `
<!DOCTYPE html>
<html>
<head>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@3.9.1/dist/chart.min.js"></script>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana;
      margin: 0;
      padding: 20px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
    }
    .container {
      max-width: 1400px;
      margin: 0 auto;
    }
    .header {
      color: white;
      margin-bottom: 25px;
      text-align: center;
    }
    .header h1 { margin: 0; font-size: 28px; }
    .header p { margin: 5px 0 0 0; opacity: 0.9; }
    .grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 20px;
      margin-bottom: 20px;
    }
    .chart-box {
      background: white;
      border-radius: 12px;
      padding: 20px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.1);
    }
    .chart-box h3 {
      margin: 0 0 15px 0;
      color: #333;
      font-size: 16px;
      border-bottom: 3px solid #667eea;
      padding-bottom: 10px;
    }
    .stats-row {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 15px;
      margin-bottom: 20px;
    }
    .stat-card {
      background: white;
      border-radius: 12px;
      padding: 15px;
      text-align: center;
      box-shadow: 0 4px 20px rgba(0,0,0,0.1);
    }
    .stat-card .number {
      font-size: 32px;
      font-weight: bold;
      color: #667eea;
      margin: 10px 0;
    }
    .stat-card .label {
      font-size: 12px;
      color: #999;
      text-transform: uppercase;
      letter-spacing: 1px;
    }
    .chart-canvas {
      position: relative;
      height: 300px;
      margin: 0 auto;
    }
    .dim-chart {
      grid-column: 1 / -1;
    }
    .footer {
      color: white;
      text-align: center;
      font-size: 12px;
      opacity: 0.8;
      margin-top: 20px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📊 Dashboard — Paso a Paso</h1>
      <p>Actualizado: ${datos.timestamp}</p>
    </div>

    <div class="stats-row">
      <div class="stat-card">
        <div class="label">Total</div>
        <div class="number">${datos.total}</div>
        <div class="label">Participantes</div>
      </div>
      <div class="stat-card">
        <div class="label">Perfil A</div>
        <div class="number">${perfiles['Perfil A']}</div>
        <div class="label">Listos para empleo</div>
      </div>
      <div class="stat-card">
        <div class="label">Perfil B</div>
        <div class="number">${perfiles['Perfil B']}</div>
        <div class="label">En orientación</div>
      </div>
      <div class="stat-card">
        <div class="label">Prom. Puntaje</div>
        <div class="number">${Math.round(dimPromedio.reduce((a,b)=>a+b,0)/dimPromedio.length)}</div>
        <div class="label">/ 60</div>
      </div>
    </div>

    <div class="grid">
      <div class="chart-box">
        <h3>🎯 Distribución por Perfil</h3>
        <div class="chart-canvas">
          <canvas id="perfilChart"></canvas>
        </div>
      </div>

      <div class="chart-box">
        <h3>📌 Participantes por Estado</h3>
        <div class="chart-canvas">
          <canvas id="estadoChart"></canvas>
        </div>
      </div>
    </div>

    <div class="chart-box dim-chart">
      <h3>📈 Promedio por Dimensión</h3>
      <div class="chart-canvas" style="height: 250px;">
        <canvas id="dimChart"></canvas>
      </div>
    </div>

    <div class="footer">
      Pulse ESC o haga clic fuera para cerrar
    </div>
  </div>

  <script>
    const chartOptions = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { usePointStyle: true, padding: 15 } } }
    };

    // Gráfico de Perfiles (Pie)
    new Chart(document.getElementById('perfilChart'), {
      type: 'doughnut',
      data: {
        labels: ['🟢 Perfil A', '🔵 Perfil B', '🟡 Perfil C', '🔴 Perfil D', '⬜ Sin perfil'],
        datasets: [{
          data: [${perfiles['Perfil A']}, ${perfiles['Perfil B']}, ${perfiles['Perfil C']}, ${perfiles['Perfil D']}, ${perfiles['Sin perfil']}],
          backgroundColor: ['#43a047', '#1e88e5', '#fb8c00', '#e53935', '#bdbdbd'],
          borderColor: 'white',
          borderWidth: 2
        }]
      },
      options: {
        ...chartOptions,
        plugins: { ...chartOptions.plugins, tooltip: { callbacks: { label: ctx => ctx.label + ': ' + ctx.parsed } } }
      }
    });

    // Gráfico de Estados (Bar)
    new Chart(document.getElementById('estadoChart'), {
      type: 'bar',
      data: {
        labels: ${JSON.stringify(estdoArray.map(e=>e[0]))},
        datasets: [{
          label: 'Cantidad',
          data: ${JSON.stringify(estdoArray.map(e=>e[1]))},
          backgroundColor: '#667eea',
          borderColor: '#667eea',
          borderWidth: 1,
          borderRadius: 6
        }]
      },
      options: {
        ...chartOptions,
        indexAxis: 'y',
        scales: { x: { beginAtZero: true, ticks: { stepSize: 1 } } }
      }
    });

    // Gráfico de Dimensiones (Line)
    new Chart(document.getElementById('dimChart'), {
      type: 'line',
      data: {
        labels: ['Capital Educativo', 'Capital Laboral', 'Habilidades Digitales', 'Claridad Vocacional', 'Barreras', 'Red de Apoyo'],
        datasets: [{
          label: 'Promedio',
          data: ${JSON.stringify(dimPromedio)},
          borderColor: '#667eea',
          backgroundColor: 'rgba(102, 126, 234, 0.1)',
          borderWidth: 3,
          pointRadius: 6,
          pointBackgroundColor: '#667eea',
          pointBorderColor: 'white',
          pointBorderWidth: 2,
          fill: true,
          tension: 0.4
        }]
      },
      options: {
        ...chartOptions,
        scales: { y: { beginAtZero: true, max: 60, ticks: { stepSize: 10 } } }
      }
    });
  </script>
</body>
</html>
  `;
}

function abrirDashboard() {
  try {
    const ss = SpreadsheetApp.getActive();
    actualizarDashboard(ss);
    ss.setActiveSheet(ss.getSheetByName('Dashboard'));
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e); }
}

function actualizarDashboard(ss) {
  try {
    if (!ss) ss = SpreadsheetApp.getActive();
    const maestro = ss.getSheetByName(CONFIG.HOJA);
    if (!maestro) return;

    let dash = ss.getSheetByName('Dashboard');
    if (!dash) dash = ss.insertSheet('Dashboard');

    dash.clear();
    dash.clearFormats();
    dash.getCharts().forEach(c => dash.removeChart(c));

    const M  = CONFIG.COL;
    const tz = Session.getScriptTimeZone();
    const ahora = new Date();
    const datos = maestro.getLastRow() > 1
      ? maestro.getRange(2,1,maestro.getLastRow()-1,26).getValues().filter(r => r[M.ID-1])
      : [];

    // ── Calcular métricas ──
    const perfiles = {'Perfil A':0,'Perfil B':0,'Perfil C':0,'Perfil D':0,'Sin perfil':0};
    const estados  = {};
    const generos  = {};
    const dims     = [[],[],[],[],[],[]];
    const puntajes = [];
    const mensuales = {};
    let enAcomp = 0, conexiones = 0, ingresadosMes = 0;
    const mesAct = ahora.getMonth(), anioAct = ahora.getFullYear();

    datos.forEach(r => {
      const p = normalizarPerfil(r[M.PERFIL-1]);
      perfiles[p] = (perfiles[p]||0)+1;
      const e = r[M.ESTADO-1]||'Sin estado';
      estados[e] = (estados[e]||0)+1;
      if (['Mentoría','Orientación'].includes(e)) enAcomp++;
      if (['Cierre','Completado'].includes(e)) conexiones++;
      const gen = String(r[M.GENERO-1]||'Sin dato');
      generos[gen] = (generos[gen]||0)+1;
      const pt = Number(r[M.PUNTAJE-1]);
      if (pt > 0) puntajes.push(pt);
      for (let i=0;i<6;i++) dims[i].push(Number(r[M.DIM1+i-1])||0);
      const f = r[M.FECHA-1];
      if (f instanceof Date) {
        const k = Utilities.formatDate(f,tz,'yyyy-MM');
        mensuales[k] = (mensuales[k]||0)+1;
        if (f.getMonth()===mesAct && f.getFullYear()===anioAct) ingresadosMes++;
      }
    });

    const total     = datos.length;
    const promPt    = puntajes.length ? Math.round(puntajes.reduce((a,b)=>a+b,0)/puntajes.length) : 0;
    const dimProm   = dims.map(d => d.length ? Math.round(d.reduce((a,b)=>a+b,0)/d.length*10)/10 : 0);
    const estList   = Object.entries(estados).sort((a,b)=>b[1]-a[1]);
    const genList   = Object.entries(generos).sort((a,b)=>b[1]-a[1]);
    const tasaGrad  = total > 0 ? (conexiones/total*100).toFixed(1)+'%' : '0%';

    // ── Helper functions ──
    const W = 6;
    const H = (r,h) => dash.setRowHeight(r,h);
    const titSec = (row, txt, bg) => {
      H(row,30);
      dash.getRange(row,1,1,W).merge().setValue(txt)
        .setFontSize(11).setFontWeight('bold').setFontColor('#fff')
        .setBackground(bg||'#3949ab').setHorizontalAlignment('center').setVerticalAlignment('middle');
    };
    const subHdr = (row, cols) => {
      H(row,22);
      cols.forEach((c,i) => dash.getRange(row,i+1).setValue(c)
        .setFontWeight('bold').setBackground('#e8eaf6').setFontColor('#333')
        .setHorizontalAlignment(i>0?'center':'left').setFontSize(10));
    };
    const dataRow = (row, cells, bgs, fgs) => {
      H(row,22);
      cells.forEach((v,i) => {
        const r = dash.getRange(row,i+1).setValue(v).setFontSize(10)
          .setHorizontalAlignment(i>0?'center':'left');
        if (bgs&&bgs[i]) r.setBackground(bgs[i]);
        if (fgs&&fgs[i]) r.setFontColor(fgs[i]).setFontWeight('bold');
      });
    };
    const totRow = (row, cells) => {
      H(row,22);
      cells.forEach((v,i) => dash.getRange(row,i+1).setValue(v)
        .setFontWeight('bold').setBackground('#e8eaf6').setFontColor('#1a237e')
        .setHorizontalAlignment(i>0?'center':'left').setFontSize(10));
    };
    const spacer = (row) => { H(row,14); };

    // ── COLUMNAS ──
    dash.setColumnWidth(1,200);
    dash.setColumnWidth(2,90);
    dash.setColumnWidth(3,90);
    dash.setColumnWidth(4,90);
    dash.setColumnWidth(5,90);
    dash.setColumnWidth(6,90);

    let r = 1;

    // ════ ENCABEZADO ════
    H(r,48);
    dash.getRange(r,1,1,W).merge()
      .setValue('📊  REPORTE DE SEGUIMIENTO — INCLUSIÓN LABORAL')
      .setFontSize(17).setFontWeight('bold').setFontColor('#fff')
      .setBackground('#1a237e').setHorizontalAlignment('center').setVerticalAlignment('middle');
    r++;
    H(r,26);
    dash.getRange(r,1,1,W).merge()
      .setValue('Actualizado: '+Utilities.formatDate(ahora,tz,'d/M/yyyy  HH:mm'))
      .setFontSize(10).setFontColor('#c5cae9').setBackground('#283593')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    r++; spacer(r); r++;

    // ════ RESUMEN GENERAL ════
    titSec(r,'📋  RESUMEN GENERAL','#1565c0'); r++;
    const kpis = [
      {lbl:'Registrados (IL.P.01)',  val:total,          bg:'#e8eaf6', fg:'#1a237e'},
      {lbl:'Acompañamiento (IL.P.02)',val:enAcomp,        bg:'#e8f5e9', fg:'#1b5e20'},
      {lbl:'Conexiones lab. (IL.R.07)',val:conexiones,    bg:'#e3f2fd', fg:'#01579b'},
      {lbl:'Puntaje prom.',           val:promPt+'/60',   bg:'#fff8e1', fg:'#e65100'},
      {lbl:'Perfil A  ✓ Empleo',     val:perfiles['Perfil A'], bg:'#d9ead3', fg:'#274e13'},
      {lbl:'Este mes',                val:ingresadosMes,  bg:'#f3e5f5', fg:'#6a1b9a'},
    ];
    // 2 filas de KPI: label y valor
    H(r,22);
    kpis.forEach((k,i) => dash.getRange(r,i+1).setValue(k.lbl)
      .setFontSize(8).setFontWeight('bold').setFontColor('#666').setHorizontalAlignment('center').setBackground(k.bg));
    r++;
    H(r,42);
    kpis.forEach((k,i) => dash.getRange(r,i+1).setValue(k.val)
      .setFontSize(22).setFontWeight('bold').setFontColor(k.fg).setHorizontalAlignment('center')
      .setBackground(k.bg).setVerticalAlignment('middle'));
    r++;
    spacer(r); r++;

    // ════ DISTRIBUCIÓN POR ESTADO ════
    titSec(r,'📌  DISTRIBUCIÓN POR ESTADO / ETAPA','#1565c0'); r++;
    subHdr(r,['Estado / Etapa','Participantes','%','','','']); r++;
    estList.forEach(([est,cnt]) => {
      const pct = total>0 ? (cnt/total*100).toFixed(1)+'%' : '0%';
      dataRow(r,[est,cnt,pct,'','','']); r++;
    });
    totRow(r,['TOTAL',total,'100%','','','']); r++;
    spacer(r); r++;

    // ════ DISTRIBUCIÓN POR PERFIL ════
    titSec(r,'🎯  DISTRIBUCIÓN POR PERFIL (Diagnóstico IL)','#1565c0'); r++;
    subHdr(r,['Perfil','Participantes','%','Descripción','','']); r++;
    const pDesc = {
      'Perfil A':'Listo para empleo',
      'Perfil B':'Orientación vocacional',
      'Perfil C':'Desarrollo de capacidades',
      'Perfil D':'Barreras críticas (URGENTE)',
      'Sin perfil':'Sin asignar'
    };
    const pCol = {
      'Perfil A':{bg:'#d9ead3',fg:'#274e13'},
      'Perfil B':{bg:'#cfe2f3',fg:'#1c4587'},
      'Perfil C':{bg:'#fff2cc',fg:'#7f6000'},
      'Perfil D':{bg:'#f4cccc',fg:'#660000'},
      'Sin perfil':{bg:'#f5f5f5',fg:'#777'}
    };
    Object.entries(perfiles).forEach(([prf,cnt]) => {
      const c = pCol[prf]||{bg:'#f5f5f5',fg:'#777'};
      const pct = total>0 ? (cnt/total*100).toFixed(1)+'%' : '0%';
      H(r,22);
      dash.getRange(r,1).setValue(prf).setFontWeight('bold').setBackground(c.bg).setFontColor(c.fg).setFontSize(10);
      dash.getRange(r,2).setValue(cnt).setHorizontalAlignment('center').setBackground(c.bg).setFontColor(c.fg).setFontWeight('bold').setFontSize(10);
      dash.getRange(r,3).setValue(pct).setHorizontalAlignment('center').setBackground(c.bg).setFontColor(c.fg).setFontSize(10);
      dash.getRange(r,4,1,3).merge().setValue(pDesc[prf]||'').setBackground(c.bg).setFontColor(c.fg).setFontSize(10).setFontStyle('italic');
      r++;
    });
    totRow(r,['TOTAL',total,'100%','','','']); r++;
    spacer(r); r++;

    // ════ DESAGREGACIÓN POR GÉNERO (requerido por indicadores IL) ════
    titSec(r,'♀♂  DESAGREGACIÓN POR GÉNERO (requerido IL.P.01, IL.R.07)','#37474f'); r++;
    subHdr(r,['Género','Participantes','%','','','']); r++;
    genList.forEach(([gen,cnt]) => {
      const pct = total>0 ? (cnt/total*100).toFixed(1)+'%' : '0%';
      dataRow(r,[gen,cnt,pct,'','','']); r++;
    });
    totRow(r,['TOTAL',total,'100%','','','']); r++;
    spacer(r); r++;

    // ════ DIMENSIONES DIAGNÓSTICO ════
    titSec(r,'📈  PROMEDIO DIMENSIONES DE DIAGNÓSTICO (sobre 10)','#00695c'); r++;
    subHdr(r,['Dimensión','Promedio','Nivel','','','']); r++;
    const dimLabels = [
      '📚 Cap. Educativo (Dim 1)',
      '💼 Cap. Laboral (Dim 2)',
      '💻 Hab. Digitales (Dim 3)',
      '🎯 Claridad Vocacional (Dim 4)',
      '🚧 Barreras Estructurales (Dim 5)',
      '🤝 Red de Apoyo (Dim 6)'
    ];
    dimLabels.forEach((dim,i) => {
      const val = dimProm[i];
      const bg = val>=7?'#c8e6c9':val>=4?'#fff9c4':val>0?'#ffcdd2':'#f5f5f5';
      const fg = val>=7?'#1b5e20':val>=4?'#f57f17':val>0?'#c62828':'#999';
      const niv = val>=7?'Alto':val>=4?'Medio':val>0?'Bajo':'Sin datos';
      H(r,22);
      dash.getRange(r,1).setValue(dim).setFontSize(10);
      dash.getRange(r,2).setValue(val).setBackground(bg).setFontColor(fg).setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10);
      dash.getRange(r,3).setValue(niv).setBackground(bg).setFontColor(fg).setHorizontalAlignment('center').setFontSize(10);
      dash.getRange(r,4,1,3).merge().setValue('').setBackground('#fafafa');
      r++;
    });
    spacer(r); r++;

    // ════ INGRESOS POR MES (últimos 12) ════
    titSec(r,'📅  INGRESOS POR MES (últimos meses)','#01579b'); r++;
    subHdr(r,['Mes','Nuevos','Acumulado','','','']); r++;
    const mesKeys = Object.keys(mensuales).sort().slice(-12);
    let acum = total - mesKeys.reduce((a,k) => a+(mensuales[k]||0),0);
    mesKeys.forEach(k => {
      const cnt = mensuales[k]||0;
      acum += cnt;
      const label = Utilities.formatDate(new Date(k+'-01'),tz,'MMM yyyy');
      dataRow(r,[label,cnt,acum,'','','']); r++;
    });
    spacer(r); r++;

    // ════ NOTA POWER BI ════
    H(r,28);
    dash.getRange(r,1,1,W).merge()
      .setValue('⚡  Para actualizar Power BI: menú  📊 PASO A PASO → Exportar para Power BI')
      .setFontSize(9).setFontColor('#fff').setBackground('#455a64')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontStyle('italic');

  } catch(e) {
    logError('actualizarDashboard', e);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

// ── Generador de Reporte Mensual ──
function generarReporteMensual() {
  try {
    const ss = SpreadsheetApp.getActive();
    const maestro = ss.getSheetByName(CONFIG.HOJA);
    if (!maestro || maestro.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ No hay datos para generar el reporte.');
      return;
    }

    const M  = CONFIG.COL;
    const tz = Session.getScriptTimeZone();
    const ahora = new Date();
    const mesNombre = Utilities.formatDate(ahora, tz, 'MMMM yyyy');
    const mes = ahora.getMonth(), anio = ahora.getFullYear();

    const datos = maestro.getRange(2,1,maestro.getLastRow()-1,26).getValues().filter(r => r[M.ID-1]);
    const perfiles = {'Perfil A':0,'Perfil B':0,'Perfil C':0,'Perfil D':0};
    const estados  = {};
    const puntajes = [];
    let ingresadosMes = 0;

    datos.forEach(r => {
      const p = normalizarPerfil(r[M.PERFIL-1]);
      if (perfiles[p] !== undefined) perfiles[p]++;
      const e = r[M.ESTADO-1]||'Sin estado';
      estados[e] = (estados[e]||0)+1;
      const pt = Number(r[M.PUNTAJE-1]);
      if (pt > 0) puntajes.push(pt);
      const f = r[M.FECHA-1];
      if (f instanceof Date && f.getMonth()===mes && f.getFullYear()===anio) ingresadosMes++;
    });

    const total     = datos.length;
    const promPt    = puntajes.length ? Math.round(puntajes.reduce((a,b)=>a+b,0)/puntajes.length) : 0;
    const enCierre  = (estados['Cierre']||0)+(estados['Completado']||0);
    const activos   = total - (estados['Inactivo']||0) - enCierre;

    const resumenDatos =
      `• Total acumulado: ${total} participantes\n`+
      `• Nuevos este mes: ${ingresadosMes}\n`+
      `• Activos en proceso: ${activos}\n`+
      `• Perfil A (listos para empleo): ${perfiles['Perfil A']}\n`+
      `• Perfil B (orientación vocacional): ${perfiles['Perfil B']}\n`+
      `• Perfil C (desarrollo de capacidades): ${perfiles['Perfil C']}\n`+
      `• Perfil D (barreras críticas): ${perfiles['Perfil D']}\n`+
      `• En Mentoría: ${estados['Mentoría']||0}\n`+
      `• En Formación: ${estados['Formación']||0}\n`+
      `• Derivados: ${estados['Derivación']||0}\n`+
      `• Finalizados (Cierre/Completado): ${enCierre}\n`+
      `• Puntaje diagnóstico promedio: ${promPt} / 60`;

    const promptFinal =
      `Actúa como redactor de impacto social. Escribe el texto de un reporte mensual para el programa PASO A PASO de Creamos Guatemala.\n\n`+
      `DATOS DEL PERÍODO — ${mesNombre.toUpperCase()}:\n`+
      `${resumenDatos}\n\n`+
      `ESCRIBE ESTAS SECCIONES:\n\n`+
      `1. TITULAR DE IMPACTO (1 oración)\n`+
      `   Formato: "[X] personas [logro] gracias a [programa]"\n\n`+
      `2. RESUMEN EJECUTIVO (3 oraciones)\n`+
      `   - Qué se hizo\n`+
      `   - Cuál fue el resultado más importante\n`+
      `   - Qué sigue\n\n`+
      `3. LOGROS DEL MES (máximo 5 viñetas)\n`+
      `   Cada viñeta: verbo en pasado + dato concreto + impacto\n`+
      `   Ejemplo: "Acompañamos a 12 personas en búsqueda activa, logrando 4 nuevos empleos"\n\n`+
      `4. DESAFÍOS (máximo 3 viñetas)\n`+
      `   Tono honesto pero constructivo, no alarmista\n\n`+
      `5. PRÓXIMAS ACCIONES (3 puntos)\n`+
      `   Específicas, con responsable implícito y fecha si aplica\n\n`+
      `TONO: Profesional, cercano, orientado a personas (no a procesos). `+
      `No usar palabras como "sinergia", "robusto", "implementar" — usar lenguaje simple. Máximo 200 palabras en total.`;

    const html = HtmlService.createHtmlOutput(
      `<!DOCTYPE html><html><head><style>
        body{font-family:'Segoe UI',Arial;font-size:12px;margin:0;background:#f5f5f5}
        .hdr{background:#1a237e;color:#fff;padding:14px 18px}
        .hdr h2{margin:0;font-size:14px}
        .hdr p{margin:4px 0 0;font-size:10px;opacity:.8}
        .body{padding:16px}
        textarea{width:100%;height:520px;font-size:11px;font-family:monospace;border:1px solid #c5cae9;border-radius:6px;padding:12px;resize:none;background:#fff;line-height:1.6}
        .btn{display:block;width:100%;margin-top:10px;padding:10px;background:#1a237e;color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:700;cursor:pointer}
        .btn:hover{background:#283593}
        .note{font-size:10px;color:#999;margin-top:8px;text-align:center}
      </style></head><body>
        <div class="hdr"><h2>📝 Reporte Mensual — ${mesNombre}</h2>
          <p>Copia este prompt y pégalo en Claude o ChatGPT para generar el reporte</p></div>
        <div class="body">
          <textarea id="pt" readonly>${promptFinal.replace(/</g,'&lt;').replace(/>/g,'&gt;')}</textarea>
          <button class="btn" onclick="document.getElementById('pt').select();document.execCommand('copy');this.textContent='✅ ¡Copiado!'">📋 Copiar Prompt</button>
          <div class="note">Después de copiar, pégalo en tu herramienta de IA favorita</div>
        </div>
      </body></html>`
    ).setWidth(600).setHeight(700);

    SpreadsheetApp.getUi().showModalDialog(html, '📝 Reporte Mensual — ' + mesNombre);

  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('generarReporteMensual', e);
  }
}

// ── Exportar datos para Power BI ──
function exportarParaPowerBI() {
  try {
    const ss      = SpreadsheetApp.getActive();
    const maestro = ss.getSheetByName(CONFIG.HOJA);
    if (!maestro) { SpreadsheetApp.getUi().alert('❌ No hay datos.'); return; }

    const M  = CONFIG.COL;
    const tz = Session.getScriptTimeZone();
    const ahora = new Date();

    // ── PBI_Participantes — tabla plana (1 fila = 1 participante) ──
    let pbiP = ss.getSheetByName('PBI_Participantes');
    if (!pbiP) pbiP = ss.insertSheet('PBI_Participantes');
    else pbiP.clear();

    const hP = ['ID','Fecha','Nombre','DPI','Edad','Genero','Telefono','Zona','Email',
                'Educacion','Situacion_Laboral','Fortalezas','Objetivo',
                'Perfil','Prioridad','Puntaje',
                'Dim_Educativo','Dim_Laboral','Dim_Digital','Dim_Vocacional','Dim_Barreras','Dim_Red_Apoyo',
                'Estado','FuenteActualizacion'];
    pbiP.getRange(1,1,1,hP.length).setValues([hP])
      .setFontWeight('bold').setBackground('#1a237e').setFontColor('#fff');

    const datos = maestro.getLastRow() > 1
      ? maestro.getRange(2,1,maestro.getLastRow()-1,26).getValues().filter(r => r[M.ID-1])
      : [];

    if (datos.length) {
      const filas = datos.map(r => [
        r[M.ID-1], r[M.FECHA-1], r[M.NOMBRE-1], r[M.DPI-1],
        Number(r[M.EDAD-1])||'', r[M.GENERO-1], r[M.TELEFONO-1], r[M.ZONA-1], r[M.EMAIL-1],
        r[M.EDUCACION-1], r[M.LABORAL-1], r[M.FORTALEZAS-1], r[M.OBJETIVO-1],
        normalizarPerfil(r[M.PERFIL-1]), normalizarPrioridad(r[M.PRIORIDAD-1]),
        Number(r[M.PUNTAJE-1])||0,
        Number(r[M.DIM1-1])||0, Number(r[M.DIM2-1])||0, Number(r[M.DIM3-1])||0,
        Number(r[M.DIM4-1])||0, Number(r[M.DIM5-1])||0, Number(r[M.DIM6-1])||0,
        r[M.ESTADO-1], ahora
      ]);
      pbiP.getRange(2,1,filas.length,hP.length).setValues(filas);
    }
    pbiP.setFrozenRows(1);
    pbiP.autoResizeColumns(1,hP.length);

    // ── PBI_Indicadores — tabla de KPIs con código IL ──
    let pbiI = ss.getSheetByName('PBI_Indicadores');
    if (!pbiI) pbiI = ss.insertSheet('PBI_Indicadores');
    else pbiI.clear();

    const hI = ['Codigo','Indicador','Valor_Texto','Valor_Num','Criterio','Periodo','Fecha_Actualizacion'];
    pbiI.getRange(1,1,1,hI.length).setValues([hI])
      .setFontWeight('bold').setBackground('#1a237e').setFontColor('#fff');

    const perfiles = {'Perfil A':0,'Perfil B':0,'Perfil C':0,'Perfil D':0};
    const estados  = {};
    const generos  = {};
    const puntajes = [];
    let enAcomp = 0, conexiones = 0, ingMes = 0;
    const mesAct = ahora.getMonth(), anioAct = ahora.getFullYear();

    datos.forEach(r => {
      const p = normalizarPerfil(r[M.PERFIL-1]);
      if (perfiles[p] !== undefined) perfiles[p]++;
      const e = r[M.ESTADO-1]||'Sin estado';
      estados[e] = (estados[e]||0)+1;
      const gen = String(r[M.GENERO-1]||'Sin dato');
      generos[gen] = (generos[gen]||0)+1;
      if (['Mentoría','Orientación'].includes(e)) enAcomp++;
      if (['Cierre','Completado'].includes(e)) conexiones++;
      const pt = Number(r[M.PUNTAJE-1]);
      if (pt > 0) puntajes.push(pt);
      const f = r[M.FECHA-1];
      if (f instanceof Date && f.getMonth()===mesAct && f.getFullYear()===anioAct) ingMes++;
    });

    const total   = datos.length;
    const promPt  = puntajes.length ? Math.round(puntajes.reduce((a,b)=>a+b,0)/puntajes.length) : 0;
    const periodo = Utilities.formatDate(ahora, tz, 'MMMM yyyy');
    const tasaGrad = total>0 ? Math.round(conexiones/total*100) : 0;

    const kpiRows = [
      ['IL.P.01','Número de participantes en el programa IL',               total,         total,    'Pertinencia', periodo, ahora],
      ['IL.P.02','Participantes en acompañamiento profesional',              enAcomp,       enAcomp,  'Pertinencia', periodo, ahora],
      ['IL.P.04','Personas alcanzadas (registradas este mes)',               ingMes,        ingMes,   'Pertinencia', periodo, ahora],
      ['IL.R.07','Número de conexiones laborales (Cierre/Completado)',       conexiones,    conexiones,'Eficacia',   periodo, ahora],
      ['IL.R.05','Tasa de graduación % (sobre total)',                       tasaGrad+'%',  tasaGrad, 'Eficacia',    periodo, ahora],
      ['IL.P.01_A','Perfil A — Listos para empleo',                         perfiles['Perfil A'], perfiles['Perfil A'], 'Desagregación', periodo, ahora],
      ['IL.P.01_B','Perfil B — Orientación vocacional',                     perfiles['Perfil B'], perfiles['Perfil B'], 'Desagregación', periodo, ahora],
      ['IL.P.01_C','Perfil C — Desarrollo de capacidades',                  perfiles['Perfil C'], perfiles['Perfil C'], 'Desagregación', periodo, ahora],
      ['IL.P.01_D','Perfil D — Barreras críticas (URGENTE)',                perfiles['Perfil D'], perfiles['Perfil D'], 'Desagregación', periodo, ahora],
      ['DIAG.PROM','Puntaje promedio diagnóstico (sobre 60)',                promPt+'/60',  promPt,   'Diagnóstico', periodo, ahora],
    ];

    // Agregar filas por estado
    Object.entries(estados).sort((a,b)=>b[1]-a[1]).forEach(([est,cnt]) => {
      kpiRows.push(['ESTADO_'+est.toUpperCase().replace(/\s/g,'_'), 'Participantes en estado: '+est, cnt, cnt, 'Estado', periodo, ahora]);
    });

    // Agregar filas por género (desagregación)
    Object.entries(generos).sort((a,b)=>b[1]-a[1]).forEach(([gen,cnt]) => {
      kpiRows.push(['GENERO_'+gen.toUpperCase().replace(/\s/g,'_'), 'Género: '+gen, cnt, cnt, 'Género', periodo, ahora]);
    });

    pbiI.getRange(2,1,kpiRows.length,hI.length).setValues(kpiRows);
    pbiI.setFrozenRows(1);
    pbiI.setColumnWidth(2,320);
    pbiI.autoResizeColumns(1,1);

    SpreadsheetApp.getUi().alert(
      '✅ EXPORTACIÓN PARA POWER BI LISTA\n\n'+
      '📊 PBI_Participantes: '+total+' registros\n'+
      '📈 PBI_Indicadores: '+kpiRows.length+' filas (IL.P.01, IL.P.02, IL.R.07…)\n\n'+
      '🔗 CONECTAR EN POWER BI DESKTOP:\n'+
      '1. Inicio → Obtener datos → Google Sheets\n'+
      '2. Pega la URL de este Google Sheets\n'+
      '3. Selecciona: PBI_Participantes y PBI_Indicadores\n'+
      '4. La conexión se actualiza automáticamente al hacer esta exportación\n\n'+
      '💡 Recomendado: ejecutar esta función 1 vez al día/semana\n'+
      '   o crear un trigger automático mensual.'
    );

  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('exportarParaPowerBI', e);
  }
}

// ============================================================================
// AUTO-ACTUALIZACIÓN POWER BI — Triggers automáticos
// ============================================================================

// Triggers
const TRIGGER_PBI_FN  = 'triggerSilenciosoExportPBI';
const TRIGGER_SYNC_FN = 'triggerAutoSyncCompleta';
const PROP_TRIGGER_FREQ = 'pbi_trigger_freq';
const PROP_SYNC_HORA    = 'sync_hora_diaria';

/**
 * Versión silenciosa de exportarParaPowerBI, ejecutada por el trigger.
 * NO muestra alertas UI — solo actualiza las hojas PBI.
 */
function triggerSilenciosoExportPBI() {
  try {
    const ss      = SpreadsheetApp.getActive();
    const maestro = ss.getSheetByName(CONFIG.HOJA);
    if (!maestro || maestro.getLastRow() < 2) return;

    const M  = CONFIG.COL;
    const tz = Session.getScriptTimeZone();
    const ahora = new Date();

    // ── PBI_Participantes ──
    let pbiP = ss.getSheetByName('PBI_Participantes');
    if (!pbiP) pbiP = ss.insertSheet('PBI_Participantes');
    else pbiP.clear();

    const hP = ['ID','Fecha','Nombre','DPI','Edad','Genero','Telefono','Zona','Email',
                'Educacion','Situacion_Laboral','Fortalezas','Objetivo',
                'Perfil','Prioridad','Puntaje',
                'Dim_Educativo','Dim_Laboral','Dim_Digital','Dim_Vocacional','Dim_Barreras','Dim_Red_Apoyo',
                'Estado','FechaActualizacion'];
    pbiP.getRange(1,1,1,hP.length).setValues([hP])
      .setFontWeight('bold').setBackground('#1a237e').setFontColor('#fff');

    const datos = maestro.getRange(2,1,maestro.getLastRow()-1,26).getValues().filter(r => r[M.ID-1]);
    if (datos.length) {
      const filas = datos.map(r => [
        r[M.ID-1], r[M.FECHA-1], r[M.NOMBRE-1], r[M.DPI-1],
        Number(r[M.EDAD-1])||'', r[M.GENERO-1], r[M.TELEFONO-1], r[M.ZONA-1], r[M.EMAIL-1],
        r[M.EDUCACION-1], r[M.LABORAL-1], r[M.FORTALEZAS-1], r[M.OBJETIVO-1],
        normalizarPerfil(r[M.PERFIL-1]), normalizarPrioridad(r[M.PRIORIDAD-1]),
        Number(r[M.PUNTAJE-1])||0,
        Number(r[M.DIM1-1])||0, Number(r[M.DIM2-1])||0, Number(r[M.DIM3-1])||0,
        Number(r[M.DIM4-1])||0, Number(r[M.DIM5-1])||0, Number(r[M.DIM6-1])||0,
        r[M.ESTADO-1], ahora
      ]);
      pbiP.getRange(2,1,filas.length,hP.length).setValues(filas);
    }
    pbiP.setFrozenRows(1);

    // ── PBI_Indicadores ──
    let pbiI = ss.getSheetByName('PBI_Indicadores');
    if (!pbiI) pbiI = ss.insertSheet('PBI_Indicadores');
    else pbiI.clear();

    const hI = ['Codigo','Indicador','Valor_Texto','Valor_Num','Criterio','Periodo','Fecha_Actualizacion'];
    pbiI.getRange(1,1,1,hI.length).setValues([hI])
      .setFontWeight('bold').setBackground('#1a237e').setFontColor('#fff');

    const perfiles  = {'Perfil A':0,'Perfil B':0,'Perfil C':0,'Perfil D':0};
    const estados   = {};
    const generos   = {};
    const puntajes  = [];
    let enAcomp = 0, conexiones = 0, ingMes = 0;
    const mesAct = ahora.getMonth(), anioAct = ahora.getFullYear();

    datos.forEach(r => {
      const p = normalizarPerfil(r[M.PERFIL-1]);
      if (perfiles[p] !== undefined) perfiles[p]++;
      const e = r[M.ESTADO-1]||'Sin estado';
      estados[e] = (estados[e]||0)+1;
      const gen = String(r[M.GENERO-1]||'Sin dato');
      generos[gen] = (generos[gen]||0)+1;
      if (['Mentoría','Orientación'].includes(e)) enAcomp++;
      if (['Cierre','Completado'].includes(e)) conexiones++;
      const pt = Number(r[M.PUNTAJE-1]);
      if (pt > 0) puntajes.push(pt);
      const f = r[M.FECHA-1];
      if (f instanceof Date && f.getMonth()===mesAct && f.getFullYear()===anioAct) ingMes++;
    });

    const total   = datos.length;
    const promPt  = puntajes.length ? Math.round(puntajes.reduce((a,b)=>a+b,0)/puntajes.length) : 0;
    const periodo = Utilities.formatDate(ahora, tz, 'MMMM yyyy');
    const tasaGrad = total > 0 ? Math.round(conexiones/total*100) : 0;

    const kpiRows = [
      ['IL.P.01', 'Número de participantes en el programa IL',            total,         total,     'Pertinencia',   periodo, ahora],
      ['IL.P.02', 'Participantes en acompañamiento profesional',           enAcomp,       enAcomp,   'Pertinencia',   periodo, ahora],
      ['IL.P.04', 'Personas registradas este mes',                         ingMes,        ingMes,    'Pertinencia',   periodo, ahora],
      ['IL.R.07', 'Conexiones laborales (Cierre/Completado)',              conexiones,    conexiones, 'Eficacia',     periodo, ahora],
      ['IL.R.05', 'Tasa de graduación % (sobre total)',                    tasaGrad+'%',  tasaGrad,   'Eficacia',     periodo, ahora],
      ['IL.P.01_A','Perfil A — Listos para empleo',                        perfiles['Perfil A'], perfiles['Perfil A'], 'Desagregación', periodo, ahora],
      ['IL.P.01_B','Perfil B — Orientación vocacional',                    perfiles['Perfil B'], perfiles['Perfil B'], 'Desagregación', periodo, ahora],
      ['IL.P.01_C','Perfil C — Desarrollo de capacidades',                 perfiles['Perfil C'], perfiles['Perfil C'], 'Desagregación', periodo, ahora],
      ['IL.P.01_D','Perfil D — Barreras críticas',                         perfiles['Perfil D'], perfiles['Perfil D'], 'Desagregación', periodo, ahora],
      ['DIAG.PROM','Puntaje promedio diagnóstico (sobre 60)',               promPt+'/60',  promPt,     'Diagnóstico',  periodo, ahora],
    ];
    Object.entries(estados).forEach(([est,cnt]) =>
      kpiRows.push(['ESTADO_'+est.toUpperCase().replace(/\s/g,'_'), 'Estado: '+est, cnt, cnt, 'Estado', periodo, ahora]));
    Object.entries(generos).forEach(([gen,cnt]) =>
      kpiRows.push(['GENERO_'+gen.toUpperCase().replace(/\s/g,'_'), 'Género: '+gen, cnt, cnt, 'Género', periodo, ahora]));

    pbiI.getRange(2,1,kpiRows.length,hI.length).setValues(kpiRows);
    pbiI.setFrozenRows(1);

    // Registrar timestamp en PropertiesService
    PropertiesService.getScriptProperties().setProperty('pbi_last_update', ahora.toISOString());

    Logger.log('✅ PBI actualizado: ' + ahora.toLocaleString());
  } catch(e) {
    logError('triggerSilenciosoExportPBI', e);
  }
}

/**
 * Muestra diálogo para configurar la frecuencia de auto-actualización.
 */
function configurarAutoActualizacion() {
  const props    = PropertiesService.getScriptProperties();
  const freqActual = props.getProperty(PROP_TRIGGER_FREQ) || 'ninguna';
  const ultimaAct  = props.getProperty('pbi_last_update');
  const ultimaTxt  = ultimaAct
    ? Utilities.formatDate(new Date(ultimaAct), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    : 'Nunca';

  const triggerActivo = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === TRIGGER_PBI_FN);

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head>
    <style>
      body{font-family:'Segoe UI',Arial;font-size:13px;margin:0;background:#f5f5f5}
      .hdr{background:#1a237e;color:#fff;padding:16px 20px}
      .hdr h2{margin:0;font-size:15px}
      .hdr p{margin:5px 0 0;font-size:10px;opacity:.8}
      .body{padding:20px}
      .status{background:${triggerActivo?'#e8f5e9':'#fff3e0'};
              border-left:4px solid ${triggerActivo?'#43a047':'#fb8c00'};
              padding:10px 14px;border-radius:4px;margin-bottom:18px;font-size:12px}
      .status strong{color:${triggerActivo?'#2e7d32':'#e65100'}}
      .opt{background:#fff;border:2px solid #e0e0e0;border-radius:8px;padding:14px;
           margin-bottom:10px;cursor:pointer;transition:all .15s;display:flex;align-items:center;gap:12px}
      .opt:hover{border-color:#1a237e;background:#f0f2ff}
      .opt.sel{border-color:#1a237e;background:#e8eaf6}
      .opt-ico{font-size:22px;flex-shrink:0}
      .opt-txt strong{display:block;font-size:13px;color:#1a237e}
      .opt-txt span{font-size:11px;color:#777}
      .btn{width:100%;padding:12px;background:#1a237e;color:#fff;border:none;
           border-radius:6px;font-size:13px;font-weight:700;cursor:pointer;margin-top:14px}
      .btn:hover{background:#283593}
      .btn.danger{background:#c62828;margin-top:8px}
      .btn.danger:hover{background:#b71c1c}
      .note{font-size:10px;color:#999;margin-top:10px;text-align:center;line-height:1.5}
    </style></head><body>
    <div class="hdr">
      <h2>⏰ Auto-actualización Power BI</h2>
      <p>Las hojas PBI_Participantes y PBI_Indicadores se actualizarán automáticamente</p>
    </div>
    <div class="body">
      <div class="status">
        Estado: <strong>${triggerActivo ? '✅ ACTIVO — cada '+freqActual : '🔴 INACTIVO'}</strong><br>
        Última actualización: <strong>${ultimaTxt}</strong>
      </div>

      <p style="margin:0 0 12px;font-weight:bold;color:#333">Selecciona la frecuencia:</p>

      <div class="opt ${freqActual==='hora'?'sel':''}" onclick="sel(this,'hora')">
        <span class="opt-ico">🕐</span>
        <div class="opt-txt"><strong>Cada hora</strong><span>Máxima frescura de datos. Ideal si el equipo registra actividad durante el día.</span></div>
      </div>
      <div class="opt ${freqActual==='6horas'?'sel':''}" onclick="sel(this,'6horas')">
        <span class="opt-ico">🕕</span>
        <div class="opt-txt"><strong>Cada 6 horas</strong><span>Balance entre frecuencia y cuota. Recomendado para la mayoría de casos.</span></div>
      </div>
      <div class="opt ${freqActual==='dia'?'sel':''}" onclick="sel(this,'dia')">
        <span class="opt-ico">📅</span>
        <div class="opt-txt"><strong>Una vez al día (medianoche)</strong><span>Ideal si Power BI se revisa cada mañana con datos del día anterior.</span></div>
      </div>
      <div class="opt ${freqActual==='semana'?'sel':''}" onclick="sel(this,'semana')">
        <span class="opt-ico">📆</span>
        <div class="opt-txt"><strong>Cada semana (lunes)</strong><span>Para reportes semanales. Usa menos cuota de ejecución.</span></div>
      </div>

      <button class="btn" onclick="guardar()">✅ Activar Auto-actualización</button>
      <button class="btn danger" onclick="desactivar()">🔴 Desactivar</button>
      <div class="note">Funciona 24/7 aunque el archivo esté cerrado.<br>Puedes ver ejecuciones en Apps Script → Ejecuciones.</div>
    </div>
    <script>
      var freq = '${freqActual==='ninguna'?'dia':freqActual}';
      function sel(el,f){
        document.querySelectorAll('.opt').forEach(o=>o.classList.remove('sel'));
        el.classList.add('sel'); freq=f;
      }
      function guardar(){
        google.script.run.withSuccessHandler(function(msg){alert(msg);google.script.host.close();})
          .activarTriggerPBI(freq);
      }
      function desactivar(){
        google.script.run.withSuccessHandler(function(msg){alert(msg);google.script.host.close();})
          .desactivarAutoActualizacion();
      }
    </script>
    </body></html>
  `).setWidth(480).setHeight(560);

  SpreadsheetApp.getUi().showModalDialog(html, '⏰ Configurar Auto-actualización Power BI');
}

/**
 * Crea el trigger con la frecuencia seleccionada.
 * @param {string} freq 'hora' | '6horas' | 'dia' | 'semana'
 */
function activarTriggerPBI(freq) {
  try {
    // Borrar triggers anteriores del mismo handler
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === TRIGGER_PBI_FN)
      .forEach(t => ScriptApp.deleteTrigger(t));

    let builder = ScriptApp.newTrigger(TRIGGER_PBI_FN).timeBased();

    switch(freq) {
      case 'hora':
        builder.everyHours(1); break;
      case '6horas':
        builder.everyHours(6); break;
      case 'dia':
        builder.everyDays(1).atHour(0); break;
      case 'semana':
        builder.everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6); break;
      default:
        builder.everyHours(6);
    }

    builder.create();

    PropertiesService.getScriptProperties().setProperty(PROP_TRIGGER_FREQ, freq);

    // Ejecutar inmediatamente la primera vez
    triggerSilenciosoExportPBI();

    const label = {hora:'cada hora', '6horas':'cada 6 horas', dia:'diariamente a medianoche', semana:'cada lunes a las 6am'}[freq] || freq;
    return '✅ Auto-actualización activada: '+label+'\n\nLas hojas PBI_Participantes y PBI_Indicadores ya fueron actualizadas ahora mismo y seguirán actualizándose automáticamente.';
  } catch(e) {
    logError('activarTriggerPBI', e);
    return '❌ Error: ' + e.message;
  }
}

/**
 * Elimina el trigger de auto-actualización.
 */
function desactivarAutoActualizacion() {
  try {
    const triggers = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === TRIGGER_PBI_FN);

    if (!triggers.length) {
      return '⚠️ No había ningún trigger activo.';
    }

    triggers.forEach(t => ScriptApp.deleteTrigger(t));
    PropertiesService.getScriptProperties().setProperty(PROP_TRIGGER_FREQ, 'ninguna');
    SpreadsheetApp.getUi().alert('🔴 Auto-actualización desactivada.\n\nPuedes volver a activarla desde ⏰ Configurar Auto-actualización.');
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('desactivarAutoActualizacion', e);
  }
}

// ============================================================================
// AUTO-SYNC COMPLETA MATUTINA (Kobo + DP + PBI + Dashboard)
// ============================================================================

/** Ejecutado por trigger: sincroniza todo sin alertas UI */
function triggerAutoSyncCompleta() {
  try {
    const ss = SpreadsheetApp.getActive();
    Logger.log('🔄 Auto-sync iniciada: ' + new Date());

    // 1. Sincronizar Kobo (silencioso)
    sincronizarConValidaciones(true);
    Logger.log('✅ Kobo sincronizado');

    // 2. Sincronizar DP_Empleabilidad → Derivados (silencioso)
    importarDPADerivados(true);
    Logger.log('✅ Derivados actualizados');

    // 3. Actualizar hojas Power BI
    triggerSilenciosoExportPBI();
    Logger.log('✅ PBI actualizado');

    // 4. Actualizar Dashboard
    actualizarDashboard(ss);
    Logger.log('✅ Dashboard actualizado');

    PropertiesService.getScriptProperties().setProperty('sync_last_run', new Date().toISOString());
    Logger.log('✅ Auto-sync completa finalizada');
  } catch(e) {
    logError('triggerAutoSyncCompleta', e);
  }
}

/** Muestra diálogo para configurar la sync matutina */
function configurarSyncMatutina() {
  const props      = PropertiesService.getScriptProperties();
  const horaActual = props.getProperty(PROP_SYNC_HORA) || '7';
  const ultimaRun  = props.getProperty('sync_last_run');
  const ultimaTxt  = ultimaRun
    ? Utilities.formatDate(new Date(ultimaRun), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    : 'Nunca';
  const activo = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === TRIGGER_SYNC_FN);

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head>
    <style>
      body{font-family:'Segoe UI',Arial;font-size:13px;margin:0;background:#f5f5f5}
      .hdr{background:#1a237e;color:#fff;padding:16px 20px}
      .hdr h2{margin:0;font-size:15px}
      .hdr p{margin:4px 0 0;font-size:10px;opacity:.85}
      .body{padding:20px}
      .status{background:${activo?'#e8f5e9':'#fff3e0'};
              border-left:4px solid ${activo?'#43a047':'#fb8c00'};
              padding:10px 14px;border-radius:4px;margin-bottom:18px;font-size:12px}
      .steps{background:#e8eaf6;border-radius:6px;padding:12px 14px;margin-bottom:16px;font-size:11px;line-height:1.8;color:#333}
      .steps strong{color:#1a237e}
      .hora-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:16px}
      .hora-opt{background:#fff;border:2px solid #e0e0e0;border-radius:8px;padding:10px 6px;
                text-align:center;cursor:pointer;transition:all .15s}
      .hora-opt:hover{border-color:#1a237e;background:#f0f2ff}
      .hora-opt.sel{border-color:#1a237e;background:#e8eaf6}
      .hora-opt strong{display:block;font-size:16px;color:#1a237e}
      .hora-opt span{font-size:9px;color:#777}
      .btn{width:100%;padding:12px;background:#1a237e;color:#fff;border:none;
           border-radius:6px;font-size:13px;font-weight:700;cursor:pointer;margin-top:4px}
      .btn:hover{background:#283593}
      .btn.danger{background:#c62828;margin-top:8px}
    </style></head><body>
    <div class="hdr">
      <h2>⚡ Auto-Sync Completa</h2>
      <p>Sincroniza Kobo + Derivados + PBI + Dashboard automáticamente</p>
    </div>
    <div class="body">
      <div class="status">
        Estado: <strong>${activo ? '✅ ACTIVO — a las '+horaActual+':00' : '🔴 INACTIVO'}</strong><br>
        Última ejecución: <strong>${ultimaTxt}</strong>
      </div>
      <div class="steps">
        <strong>Cada mañana ejecuta automáticamente:</strong><br>
        1️⃣ Descarga nuevas respuestas de Kobo<br>
        2️⃣ Importa nuevos registros de DP_Empleabilidad → Derivados<br>
        3️⃣ Actualiza hojas PBI_Participantes e PBI_Indicadores<br>
        4️⃣ Reconstruye el Dashboard con datos frescos
      </div>
      <p style="margin:0 0 10px;font-weight:bold;color:#333">Hora de ejecución:</p>
      <div class="hora-grid">
        <div class="hora-opt ${horaActual==='5'?'sel':''}" onclick="sel(this,'5')"><strong>5:00</strong><span>Madrugada</span></div>
        <div class="hora-opt ${horaActual==='6'?'sel':''}" onclick="sel(this,'6')"><strong>6:00</strong><span>Temprano</span></div>
        <div class="hora-opt ${horaActual==='7'?'sel':''}" onclick="sel(this,'7')"><strong>7:00</strong><span>✅ Recomen.</span></div>
        <div class="hora-opt ${horaActual==='8'?'sel':''}" onclick="sel(this,'8')"><strong>8:00</strong><span>Al llegar</span></div>
      </div>
      <button class="btn" onclick="guardar()">⚡ Activar Auto-Sync Diaria</button>
      <button class="btn danger" onclick="desactivar()">🔴 Desactivar</button>
    </div>
    <script>
      var hora = '${horaActual}';
      function sel(el,h){ document.querySelectorAll('.hora-opt').forEach(o=>o.classList.remove('sel')); el.classList.add('sel'); hora=h; }
      function guardar(){
        google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).activarSyncDiaria(hora);
      }
      function desactivar(){
        google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).desactivarSyncDiaria();
      }
    </script></body></html>
  `).setWidth(460).setHeight(490);
  SpreadsheetApp.getUi().showModalDialog(html, '⚡ Auto-Sync Completa');
}

function activarSyncDiaria(hora) {
  try {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === TRIGGER_SYNC_FN)
      .forEach(t => ScriptApp.deleteTrigger(t));

    ScriptApp.newTrigger(TRIGGER_SYNC_FN).timeBased()
      .everyDays(1).atHour(parseInt(hora)||7).create();

    PropertiesService.getScriptProperties().setProperty(PROP_SYNC_HORA, String(hora));

    // Ejecutar ahora mismo
    triggerAutoSyncCompleta();

    return '✅ Auto-Sync activada a las '+hora+':00 cada día.\n\nYa se ejecutó una primera sincronización completa ahora mismo.';
  } catch(e) {
    logError('activarSyncDiaria', e);
    return '❌ Error: ' + e.message;
  }
}

function desactivarSyncDiaria() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === TRIGGER_SYNC_FN);
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  PropertiesService.getScriptProperties().setProperty(PROP_SYNC_HORA, '');
  SpreadsheetApp.getUi().alert('🔴 Auto-Sync diaria desactivada.');
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

// Abre la app de fichas — no requiere seleccionar fila
function verFicha() {
  try {
    const datos = obtenerParticipantes();
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(construirFichaHtml(datos)).setWidth(900).setHeight(720),
      '👤 Participantes — Paso a Paso'
    );
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e.message); }
}

function construirFichaHtml(datos) {
  const PERFIL_COLOR = {
    'Perfil A': { bg:'#274e13', fg:'#d9ead3' },
    'Perfil B': { bg:'#1c4587', fg:'#cfe2f3' },
    'Perfil C': { bg:'#7f6000', fg:'#fff2cc' },
    'Perfil D': { bg:'#660000', fg:'#f4cccc' }
  };

  function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function ini(n){ return String(n||'').split(' ').slice(0,2).map(w=>w[0]||'').join('').toUpperCase(); }
  function tel502(t){ const s=String(t||'').replace(/\D/g,''); return s.length===8?'502'+s:s; }

  const filas = datos.map(p => {
    const c = PERFIL_COLOR[p.perfil] || { bg:'#555', fg:'#f0f0f0' };
    const tel = tel502(p.telefono);
    const waMsg = encodeURIComponent('Hola '+p.nombre+', somos el equipo de Paso a Paso de Creamos Guatemala.');
    const waBtn = tel
      ? '<a href="https://wa.me/'+tel+'?text='+waMsg+'" target="_blank" style="display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;border-radius:50%;background:#25d366;color:#fff;text-decoration:none;font-size:13px">💬</a>'
      : '';
    const docBtn = p.docUrl
      ? '<a href="'+p.docUrl+'" target="_blank" style="display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;border-radius:50%;background:#e8eaf6;color:#1a237e;text-decoration:none;font-size:13px;margin-left:3px">📄</a>'
      : '';
    const urgente = p.prioridad === 'CRÍTICO'
      ? '<span style="background:#c62828;color:#fff;padding:1px 5px;border-radius:6px;font-size:9px;font-weight:700;margin-left:3px">URGENTE</span>' : '';
    return `<tr class="fila" style="cursor:pointer" onclick="abrirPerfil(this)" data-perfil='${JSON.stringify(p).replace(/'/g,"&apos;")}' data-buscar="${esc(p.nombre+' '+p.id+' '+(p.zona||'')).toLowerCase()}">
      <td style="padding:8px 10px">
        <div style="display:flex;align-items:center;gap:8px">
          <div style="width:32px;height:32px;border-radius:50%;background:${c.bg};color:${c.fg};display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0">${ini(p.nombre)}</div>
          <div>
            <div style="font-size:12px;font-weight:700;color:#1a237e">${esc(p.nombre)}</div>
            <div style="font-size:9px;color:#9e9e9e">Creamos ID: ${esc(p.id)}</div>
          </div>
        </div>
      </td>
      <td style="padding:8px 6px;vertical-align:middle">
        <span style="background:${c.fg};color:${c.bg};padding:2px 7px;border-radius:8px;font-size:10px;font-weight:700">${esc(p.perfil||'—')}</span>${urgente}
      </td>
      <td style="padding:8px 6px;vertical-align:middle">
        <span style="background:#e8eaf6;color:#3949ab;padding:2px 7px;border-radius:8px;font-size:10px;font-weight:700">${esc(p.estado||'—')}</span>
      </td>
      <td style="padding:8px 6px;vertical-align:middle;font-size:11px;color:#555">${esc(p.telefono||'—')}</td>
      <td style="padding:8px 6px;vertical-align:middle;font-size:11px;color:#555">${esc(p.zona||'—')}</td>
      <td style="padding:8px 6px;vertical-align:middle;text-align:center">${waBtn}${docBtn}</td>
    </tr>`;
  }).join('');

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;background:#f0f2ff;height:100vh;display:flex;flex-direction:column;overflow:hidden}
.tb{background:#1a237e;color:#fff;padding:10px 14px;display:flex;align-items:center;gap:8px;flex-shrink:0}
.tb h2{font-size:13px;font-weight:700;flex:1}
.badge{font-size:10px;background:rgba(255,255,255,.2);padding:2px 8px;border-radius:10px}
.bar{padding:7px 12px;background:#fff;border-bottom:1px solid #e8eaf6;display:flex;gap:6px;align-items:center;flex-shrink:0}
.srch{padding:7px 12px;background:#fff;border-bottom:1px solid #e8eaf6;flex-shrink:0}
.srch input{width:100%;padding:7px 12px;border:1px solid #c5cae9;border-radius:20px;font-size:12px;outline:none}
.srch input:focus{border-color:#1a237e}
.filtros{padding:5px 12px;background:#fff;border-bottom:1px solid #e8eaf6;display:flex;gap:4px;flex-wrap:wrap;flex-shrink:0}
.flt{padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;border:1.5px solid #c5cae9;background:#fff;cursor:pointer;color:#555}
.flt.on{background:#1a237e;color:#fff;border-color:#1a237e}
.tabla-wrap{overflow-y:auto;flex:1}
table{width:100%;border-collapse:collapse;background:#fff}
thead th{padding:7px 10px;background:#e8eaf6;color:#1a237e;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;position:sticky;top:0;z-index:1;border-bottom:2px solid #c5cae9;text-align:left}
.fila:hover{background:#f0f2ff}
.fila.oculta{display:none}
.empty{padding:40px;text-align:center;color:#9e9e9e;font-size:12px}
</style></head><body>
<div class="tb"><h2>👤 Participantes</h2><span class="badge" id="cnt">${datos.length} participantes</span></div>
<div class="bar">
  <span style="font-size:11px;color:#666">Total cargados: <strong>${datos.length}</strong></span>
</div>
<div class="srch"><input id="q" placeholder="🔍 Buscar por nombre, Creamos ID, zona…" oninput="buscar(this.value)"></div>
<div class="filtros" id="filtros">
  <button class="flt on" onclick="filtrar(this,'')">Todos</button>
  <button class="flt" onclick="filtrar(this,'Perfil A')" style="color:#274e13;border-color:#a8d5a2">A</button>
  <button class="flt" onclick="filtrar(this,'Perfil B')" style="color:#1c4587;border-color:#9bbfe0">B</button>
  <button class="flt" onclick="filtrar(this,'Perfil C')" style="color:#7f6000;border-color:#f0d060">C</button>
  <button class="flt" onclick="filtrar(this,'Perfil D')" style="color:#660000;border-color:#e08080">D</button>
  <button class="flt" onclick="filtrar(this,'Orientación')">Orient.</button>
  <button class="flt" onclick="filtrar(this,'Mentoría')">Ment.</button>
  <button class="flt" onclick="filtrar(this,'Formación')">Form.</button>
  <button class="flt" onclick="filtrar(this,'Inactivo')">Inact.</button>
</div>
<div class="tabla-wrap">
  ${datos.length === 0 ? '<div class="empty">📭 No hay participantes en la hoja Maestro</div>' :
  `<table><thead><tr>
    <th>Participante</th><th>Perfil</th><th>Estado</th><th>Teléfono</th><th>Zona</th><th>Acciones</th>
  </tr></thead><tbody id="tbody">${filas}</tbody></table>`}
</div>
<script>
var filtroActual='';
function buscar(q){
  var filas=document.querySelectorAll('.fila');
  var v=0;
  filas.forEach(function(f){
    var b=f.dataset.buscar||'';
    var ok=(!q||b.includes(q.toLowerCase()))&&(!filtroActual||b.includes(filtroActual.toLowerCase()));
    f.classList.toggle('oculta',!ok);
    if(ok)v++;
  });
  document.getElementById('cnt').textContent=v+' / ${datos.length}';
}
function filtrar(btn,val){
  filtroActual=val;
  document.querySelectorAll('.flt').forEach(function(b){b.classList.remove('on');});
  btn.classList.add('on');
  buscar(document.getElementById('q').value);
}
function abrirPerfil(row) {
  const p = JSON.parse(row.dataset.perfil);
  const modal = document.getElementById('perfilModal');
  const contenido = document.getElementById('perfilContenido');
  const PCOL = {'Perfil A':{bg:'#274e13',fg:'#d9ead3'},'Perfil B':{bg:'#1c4587',fg:'#cfe2f3'},'Perfil C':{bg:'#7f6000',fg:'#fff2cc'},'Perfil D':{bg:'#660000',fg:'#fce5cd'},'Orientación':{bg:'#4a148c',fg:'#e1bee7'},'Mentoría':{bg:'#006064',fg:'#e0f7fa'},'Formación':{bg:'#bf360c',fg:'#fbe9e7'},'Inactivo':{bg:'#424242',fg:'#eeeeee'}};
  const c = PCOL[p.perfil] || {bg:'#37474f',fg:'#eceff1'};
  const ini = p.nombre.split(' ').slice(0,2).map(function(w){return (w[0]||'').toUpperCase();}).join('');
  const tel = (function(t){var s=String(t||'').replace(/\D/g,'');return s.length===8?'502'+s:s;})(p.telefono);
  const waUrl = tel ? 'https://wa.me/'+tel+'?text='+encodeURIComponent('Hola '+p.nombre+', somos el equipo de Paso a Paso de Creamos Guatemala.') : '';

  function fila(label, val) {
    return '<div style="padding:10px 0;display:flex;justify-content:space-between;border-bottom:1px solid #f5f5f5;"><span style="font-size:9px;font-weight:600;color:#9e9e9e;text-transform:uppercase;">'+label+'</span><span style="font-size:12px;color:#212121;font-weight:500;text-align:right;">'+(val||'<span style="color:#bdbdbd;">—</span>')+'</span></div>';
  }

  var html =
    '<div style="background:linear-gradient(135deg,#1a237e 0%,#283593 100%);padding:20px;color:#fff;margin:-14px -16px 0;margin-bottom:0;">' +
      '<div style="display:flex;align-items:center;gap:12px;margin-bottom:12px;">' +
        '<div style="width:52px;height:52px;border-radius:50%;background:rgba(255,255,255,.2);border:2px solid rgba(255,255,255,.4);display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:700;flex-shrink:0;">'+ini+'</div>' +
        '<div style="flex:1;">' +
          '<div style="font-size:14px;font-weight:700;line-height:1.2;">'+p.nombre+'</div>' +
          '<div style="font-size:9px;opacity:.8;margin-top:2px;">Creamos ID: '+p.id+'</div>' +
        '</div>' +
        (p.puntaje ? '<div style="text-align:center;background:rgba(255,255,255,.2);border-radius:8px;padding:6px 10px;flex-shrink:0;"><div style="font-size:18px;font-weight:800;">'+p.puntaje+'</div><div style="font-size:7px;opacity:.85;text-transform:uppercase;">Puntaje</div></div>' : '') +
      '</div>' +
      '<div style="display:flex;gap:4px;flex-wrap:wrap;">' +
        (p.perfil ? '<span style="background:'+c.fg+';color:'+c.bg+';padding:3px 9px;border-radius:10px;font-size:9px;font-weight:700;">'+p.perfil+'</span>' : '') +
        (p.estado ? '<span style="background:rgba(255,255,255,.2);padding:3px 9px;border-radius:10px;font-size:9px;">'+p.estado+'</span>' : '') +
        (p.prioridad==='CRÍTICO' ? '<span style="background:#c62828;padding:3px 9px;border-radius:10px;font-size:9px;font-weight:700;">URGENTE</span>' : '') +
      '</div>' +
    '</div>' +

    '<div style="padding:12px 16px;border-bottom:1px solid #e8eaf6;display:flex;gap:0;flex-shrink:0;">' +
      '<button onclick="tabPerfil()" style="flex:1;padding:8px;background:transparent;border:none;border-bottom:2px solid #1a237e;color:#1a237e;font-size:11px;font-weight:700;cursor:pointer;text-transform:uppercase;letter-spacing:.3px;">Perfil</button>' +
      '<button onclick="tabDimensiones()" style="flex:1;padding:8px;background:transparent;border:none;border-bottom:2px solid transparent;color:#9e9e9e;font-size:11px;font-weight:700;cursor:pointer;text-transform:uppercase;letter-spacing:.3px;">Dimensiones</button>' +
    '</div>' +

    '<div id="tabPerfil" style="padding:14px 16px;overflow-y:auto;max-height:300px;">' +
      '<div style="font-size:9px;font-weight:700;color:#9e9e9e;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;margin-top:0;">📋 Datos Personales</div>' +
      fila('DPI', p.dpi) +
      fila('Edad', p.edad ? p.edad+' años' : '') +
      fila('Género', p.genero) +
      fila('Teléfono', p.telefono) +
      fila('Zona', p.zona) +
      fila('Email', p.email) +
      '<div style="font-size:9px;font-weight:700;color:#9e9e9e;text-transform:uppercase;letter-spacing:.5px;margin:10px 0 8px;">📚 Perfil Socioeconómico</div>' +
      fila('Educación', p.educacion) +
      fila('Situación Laboral', p.laboral) +
      fila('Fortalezas', p.fortalezas) +
      fila('Objetivo', p.objetivo) +
    '</div>' +

    '<div id="tabDimensiones" style="display:none;padding:16px;overflow-y:hidden;">' +
      '<div style="position:relative;width:100%;height:260px;">' +
        '<canvas id="radarChart" style="position:absolute;top:0;left:0;width:100%!important;height:260px!important;"></canvas>' +
      '</div>' +
      '<div id="dimList" style="margin-top:10px;display:grid;grid-template-columns:1fr 1fr;gap:4px;"></div>' +
    '</div>' +

    '<div style="padding:10px 16px;border-top:1px solid #e8eaf6;background:#f9f9f9;display:flex;gap:6px;flex-shrink:0;">' +
      (waUrl ? '<a href="'+waUrl+'" target="_blank" style="flex:1;padding:8px;background:#25d366;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:700;font-size:10px;text-align:center;text-decoration:none;display:block;line-height:1.3;">💬<br>WhatsApp</a>' : '') +
      (p.docUrl ? '<a href="'+p.docUrl+'" target="_blank" style="flex:1;padding:8px;background:#e8eaf6;color:#1a237e;border:none;border-radius:6px;cursor:pointer;font-weight:700;font-size:10px;text-align:center;text-decoration:none;display:block;line-height:1.3;">📄<br>Expediente</a>' : '') +
      '<button onclick="cerrarPerfil()" style="flex:1;padding:8px;background:#f5f5f5;color:#555;border:none;border-radius:6px;cursor:pointer;font-weight:600;font-size:10px;line-height:1.3;">✕<br>Cerrar</button>' +
    '</div>';

  contenido.innerHTML = html;

  // Agregar funciones de pestaña al window para que se puedan llamar
  window.tabPerfil = function(){
    document.getElementById('tabPerfil').style.display = 'block';
    document.getElementById('tabDimensiones').style.display = 'none';
    document.querySelectorAll('#perfilContenido button').forEach(function(btn,i){
      if(i===0) { btn.style.borderBottomColor = '#1a237e'; btn.style.color = '#1a237e'; }
      else if(i===1) { btn.style.borderBottomColor = 'transparent'; btn.style.color = '#9e9e9e'; }
    });
  };
  window.tabDimensiones = function(){
    document.getElementById('tabPerfil').style.display = 'none';
    document.getElementById('tabDimensiones').style.display = 'block';
    document.querySelectorAll('#perfilContenido button').forEach(function(btn,i){
      if(i===0) { btn.style.borderBottomColor = 'transparent'; btn.style.color = '#9e9e9e'; }
      else if(i===1) { btn.style.borderBottomColor = '#1a237e'; btn.style.color = '#1a237e'; }
    });
    setTimeout(function(){ dibujarRadar(p); }, 100);
  };

  modal.style.display = 'flex';
}
function cerrarPerfil() {
  document.getElementById('perfilModal').style.display = 'none';
}
function dibujarRadar(p) {
  const dims = p.dims || [0,0,0,0,0,0];
  const dimLabels = ['Educativo', 'Laboral', 'Digital', 'Vocacional', 'Barreras', 'Red Apoyo'];
  if (typeof Chart === 'undefined') {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/chart.js@3.9.1/dist/chart.min.js';
    script.onload = () => { hacerRadar(p, dims, dimLabels); };
    document.head.appendChild(script);
  } else {
    hacerRadar(p, dims, dimLabels);
  }
}
function hacerRadar(p, dims, dimLabels) {
  const canvas = document.getElementById('radarChart');
  if (!canvas || !canvas.getContext) return;
  if (window.radarChartInstance) { window.radarChartInstance.destroy(); window.radarChartInstance = null; }
  const hayDatos = dims.some(function(d){ return d > 0; });
  const ctx = canvas.getContext('2d');
  const perfColors = {'Perfil A':'#2e7d32','Perfil B':'#1565c0','Perfil C':'#f57f17','Perfil D':'#b71c1c'};
  const baseColor = perfColors[p.perfil] || '#3949ab';
  function hexToRgba(h, a) {
    var r=parseInt(h.slice(1,3),16), g=parseInt(h.slice(3,5),16), b=parseInt(h.slice(5,7),16);
    return 'rgba('+r+','+g+','+b+','+a+')';
  }
  window.radarChartInstance = new Chart(ctx, {
    type: 'radar',
    data: {
      labels: dimLabels,
      datasets: [{
        label: p.nombre,
        data: hayDatos ? dims : [1,1,1,1,1,1],
        borderColor: baseColor,
        backgroundColor: hexToRgba(baseColor, 0.18),
        borderWidth: 2.5,
        fill: true,
        pointRadius: 4,
        pointHoverRadius: 6,
        pointBackgroundColor: baseColor,
        pointBorderColor: '#fff',
        pointBorderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 400 },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(26,35,126,.9)',
          padding: 8,
          callbacks: { label: function(c){ return ' '+c.raw+' / 10'; } }
        }
      },
      scales: {
        r: {
          beginAtZero: true,
          max: 10,
          min: 0,
          ticks: { stepSize: 2, font: { size: 8 }, color: '#aaa', backdropColor: 'rgba(255,255,255,.0)' },
          grid: { color: 'rgba(26,35,126,.1)', lineWidth: 1 },
          angleLines: { color: 'rgba(26,35,126,.12)', lineWidth: 1 },
          pointLabels: { font: { size: 10, weight: '600' }, color: '#3949ab', padding: 4 }
        }
      }
    }
  });
  var listEl = document.getElementById('dimList');
  if (!listEl) return;
  var colors = ['#3949ab','#1e88e5','#00897b','#43a047','#fb8c00','#e53935'];
  listEl.innerHTML = dimLabels.map(function(lbl, i){
    var v = hayDatos ? (dims[i]||0) : 0;
    var pct = (v/10)*100;
    return '<div style="padding:4px 0;">'+
      '<div style="display:flex;justify-content:space-between;font-size:9px;color:#555;margin-bottom:2px;"><span style="font-weight:600;">'+lbl+'</span><span style="color:'+colors[i]+';font-weight:700;">'+v+'/10</span></div>'+
      '<div style="background:#f0f2ff;border-radius:4px;height:5px;overflow:hidden;"><div style="width:'+pct+'%;height:100%;background:'+colors[i]+';border-radius:4px;transition:width .3s;"></div></div>'+
    '</div>';
  }).join('');
}
document.getElementById('perfilModal').addEventListener('click', function(e) { if (e.target === this) cerrarPerfil(); });
<\/script>
<div id="perfilModal" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);z-index:1000;align-items:center;justify-content:center;padding:16px;">
  <div style="background:#fff;border-radius:12px;width:100%;max-width:420px;max-height:90vh;overflow:hidden;box-shadow:0 10px 40px rgba(0,0,0,0.3);display:flex;flex-direction:column;">
    <div id="perfilContenido" style="display:flex;flex-direction:column;overflow:hidden;flex:1;"></div>
  </div>
</div>
</body></html>`;
}

// Lee SOLO el Maestro (sin caché para máxima confiabilidad)
function obtenerParticipantes() {
  const ss = SpreadsheetApp.getActive();
  const maestro = ss.getSheetByName(CONFIG.HOJA);
  if (!maestro) throw new Error('Hoja "Maestro" no existe. Usa 📥 Instalar / Reparar Sistema.');
  if (maestro.getLastRow() < 2) return [];
  const M = CONFIG.COL;
  return maestro.getRange(2, 1, maestro.getLastRow()-1, 26).getValues()
    .filter(r => r[M.ID-1])
    .map(r => ({
      id:         String(r[M.ID-1]         || ''),
      nombre:     String(r[M.NOMBRE-1]     || ''),
      perfil:     String(r[M.PERFIL-1]     || ''),
      prioridad:  String(r[M.PRIORIDAD-1]  || ''),
      puntaje:    Number(r[M.PUNTAJE-1])   || 0,
      estado:     String(r[M.ESTADO-1]     || ''),
      dpi:        String(r[M.DPI-1]        || ''),
      edad:       String(r[M.EDAD-1]       || ''),
      genero:     String(r[M.GENERO-1]     || ''),
      telefono:   String(r[M.TELEFONO-1]   || ''),
      zona:       String(r[M.ZONA-1]       || ''),
      email:      String(r[M.EMAIL-1]      || ''),
      educacion:  String(r[M.EDUCACION-1]  || ''),
      laboral:    String(r[M.LABORAL-1]    || ''),
      fortalezas: String(r[M.FORTALEZAS-1] || ''),
      objetivo:   String(r[M.OBJETIVO-1]   || ''),
      docUrl:     String(r[M.DOC_URL-1]    || ''),
      dims: [
        Number(r[M.DIM1-1]) || 0,
        Number(r[M.DIM2-1]) || 0,
        Number(r[M.DIM3-1]) || 0,
        Number(r[M.DIM4-1]) || 0,
        Number(r[M.DIM5-1]) || 0,
        Number(r[M.DIM6-1]) || 0
      ]
    }));
}

// Lee DP_Empleabilidad si es necesario (llamada separada)
function obtenerParticipantesDP() {
  try {
    const ext = SpreadsheetApp.openById(DP_EMPLEABILIDAD.SPREADSHEET_ID)
                  .getSheetByName(DP_EMPLEABILIDAD.HOJA);
    if (!ext || ext.getLastRow() < 2) return [];
    const C = DP_EMPLEABILIDAD.C;
    return ext.getRange(2, 1, ext.getLastRow()-1, 12).getValues()
      .filter(r => r[C.ID])
      .map(r => ({
        id: String(r[C.ID]||''), nombre: String(r[C.NOMBRE]||''), fuente: 'DP_Emplea',
        perfil:'', prioridad:'', puntaje:0, estado:r[C.ACTIVO]?'Activo':'Inactivo',
        dpi:String(r[C.DPI]||''), edad:String(r[C.EDAD]||''), genero:String(r[C.GENERO]||''),
        telefono:String(r[C.TELEFONO]||''), zona:'', email:'',
        educacion:String(r[C.EDUCACION]||''), laboral:'', fortalezas:'',
        objetivo:String(r[C.FORMACION]||''), docUrl:'', dims:[0,0,0,0,0,0]
      }));
  } catch(e) { return []; }
}

// Derivados con indicador de si ya están en Maestro
function obtenerDerivados() {
  try {
    const ss = SpreadsheetApp.getActive();
    const hDeriv = ss.getSheetByName('Derivados');
    if (!hDeriv || hDeriv.getLastRow() < 2) return [];

    // IDs que ya están en Maestro
    const maestro = ss.getSheetByName(CONFIG.HOJA);
    const idsEnMaestro = new Set();
    if (maestro && maestro.getLastRow() > 1) {
      maestro.getRange(2, CONFIG.COL.ID, maestro.getLastRow()-1, 1).getValues()
        .forEach(r => { if (r[0]) idsEnMaestro.add(String(r[0]).trim()); });
    }
    // Leer también teléfonos para cruzar
    const telEnMaestro = new Set();
    if (maestro && maestro.getLastRow() > 1) {
      maestro.getRange(2, CONFIG.COL.TELEFONO, maestro.getLastRow()-1, 1).getValues()
        .forEach(r => { const t = String(r[0]||'').replace(/\D/g,''); if (t) telEnMaestro.add(t); });
    }

    return hDeriv.getRange(2, 1, hDeriv.getLastRow()-1, 14).getValues()
      .filter(r => r[2]) // tiene nombre
      .map(r => {
        const tel = String(r[3]||'').replace(/\D/g,'');
        const telN = tel.length === 8 ? '502'+tel : tel;
        const enM = idsEnMaestro.has(String(r[1]).trim()) || (tel && telEnMaestro.has(tel)) || (telN && telEnMaestro.has(telN));
        return {
          nombre:    String(r[2]||''),
          telefono:  String(r[3]||''),
          genero:    String(r[4]||''),
          educacion: String(r[6]||''),
          formacion: String(r[8]||''),
          estadoDeriv: String(r[12]||'Pendiente'),
          enMaestro: enM
        };
      });
  } catch(e) {
    logError('obtenerDerivados', e);
    throw e;
  }
}

// HTML completo de la app de fichas — modal centrado, dropdown de acciones, multi-fuente
const FICHA_HTML = `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;background:#f0f2ff;color:#333;height:100vh;display:flex;flex-direction:column;overflow:hidden}
.tb{background:#1a237e;color:#fff;padding:10px 14px;display:flex;align-items:center;gap:8px;flex-shrink:0}
.tb h2{font-size:13px;font-weight:700;flex:1}.badge{font-size:10px;background:rgba(255,255,255,.2);padding:2px 8px;border-radius:10px}
.tabs{display:flex;background:#fff;border-bottom:2px solid #e8eaf6;flex-shrink:0}
.tab{flex:1;padding:8px;text-align:center;font-size:11px;font-weight:700;cursor:pointer;color:#9e9e9e;border-bottom:2px solid transparent;margin-bottom:-2px}
.tab.on{color:#1a237e;border-bottom-color:#1a237e}
.accbar{padding:6px 12px;background:#fff;border-bottom:1px solid #e8eaf6;display:flex;gap:6px;flex-shrink:0;position:relative}
.btn{padding:5px 11px;border-radius:5px;font-size:11px;font-weight:600;border:none;cursor:pointer}
.btn-p{background:#1a237e;color:#fff}.btn-p:hover{background:#283593}
.btn-g{background:#f0f2ff;color:#1a237e;border:1px solid #c5cae9}.btn-g:hover{background:#e8eaf6}
.btn-wa{background:#25d366;color:#fff}.btn-wa:hover{background:#1da851}
.dd{position:relative}
.ddm{display:none;position:absolute;top:calc(100% + 4px);left:0;background:#fff;border:1px solid #e0e0e0;border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.14);z-index:999;min-width:250px;overflow:hidden}
.ddm.open{display:block}
.dmi{display:flex;align-items:center;gap:8px;padding:9px 14px;cursor:pointer;border-bottom:1px solid #f5f5f5}
.dmi:last-child{border:none}.dmi:hover{background:#e8eaf6;color:#1a237e}
.dmi-ico{font-size:14px;width:20px;text-align:center}
.dmi-t{display:flex;flex-direction:column}.dmi-l{font-size:12px;font-weight:600}.dmi-s{font-size:9px;color:#9e9e9e}
.buscar{padding:7px 12px;background:#fff;border-bottom:1px solid #e8eaf6;flex-shrink:0}
.buscar input{width:100%;padding:6px 12px;border:1px solid #c5cae9;border-radius:20px;font-size:12px;outline:none}
.buscar input:focus{border-color:#1a237e}
.filtros{padding:5px 12px;background:#fff;border-bottom:1px solid #e8eaf6;display:flex;gap:3px;flex-wrap:wrap;flex-shrink:0}
.flt{padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;border:1.5px solid #c5cae9;background:#fff;cursor:pointer;color:#555}
.flt.on{background:#1a237e;color:#fff;border-color:#1a237e}
.lista{overflow-y:auto;flex:1}
.card{padding:9px 14px;background:#fff;border-bottom:1px solid #f0f2ff;display:flex;align-items:center;gap:9px}
.card:hover{background:#f0f2ff}
.av{width:34px;height:34px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;flex-shrink:0}
.ci{flex:1;min-width:0;cursor:pointer}
.cn{font-size:12px;font-weight:700;color:#1a237e;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cid{font-size:9px;color:#9e9e9e;margin-top:1px}
.cs{display:flex;gap:3px;margin-top:3px;flex-wrap:wrap}
.tag{padding:1px 6px;border-radius:8px;font-size:9px;font-weight:700}
.ca{display:flex;gap:5px;flex-shrink:0}
.ico-btn{width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;text-decoration:none;border:none;cursor:pointer;flex-shrink:0}
.wa-b{background:#25d366;color:#fff}.wa-b:hover{background:#1da851}
.exp-b{background:#e8eaf6;color:#1a237e}.exp-b:hover{background:#c5cae9}
/* Detalle expandible */
.detail{display:none;padding:10px 14px;background:#f8f9ff;border-bottom:2px solid #e8eaf6}
.detail.open{display:block}
.drow{display:flex;justify-content:space-between;padding:2px 0;font-size:11px}
.dl{color:#888}.dv{color:#333;font-weight:500;text-align:right;max-width:60%;word-break:break-word}
.dsec{font-size:9px;font-weight:700;color:#9e9e9e;text-transform:uppercase;letter-spacing:.5px;margin:7px 0 3px}
.dtxt{font-size:11px;color:#444;line-height:1.5;margin-top:3px}
.d-btns{display:flex;gap:6px;margin-top:8px;flex-wrap:wrap}
.d-btn{padding:7px 12px;border-radius:5px;font-size:11px;font-weight:600;border:none;cursor:pointer;text-decoration:none;display:inline-block;text-align:center}
/* Derivaciones */
.drv-card{padding:9px 14px;background:#fff;border-bottom:1px solid #f0f2ff;display:flex;align-items:center;gap:8px}
.drv-card:hover{background:#f0f2ff}
.drv-estado{padding:2px 7px;border-radius:8px;font-size:10px;font-weight:700}
.drv-ok{background:#d9ead3;color:#274e13}.drv-pend{background:#fff2cc;color:#7f6000}
.drv-info{flex:1;min-width:0}
.drv-nom{font-size:12px;font-weight:700;color:#1a237e}
.drv-sub{font-size:10px;color:#9e9e9e;margin-top:1px}
.empty{padding:40px 20px;text-align:center;color:#9e9e9e}
.spin{font-size:22px;animation:sp 1s linear infinite;display:inline-block}
@keyframes sp{to{transform:rotate(360deg)}}
/* Modal sesión */
.ses-overlay{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;display:flex;align-items:center;justify-content:center}
.ses-box{background:#fff;border-radius:10px;padding:18px;width:320px;max-width:95vw;box-shadow:0 8px 32px rgba(0,0,0,.25)}
.ses-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px}
.ses-hdr strong{color:#1a237e;font-size:13px}.ses-x{border:none;background:none;font-size:18px;cursor:pointer;color:#888}
.ses-lbl{font-size:11px;font-weight:700;color:#333;display:block;margin-bottom:3px;margin-top:9px}
.ses-inp{width:100%;padding:7px;border:1px solid #c5cae9;border-radius:4px;font-size:11px}
.ses-ta{width:100%;padding:7px;border:1px solid #c5cae9;border-radius:4px;font-size:11px;height:60px;resize:none}
.ses-ok{width:100%;padding:9px;background:#7e57c2;color:#fff;border:none;border-radius:5px;font-size:12px;font-weight:700;cursor:pointer;margin-top:12px}
</style></head>
<body onclick="cerrarDD(event)">

<!-- ═══ TOPBAR ═══ -->
<div class="tb">
  <h2>👤 Paso a Paso</h2>
  <span class="badge" id="cnt">Cargando…</span>
</div>

<!-- ═══ TABS ═══ -->
<div class="tabs">
  <div class="tab on" id="tab-p" onclick="cambiarTab('p')">👤 Participantes</div>
  <div class="tab" id="tab-d" onclick="cambiarTab('d')">📋 Derivaciones</div>
</div>

<!-- ═══ VISTA PARTICIPANTES ═══ -->
<div id="vista-p" style="display:flex;flex-direction:column;flex:1;overflow:hidden">
  <div class="accbar">
    <div class="dd" id="dd">
      <button class="btn btn-p" onclick="togDD(event)">⚡ Gestión ▾</button>
      <div class="ddm" id="ddm">
        <div class="dmi" onclick="run('agregarParticipanteManual')"><span class="dmi-ico">➕</span><div class="dmi-t"><span class="dmi-l">Agregar Participante</span><span class="dmi-s">Registrar nuevo manualmente</span></div></div>
        <div class="dmi" onclick="run('editarParticipante')"><span class="dmi-ico">✏️</span><div class="dmi-t"><span class="dmi-l">Editar Datos</span><span class="dmi-s">Modificar datos del participante</span></div></div>
        <div class="dmi" onclick="run('abrirFormDerivacion')"><span class="dmi-ico">➡️</span><div class="dmi-t"><span class="dmi-l">Derivar a Otro Programa</span><span class="dmi-s">Enviar a otra institución</span></div></div>
      </div>
    </div>
    <button class="btn btn-g" onclick="recargar()">🔄 Actualizar</button>
  </div>
  <div class="buscar">
    <input id="q" placeholder="🔍 Buscar por nombre, Creamos ID, zona…" oninput="fil()" autofocus>
  </div>
  <div class="filtros">
    <button class="flt on" onclick="setF(this,'','')">Todos</button>
    <button class="flt" onclick="setF(this,'p','Perfil A')" style="color:#274e13;border-color:#a8d5a2">A</button>
    <button class="flt" onclick="setF(this,'p','Perfil B')" style="color:#1c4587;border-color:#9bbfe0">B</button>
    <button class="flt" onclick="setF(this,'p','Perfil C')" style="color:#7f6000;border-color:#f0d060">C</button>
    <button class="flt" onclick="setF(this,'p','Perfil D')" style="color:#660000;border-color:#e08080">D</button>
    <button class="flt" onclick="setF(this,'e','Orientación')">Orient.</button>
    <button class="flt" onclick="setF(this,'e','Mentoría')">Ment.</button>
    <button class="flt" onclick="setF(this,'e','Formación')">Form.</button>
    <button class="flt" onclick="setF(this,'e','Inactivo')">Inact.</button>
  </div>
  <div class="lista" id="lista">
    <div class="empty"><div class="spin">⏳</div><br>Cargando participantes…</div>
  </div>
</div>

<!-- ═══ VISTA DERIVACIONES ═══ -->
<div id="vista-d" style="display:none;flex-direction:column;flex:1;overflow:hidden">
  <div class="accbar">
    <a class="btn btn-p" href="https://ee.kobotoolbox.org/x/M9M734If" target="_blank">📝 Abrir Formulario Paso a Paso</a>
    <button class="btn btn-g" onclick="cargarDeriv()">🔄 Actualizar</button>
  </div>
  <div class="lista" id="lista-d">
    <div class="empty"><div class="spin">⏳</div><br>Cargando derivaciones…</div>
  </div>
</div>

<script>
var todos=[], derivs=[], tabActual='p', ft='', fv='', abierto=null;
/*__DATOS__*/
var COLS={'Perfil A':{bg:'#274e13',fg:'#d9ead3'},'Perfil B':{bg:'#1c4587',fg:'#cfe2f3'},'Perfil C':{bg:'#7f6000',fg:'#fff2cc'},'Perfil D':{bg:'#660000',fg:'#f4cccc'}};
function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function tel502(t){var s=(t||'').replace(/\\D/g,'');return s.length===8?'502'+s:s;}
function waMens(nom){return encodeURIComponent('Hola '+nom+', somos el equipo de Paso a Paso de Creamos Guatemala. ¿Cómo estás? Nos gustaría ponernos en contacto contigo.');}

/* ── Carga participantes ── */
function iniciar(){
  // Datos inyectados por verFicha() al servir el HTML
  document.getElementById('cnt').textContent=todos.length+' participantes';
  fil();
}
function recargar(){
  abierto=null;
  document.getElementById('lista').innerHTML='<div class="empty"><div class="spin">⏳</div><br>Actualizando…</div>';
  google.script.run
    .withSuccessHandler(function(d){
      todos=d||[];
      document.getElementById('cnt').textContent=todos.length+' participantes';
      fil();
    })
    .withFailureHandler(function(e){
      document.getElementById('lista').innerHTML='<div class="empty" style="color:#c62828">❌ '+esc(String(e.message||e))+'<br><button onclick="recargar()" style="margin-top:8px;padding:5px 12px;background:#1a237e;color:#fff;border:none;border-radius:4px;cursor:pointer">🔄 Reintentar</button></div>';
    })
    .obtenerParticipantes();
}
/* ── Filtros & render participantes ── */
function fil(){
  var q=document.getElementById('q').value.toLowerCase();
  var r=todos.filter(function(p){
    if(q&&!(p.nombre.toLowerCase().includes(q)||(p.id||'').toLowerCase().includes(q)||(p.zona||'').toLowerCase().includes(q)))return false;
    if(ft==='p'&&p.perfil!==fv)return false;
    if(ft==='e'&&p.estado!==fv)return false;
    return true;
  });
  renderLista(r);
}
function setF(b,t,v){
  ft=t;fv=v;
  document.querySelectorAll('.flt').forEach(function(x){x.classList.remove('on');});
  b.classList.add('on');
  fil();
}
function renderLista(r){
  abierto=null;
  var el=document.getElementById('lista');
  if(!r.length){el.innerHTML='<div class="empty">Sin resultados</div>';return;}
  el.innerHTML=r.map(function(p,i){
    var c=COLS[p.perfil]||{bg:'#555',fg:'#f0f2ff'};
    var ini=p.nombre.split(' ').slice(0,2).map(function(w){return(w[0]||'').toUpperCase();}).join('');
    var t=tel502(p.telefono);
    var waBtn=t?'<a class="ico-btn wa-b" href="https://wa.me/'+t+'?text='+waMens(p.nombre)+'" target="_blank" title="WhatsApp">💬</a>':'<span style="width:28px"></span>';
    var expBtn=p.docUrl?'<a class="ico-btn exp-b" href="'+p.docUrl+'" target="_blank" title="Expediente">📄</a>':'';
    var urgTag=p.prioridad==='CRÍTICO'?'<span class="tag" style="background:#c62828;color:#fff">URGENTE</span>':'';
    return '<div id="c'+i+'">'+
      '<div class="card">'+
      '<div class="av" style="background:'+c.bg+';color:'+c.fg+'">'+ini+'</div>'+
      '<div class="ci" onclick="toggleDetail('+i+',this.closest(\'.card\').parentElement,'+JSON.stringify(p).replace(/</g,'\\u003c')+')">'+
        '<div class="cn">'+esc(p.nombre)+'</div>'+
        '<div class="cid">Creamos ID: '+esc(p.id)+'</div>'+
        '<div class="cs">'+
          '<span class="tag" style="background:'+c.fg+';color:'+c.bg+'">'+esc(p.perfil||'Sin perfil')+'</span>'+
          '<span class="tag" style="background:#e8eaf6;color:#3949ab">'+esc(p.estado||'—')+'</span>'+
          urgTag+
        '</div>'+
      '</div>'+
      '<div class="ca">'+waBtn+expBtn+'</div>'+
      '</div>'+
      '<div class="detail" id="det'+i+'"></div>'+
      '</div>';
  }).join('');
}

/* ── Detalle expandible ── */
function toggleDetail(i,wrap,p){
  var det=document.getElementById('det'+i);
  if(det.classList.contains('open')){det.classList.remove('open');abierto=null;return;}
  if(abierto!==null){var prev=document.getElementById('det'+abierto);if(prev)prev.classList.remove('open');}
  abierto=i;
  var t=tel502(p.telefono);
  var waMsg=t?'<a class="d-btn btn-wa" href="https://wa.me/'+t+'?text='+waMens(p.nombre)+'" target="_blank">💬 Enviar WhatsApp</a>':'';
  var sesBtn='<button class="d-btn btn-p" onclick="abrirSes(\''+esc(p.id)+'\',\''+esc(p.nombre)+'\')">📋 Registrar Sesión</button>';
  var docBtn=p.docUrl?'<a class="d-btn" style="background:#e8eaf6;color:#1a237e" href="'+p.docUrl+'" target="_blank">📄 Ver Expediente</a>':'';
  det.innerHTML=
    '<div class="dsec">Datos personales</div>'+
    '<div class="drow"><span class="dl">Creamos ID</span><span class="dv">'+esc(p.id)+'</span></div>'+
    '<div class="drow"><span class="dl">DPI</span><span class="dv">'+esc(p.dpi)+'</span></div>'+
    '<div class="drow"><span class="dl">Edad</span><span class="dv">'+esc(p.edad)+'</span></div>'+
    '<div class="drow"><span class="dl">Género</span><span class="dv">'+esc(p.genero)+'</span></div>'+
    '<div class="drow"><span class="dl">Teléfono</span><span class="dv">'+esc(p.telefono)+'</span></div>'+
    '<div class="drow"><span class="dl">Zona</span><span class="dv">'+esc(p.zona)+'</span></div>'+
    '<div class="drow"><span class="dl">Email</span><span class="dv">'+esc(p.email)+'</span></div>'+
    '<div class="dsec">Perfil profesional</div>'+
    '<div class="drow"><span class="dl">Educación</span><span class="dv">'+esc(p.educacion)+'</span></div>'+
    '<div class="drow"><span class="dl">Situación laboral</span><span class="dv">'+esc(p.laboral)+'</span></div>'+
    '<div class="dsec">Fortalezas</div><div class="dtxt">'+esc(p.fortalezas||'—')+'</div>'+
    '<div class="dsec">Objetivo laboral</div><div class="dtxt">'+esc(p.objetivo||'—')+'</div>'+
    '<div class="dsec">Puntaje diagnóstico: '+p.puntaje+' / 60</div>'+
    '<div class="d-btns">'+waMsg+sesBtn+docBtn+'</div>';
  det.classList.add('open');
}

/* ── Tabs ── */
function cambiarTab(t){
  tabActual=t;
  document.getElementById('tab-p').classList.toggle('on',t==='p');
  document.getElementById('tab-d').classList.toggle('on',t==='d');
  document.getElementById('vista-p').style.display=t==='p'?'flex':'none';
  document.getElementById('vista-d').style.display=t==='d'?'flex':'none';
  if(t==='d'&&derivs.length===0)cargarDeriv();
}

/* ── Derivaciones ── */
function cargarDeriv(){
  document.getElementById('lista-d').innerHTML='<div class="empty"><div class="spin">⏳</div><br>Cargando…</div>';
  google.script.run
    .withSuccessHandler(function(d){
      derivs=d||[];
      renderDeriv();
    })
    .withFailureHandler(function(e){
      document.getElementById('lista-d').innerHTML='<div class="empty" style="color:#c62828">❌ '+esc(String(e.message||e))+'</div>';
    })
    .obtenerDerivados();
}
function renderDeriv(){
  var el=document.getElementById('lista-d');
  if(!derivs.length){
    el.innerHTML='<div class="empty">📭 Sin derivaciones<br><small>Importa desde DP_Empleabilidad usando 🔄 Importar Todo</small></div>';
    return;
  }
  el.innerHTML=derivs.map(function(d){
    var enMaestro=d.enMaestro;
    var t=tel502(d.telefono);
    var waMsg=encodeURIComponent('Hola '+d.nombre+', somos el equipo de Paso a Paso de Creamos Guatemala. Te invitamos a llenar el formulario de inscripción para ingresar al programa.');
    var waForm=encodeURIComponent('Hola '+d.nombre+', somos Paso a Paso de Creamos Guatemala. Aquí el enlace para inscribirte: https://ee.kobotoolbox.org/x/M9M734If');
    return '<div class="drv-card">'+
      '<div class="drv-info">'+
        '<div class="drv-nom">'+esc(d.nombre)+'</div>'+
        '<div class="drv-sub">'+esc(d.telefono||'Sin teléfono')+' · '+esc(d.formacion||'—')+'</div>'+
      '</div>'+
      (enMaestro?'<span class="drv-estado drv-ok">✅ En Maestro</span>':'<span class="drv-estado drv-pend">⏳ Pendiente</span>')+
      (t?'<a class="ico-btn wa-b" href="https://wa.me/'+t+'?text='+waMsg+'" target="_blank" title="WhatsApp seguimiento">💬</a>':'')+
      (t?'<a class="ico-btn" style="background:#e8f5e9;color:#2e7d32;width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;text-decoration:none;font-size:12px" href="https://wa.me/'+t+'?text='+waForm+'" target="_blank" title="Enviar formulario por WhatsApp">📝</a>':'')+
      '</div>';
  }).join('');
}

/* ── Sesión ── */
function abrirSes(pid,pnom){
  var hoy=new Date().toISOString().slice(0,10);
  var ov=document.createElement('div');ov.className='ses-overlay';ov.id='sesOv';
  ov.onclick=function(e){if(e.target===ov)ov.remove();};
  ov.innerHTML='<div class="ses-box">'+
    '<div class="ses-hdr"><strong>📋 Registrar Sesión</strong><button class="ses-x" onclick="document.getElementById(\'sesOv\').remove()">×</button></div>'+
    '<div style="font-size:11px;background:#e8eaf6;padding:6px 10px;border-radius:4px;margin-bottom:4px"><strong>Participante:</strong> '+pnom+'</div>'+
    '<label class="ses-lbl">Fecha</label><input id="sf" class="ses-inp" type="date" value="'+hoy+'">'+
    '<label class="ses-lbl">Tipo de sesión</label>'+
    '<select id="st" class="ses-inp"><option>Orientación Laboral</option><option>Mentoría Individual</option><option>Seguimiento</option><option>Taller/Actividad</option><option>Derivación</option><option>Otro</option></select>'+
    '<label class="ses-lbl">Notas (opcional)</label><textarea id="sn" class="ses-ta" placeholder="Resumen..."></textarea>'+
    '<button class="ses-ok" onclick="guardarSes(\''+pid+'\',\''+pnom+'\')">✅ Guardar Sesión</button>'+
    '</div>';
  document.body.appendChild(ov);
}
function guardarSes(pid,pnom){
  var f=document.getElementById('sf').value;
  var t=document.getElementById('st').value;
  var n=document.getElementById('sn').value;
  if(!f||!t){alert('Completa fecha y tipo');return;}
  var ov=document.getElementById('sesOv');if(ov)ov.remove();
  google.script.run
    .withFailureHandler(function(e){alert('❌ Error: '+e.message);})
    .guardarSesionInterna(pid,pnom,f,t,n);
}

/* ── Dropdown ── */
function togDD(e){e.stopPropagation();document.getElementById('ddm').classList.toggle('open');}
function cerrarDD(e){var dd=document.getElementById('dd');if(dd&&!dd.contains(e.target))document.getElementById('ddm').classList.remove('open');}
function run(fn){document.getElementById('ddm').classList.remove('open');google.script.run[fn]();}

iniciar();
<div id="perfilModal" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.4); z-index:1000; align-items:center; justify-content:center;">
  <div style="background:#fff; border-radius:8px; width:90%; max-width:400px; box-shadow:0 4px 16px rgba(0,0,0,0.2);">
    <div style="padding:12px 16px; background:#1a237e; color:#fff; border-radius:8px 8px 0 0; display:flex; justify-content:space-between; align-items:center;">
      <strong>Perfil del Participante</strong>
      <button onclick="cerrarPerfil()" style="background:none; border:none; color:#fff; font-size:18px; cursor:pointer;">&times;</button>
    </div>
    <div id="perfilContenido"></div>
  </div>
</div>
<script>
document.getElementById('perfilModal').addEventListener('click', function(e) {
  if (e.target === this) cerrarPerfil();
});
</script>
</script></body></html>`;

// ============================================================================
// SESIONES DE ACOMPAÑAMIENTO (registro interno)
// ============================================================================

function guardarSesionInterna(pid, pnom, fecha, tipo, notas) {
  try {
    const ss  = SpreadsheetApp.getActive();
    let hSes  = ss.getSheetByName('Sesiones');
    if (!hSes) {
      hSes = ss.insertSheet('Sesiones');
      hSes.getRange(1,1,1,7).setValues([['Fecha','ID_Participante','Nombre','Tipo_Sesion','Notas','Orientador','Timestamp']])
        .setFontWeight('bold').setBackground('#1a237e').setFontColor('#fff');
      hSes.setFrozenRows(1);
      hSes.setColumnWidth(1,100); hSes.setColumnWidth(2,120); hSes.setColumnWidth(3,180);
      hSes.setColumnWidth(4,160); hSes.setColumnWidth(5,280);
    }
    const orientador = Session.getEffectiveUser().getEmail();
    hSes.appendRow([new Date(fecha), pid, pnom, tipo, notas||'', orientador, new Date()]);
    // Limpiar caché para refrescar datos del participante si es necesario
    CacheService.getScriptCache().remove('p_maestro');
  } catch(e) {
    logError('guardarSesionInterna', e);
    throw e;
  }
}

function verSesionesParticipante() {
  try {
    const ss   = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja || ss.getActiveRange().getRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona primero un participante en la hoja Maestro.');
      return;
    }
    const id   = hoja.getRange(ss.getActiveRange().getRow(), CONFIG.COL.ID).getValue();
    const nom  = hoja.getRange(ss.getActiveRange().getRow(), CONFIG.COL.NOMBRE).getValue();
    const hSes = ss.getSheetByName('Sesiones');
    if (!hSes || hSes.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('📋 No hay sesiones registradas todavía.\n\nUsa "Ver Ficha" → botón "Registrar Sesión" para agregar la primera.');
      return;
    }
    const datos = hSes.getRange(2,1,hSes.getLastRow()-1,7).getValues()
      .filter(r => String(r[1]) === String(id));
    if (!datos.length) {
      SpreadsheetApp.getUi().alert('📋 '+nom+' no tiene sesiones registradas aún.');
      return;
    }
    let msg = '📋 SESIONES — ' + nom + ' (' + datos.length + ' sesiones)\n';
    msg += '═══════════════════════════════════\n\n';
    datos.sort((a,b)=>new Date(b[0])-new Date(a[0])).forEach(r => {
      msg += Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
      msg += '  ·  ' + r[3] + '\n';
      if (r[4]) msg += '    ' + String(r[4]).substring(0,80) + '\n';
      msg += '\n';
    });
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e.message); }
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
  ['Maestro','Derivados','Derivaciones','Sesiones','Log','Dashboard','Analytics',
   'Fuentes_Externas','PBI_Participantes','PBI_Indicadores'].forEach(n => {
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
// DIAGNÓSTICO DE CAMPOS KOBO & CALIDAD DE DATOS
// ============================================================================

function diagnosticoCamposKobo() {
  try {
    const key = getKoboKey();
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json&limit=1',
      { headers: { Authorization: 'Token ' + key }, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert('❌ Error Kobo: ' + resp.getResponseCode()); return;
    }
    const primer = JSON.parse(resp.getContentText()).results[0] || {};
    const campos = Object.keys(primer).filter(k => !k.startsWith('_')).sort();

    const CRITICOS = ['nombre', 'dpi', 'telefono', 'edad', 'genero', 'email'];
    let reporte = '🔬 CAMPOS DISPONIBLES EN KOBO\n══════════════════════════════\n\n';
    reporte += '🔵 CRÍTICOS:\n';
    CRITICOS.forEach(c => {
      const match = campos.find(k => k.toLowerCase().includes(c));
      reporte += `  ${c.padEnd(10)} → ${match || '❌ NO ENCONTRADO'}\n`;
    });
    reporte += `\n🟢 TODOS LOS CAMPOS (${campos.length}):\n`;
    campos.forEach(c => { reporte += `  • ${c}\n`; });

    const ui = SpreadsheetApp.getUi();
    if (ui.alert(reporte, ui.ButtonSet.OK_CANCEL) === ui.Button.OK) {
      const ss = SpreadsheetApp.getActive();
      let log = ss.getSheetByName('Log');
      if (!log) { log = ss.insertSheet('Log'); log.appendRow(['Timestamp','Tipo','Detalle','Usuario']); }
      log.appendRow([new Date(), '🔬 DIAGNÓSTICO KOBO', reporte, Session.getEffectiveUser().getEmail()]);
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('diagnosticoCamposKobo', e);
  }
}

function analizarCalidadDatos() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ No hay datos para analizar'); return;
    }
    const M = CONFIG.COL;
    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 10).getValues();
    const stats = { total: datos.length, nombre: 0, dpi: 0, tel: 0, email: 0, completos: 0, incompletos: [] };

    datos.forEach((f, i) => {
      if (f[M.NOMBRE-1]) stats.nombre++;
      if (f[M.DPI-1])    stats.dpi++;
      if (f[M.TELEFONO-1]) stats.tel++;
      if (f[M.EMAIL-1])  stats.email++;
      if (f[M.NOMBRE-1] && f[M.DPI-1] && f[M.TELEFONO-1]) {
        stats.completos++;
      } else {
        const falta = [!f[M.NOMBRE-1]?'Nombre':'', !f[M.DPI-1]?'DPI':'', !f[M.TELEFONO-1]?'Teléfono':''].filter(x=>x).join(', ');
        stats.incompletos.push({ id: f[M.ID-1], nombre: f[M.NOMBRE-1], fila: i+2, falta });
      }
    });

    const pct = n => ((n / stats.total) * 100).toFixed(0) + '%';
    let panel = `📊 CALIDAD DE DATOS\n${'═'.repeat(40)}\n\nTotal participantes: ${stats.total}\n\n`;
    panel += `✅ COMPLETITUD:\n`;
    panel += `  Nombre    ${pct(stats.nombre).padStart(5)} (${stats.nombre}/${stats.total})\n`;
    panel += `  DPI       ${pct(stats.dpi).padStart(5)} (${stats.dpi}/${stats.total})\n`;
    panel += `  Teléfono  ${pct(stats.tel).padStart(5)} (${stats.tel}/${stats.total})\n`;
    panel += `  Email     ${pct(stats.email).padStart(5)} (${stats.email}/${stats.total})\n`;
    panel += `\n🟢 Registros completos: ${pct(stats.completos)} (${stats.completos}/${stats.total})\n`;
    if (stats.incompletos.length) {
      panel += `\n🔴 Incompletos (${stats.incompletos.length}):\n`;
      stats.incompletos.slice(0, 8).forEach(r => {
        panel += `  Fila ${r.fila}: ${r.nombre || '(sin nombre)'} — falta: ${r.falta}\n`;
      });
      if (stats.incompletos.length > 8) panel += `  ... y ${stats.incompletos.length - 8} más\n`;
    }
    SpreadsheetApp.getUi().alert(panel);
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('analizarCalidadDatos', e);
  }
}

// ============================================================================
// SINCRONIZAR CON VALIDACIONES
// ============================================================================

function sincronizarConValidaciones(silencioso) {
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ Instala primero.');
      return;
    }

    let apiKey;
    try { apiKey = getKoboKey(); }
    catch (e) { if (!silencioso) SpreadsheetApp.getUi().alert('❌ ' + e.message); return; }

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
    let agregados = 0, rechazados = 0;

    registros.forEach(r => {
      const id = String(r['creamos_id'] || r['_id'] || '');
      if (!id || existentes.has(id)) return;

      const nombre   = r['nombre_completo'] || '';
      const dpi      = r['numero_dpi']     || '';
      const telefono = r['telefono']       || '';

      // VALIDAR CAMPOS CRÍTICOS
      if (!nombre.trim() || !dpi.trim() || !telefono.trim()) {
        rechazados++;
        logError('sincronizarConValidaciones',
          `Rechazado [${id}]: Faltan datos críticos (nombre/dpi/telefono)`);
        return;
      }

      const fila = new Array(26).fill('');
      fila[M.ID-1]         = id;
      fila[M.FECHA-1]      = new Date();
      fila[M.NOMBRE-1]     = nombre;
      fila[M.DPI-1]        = dpi;
      fila[M.EDAD-1]       = r['edad']                        || '';
      fila[M.GENERO-1]     = r['genero']                      || '';
      fila[M.TELEFONO-1]   = telefono;
      fila[M.ZONA-1]       = r['lugar_residencia']            || '';
      fila[M.EMAIL-1]      = r['email']                       || '';
      fila[M.EDUCACION-1]  = r['ultimo_grado_aprobado']       || '';
      fila[M.LABORAL-1]    = r['situacion_laboral_actual']    || '';
      fila[M.FORTALEZAS-1] = r['habilidades_especificas']     || '';
      fila[M.OBJETIVO-1]   = r['objetivo_laboral_especifico'] || '';
      fila[M.PERFIL-1]     = normalizarPerfil(r['perfil_asignado']);
      fila[M.PRIORIDAD-1]  = normalizarPrioridad(r['prioridad_caso']);
      fila[M.PUNTAJE-1]    = Number(r['puntaje_total_60']) || 0;
      fila[M.DIM1-1]       = Number(r['dimension_1_capital_educativo']) || 0;
      fila[M.DIM2-1]       = Number(r['dimension_2_capital_laboral']) || 0;
      fila[M.DIM3-1]       = Number(r['dimension_3_habilidades_digitales']) || 0;
      fila[M.DIM4-1]       = Number(r['dimension_4_claridad_vocacional']) || 0;
      fila[M.DIM5-1]       = Number(r['dimension_5_barreras_estructurales']) || 0;
      fila[M.DIM6-1]       = Number(r['dimension_6_red_apoyo']) || 0;
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
      SpreadsheetApp.getUi().alert(`✅ Sincronización completada\n📥 ${agregados} nuevos\n📊 ${registros.length} total en Kobo${rechazados > 0 ? `\n🔴 ${rechazados} rechazados` : ''}`);
    }
  } catch(e) {
    logError('sincronizarConValidaciones', e);
    if (!silencioso) SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

// ============================================================================
// INSPECTOR DE DATOS KOBO — Descubre qué información realmente envía tu form
// ============================================================================

function inspeccionarPrimerRegistroKobo() {
  try {
    const key = getKoboKey();
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json&limit=1',
      { headers: { Authorization: 'Token ' + key }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert('❌ Error: ' + resp.getResponseCode());
      return;
    }

    const primer = JSON.parse(resp.getContentText()).results[0] || {};

    let reporte = '📋 PRIMER REGISTRO KOBO — QUÉ DATOS TIENES\n';
    reporte += '═════════════════════════════════════════════\n\n';

    // Separar por si tienen datos
    const conDatos = {};
    const sinDatos = [];

    Object.entries(primer).forEach(([clave, valor]) => {
      if (clave.startsWith('_')) return;
      if (valor && valor.toString().trim() !== '') {
        conDatos[clave] = valor;
      } else {
        sinDatos.push(clave);
      }
    });

    reporte += '✅ CAMPOS CON DATOS (' + Object.keys(conDatos).length + '):\n';
    Object.entries(conDatos).forEach(([clave, valor]) => {
      const txt = String(valor).substring(0, 40);
      reporte += `  • ${clave}\n    → ${txt}\n`;
    });

    reporte += '\n❌ CAMPOS VACÍOS (' + sinDatos.length + '):\n';
    sinDatos.slice(0, 10).forEach(c => {
      reporte += `  • ${c}\n`;
    });

    SpreadsheetApp.getUi().alert(reporte);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('inspeccionarPrimerRegistroKobo', e);
  }
}

function compararMapeosKobo() {
  try {
    const key = getKoboKey();
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json&limit=1',
      { headers: { Authorization: 'Token ' + key }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert('❌ Error: ' + resp.getResponseCode());
      return;
    }

    const primer = JSON.parse(resp.getContentText()).results[0] || {};
    const camposReales = Object.keys(primer);

    // Campos que el código BUSCA actualmente
    const MAPEOS_ESPERADOS = {
      'Nombre':     'nombre_completo',
      'DPI':        'numero_dpi',
      'Edad':       'edad',
      'Género':     'genero',
      'Teléfono':   'telefono',
      'Zona':       'lugar_residencia',
      'Email':      'email',
      'Educación':  'ultimo_grado_aprobado',
      'Laboral':    'situacion_laboral_actual',
      'Fortalezas': 'habilidades_especificas',
      'Objetivo':   'objetivo_laboral_especifico',
      'Perfil':     'perfil_asignado',
      'Prioridad':  'prioridad_caso',
      'Puntaje':    'puntaje_total_60',
      'Dim1':       'dimension_1_capital_educativo',
      'Dim2':       'dimension_2_capital_laboral',
      'Dim3':       'dimension_3_habilidades_digitales',
      'Dim4':       'dimension_4_claridad_vocacional',
      'Dim5':       'dimension_5_barreras_estructurales',
      'Dim6':       'dimension_6_red_apoyo'
    };

    let reporte = '⚖️ COMPARACIÓN: CÓDIGO vs TU KOBO\n';
    reporte += '═════════════════════════════════════════════\n\n';

    let ok = 0, falta = 0;

    Object.entries(MAPEOS_ESPERADOS).forEach(([nombreSistema, nombreKobo]) => {
      const existe = camposReales.includes(nombreKobo);
      if (existe) {
        reporte += `✅ ${nombreSistema.padEnd(12)} OK\n`;
        ok++;
      } else {
        reporte += `❌ ${nombreSistema.padEnd(12)} NO EXISTE\n    Busca: "${nombreKobo}"\n`;
        falta++;
      }
    });

    reporte += `\n✅ ${ok} correctos  |  ❌ ${falta} faltantes\n`;

    if (falta > 0) {
      reporte += '\n💡 SOLUCIÓN:\n';
      reporte += '1. Haz click en "📋 Exportar Muestra Kobo"\n';
      reporte += '2. Mira la nueva hoja "INSPECCION_KOBO"\n';
      reporte += '3. Copia los nombres EXACTOS de los campos\n';
      reporte += '4. Actualiza Paso_a_Paso.gs línea ~590\n';
    }

    SpreadsheetApp.getUi().alert(reporte);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('compararMapeosKobo', e);
  }
}

function exportarMuestraKoboASheet() {
  try {
    const key = getKoboKey();
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json&limit=10',
      { headers: { Authorization: 'Token ' + key }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert('❌ Error: ' + resp.getResponseCode());
      return;
    }

    const registros = JSON.parse(resp.getContentText()).results || [];
    const ss = SpreadsheetApp.getActive();

    // Crear nueva hoja
    let hojaIns = ss.getSheetByName('INSPECCION_KOBO');
    if (hojaIns) ss.deleteSheet(hojaIns);
    hojaIns = ss.insertSheet('INSPECCION_KOBO');

    // Obtener campos únicos
    const camposUnicos = new Set();
    registros.forEach(r => {
      Object.keys(r).forEach(k => {
        if (!k.startsWith('_')) camposUnicos.add(k);
      });
    });

    const campos = Array.from(camposUnicos).sort();

    // Encabezados
    hojaIns.appendRow(campos);
    hojaIns.getRange(1, 1, 1, campos.length).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');

    // Datos (primeros 10)
    registros.forEach(r => {
      const fila = campos.map(c => r[c] || '');
      hojaIns.appendRow(fila);
    });

    SpreadsheetApp.getUi().alert(
      `✅ Creada hoja "INSPECCION_KOBO"\n\n` +
      `📊 ${registros.length} registros de muestra\n` +
      `📋 ${campos.length} campos distintos\n\n` +
      `Ahora puedes ver exactamente qué datos tiene tu Kobo.`
    );

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('exportarMuestraKoboASheet', e);
  }
}

// ============================================================================
// ACCIÓN RÁPIDA DESDE LA HOJA — registrar sesión sin abrir modal completo
// ============================================================================

function abrirSesionDesdeHoja(pid, pnom) {
  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8">'+
    '<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:"Segoe UI",sans-serif;font-size:12px;padding:18px;background:#f8f9ff}'+
    'label{display:block;font-weight:700;color:#333;font-size:11px;margin-bottom:3px;margin-top:10px}'+
    'input,select,textarea{width:100%;padding:7px;border:1px solid #c5cae9;border-radius:4px;font-size:11px}'+
    'textarea{height:60px;resize:none}'+
    'button{width:100%;padding:10px;background:#7e57c2;color:#fff;border:none;border-radius:5px;font-size:12px;font-weight:700;cursor:pointer;margin-top:14px}'+
    '.info{background:#e8eaf6;padding:7px 10px;border-radius:5px;font-size:11px;color:#555;margin-bottom:8px}'+
    '</style></head><body>'+
    '<div class="info"><strong>Participante:</strong> '+pnom+'</div>'+
    '<label>Fecha</label><input id="sf" type="date" value="'+new Date().toISOString().slice(0,10)+'">'+
    '<label>Tipo de Sesión</label><select id="st">'+
    '<option>Orientación Laboral</option><option>Mentoría Individual</option><option>Seguimiento</option>'+
    '<option>Taller/Actividad</option><option>Derivación</option><option>Otro</option></select>'+
    '<label>Notas (opcional)</label><textarea id="sn" placeholder="Resumen de la sesión..."></textarea>'+
    '<button onclick="guardar()">✅ Guardar Sesión</button>'+
    '<script>function guardar(){'+
    'var f=document.getElementById("sf").value;'+
    'var t=document.getElementById("st").value;'+
    'var n=document.getElementById("sn").value;'+
    'if(!f||!t){alert("Completa la fecha y tipo");return;}'+
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})'+
    '.withFailureHandler(function(e){alert("Error: "+e.message);})'+
    '.guardarSesionInterna("'+pid+'","'+pnom+'",f,t,n);}'+
    '<\/script>'+
    '</body></html>'
  ).setWidth(320).setHeight(310).setTitle('📋 Registrar Sesión — '+pnom.split(' ')[0]);
  SpreadsheetApp.getUi().showModelessDialog(html, '📋 Registrar Sesión');
}

// ============================================================================
// REINSTALAR COMPLETO — crea todas las hojas, validaciones y columna de acción
// ============================================================================

function reinstalarCompleto() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('📥 Reinstalar / Reparar Todo',
    'Esto crea todas las hojas faltantes y aplica formato y validaciones.\n\nLos datos existentes NO se borran.\n\n¿Continuar?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  const creadas = [];

  // ── Maestro ──────────────────────────────────────────────────────────────
  let maestro = ss.getSheetByName(CONFIG.HOJA);
  if (!maestro) { maestro = ss.insertSheet(CONFIG.HOJA); creadas.push('Maestro'); }
  if (maestro.getLastRow() === 0) {
    const HEADERS = ['ID_Creamos','Fecha_Registro','Nombre_Completo','DPI','Edad','Género',
      'Teléfono','Zona','Email','Nivel_Educativo','Situación_Laboral','Fortalezas',
      'Objetivo_Laboral','Perfil_Asignado','Prioridad','Puntaje_Total',
      'Dim_Educativo','Dim_Laboral','Dim_Digital','Dim_Vocacional',
      'Dim_Barreras','Dim_Apoyo','Estado','Carpeta_Drive_ID','Doc_Perfil_ID','Doc_Perfil_URL'];
    maestro.appendRow(HEADERS);
    maestro.getRange(1,1,1,26).setBackground('#1a237e').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    maestro.setFrozenRows(1);
    maestro.setColumnWidths(1,26,100);
    maestro.setColumnWidth(3,220).setColumnWidth(13,200).setColumnWidth(26,280);
  }
  // Columna 27 — Acción Rápida
  const hdrAcc = maestro.getRange(1, 27);
  hdrAcc.setValue('⚡ Acción Rápida').setBackground('#283593').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
  maestro.setColumnWidth(27, 160);
  const lastDataRow = Math.max(maestro.getLastRow(), 500);
  const rngAcc = maestro.getRange(2, 27, lastDataRow - 1, 1);
  rngAcc.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['💬 WhatsApp', '👤 Ver Perfil', '📋 Registrar Sesión'], true)
    .setAllowInvalid(false).build());

  // Validaciones en columnas Maestro
  const rngEstado = maestro.getRange(2, CONFIG.COL.ESTADO, lastDataRow - 1, 1);
  rngEstado.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Orientación','Mentoría','Formación','Derivación','Inactivo','Cierre','Completado'], true).build());
  const rngPerfil = maestro.getRange(2, CONFIG.COL.PERFIL, lastDataRow - 1, 1);
  rngPerfil.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Perfil A','Perfil B','Perfil C','Perfil D'], true).build());
  const rngPrior = maestro.getRange(2, CONFIG.COL.PRIORIDAD, lastDataRow - 1, 1);
  rngPrior.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['CRÍTICO','ALTO','MEDIO','BAJO'], true).build());
  const rngGen = maestro.getRange(2, CONFIG.COL.GENERO, lastDataRow - 1, 1);
  rngGen.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Femenino','Masculino','No binario','Prefiero no decir'], true).build());

  // ── Derivados ────────────────────────────────────────────────────────────
  {
    let hd = ss.getSheetByName('Derivados');
    if (!hd) {
      hd = ss.insertSheet('Derivados');
      hd.appendRow(['Fecha_Import','ID','Nombre','Teléfono','Género','Edad','Educación','DPI','Formación','Cohorte','Notas','Activo','Estado_Derivado','Fecha_Aprobación','⚡ Acción']);
      hd.getRange(1,1,1,15).setBackground('#880e4f').setFontColor('#fff').setFontWeight('bold');
      hd.setFrozenRows(1);
      creadas.push('Derivados');
    } else {
      // Asegurar columna Acción si no existe
      if (hd.getRange(1,15).getValue() !== '⚡ Acción') {
        hd.getRange(1,15).setValue('⚡ Acción').setBackground('#880e4f').setFontColor('#fff').setFontWeight('bold');
      }
    }
    hd.getRange(2,13,500,1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente formulario','Formulario enviado','Completó formulario','Rechazado'], true).build());
    hd.getRange(2,15,500,1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['📲 Enviar formulario Kobo','🔁 Reenviar recordatorio','✅ Ya completó el formulario'], true).build());
    hd.setColumnWidth(3,180).setColumnWidth(15,200);
  }

  // ── Sesiones ─────────────────────────────────────────────────────────────
  if (!ss.getSheetByName('Sesiones')) {
    const hs = ss.insertSheet('Sesiones');
    hs.appendRow(['Fecha','ID_Participante','Nombre','Tipo_Sesion','Notas','Orientador','Timestamp']);
    hs.getRange(1,1,1,7).setBackground('#4a148c').setFontColor('#fff').setFontWeight('bold');
    hs.setFrozenRows(1);
    hs.setColumnWidth(3,180).setColumnWidth(4,160).setColumnWidth(5,280);
    creadas.push('Sesiones');
  }

  // ── Derivaciones ─────────────────────────────────────────────────────────
  if (!ss.getSheetByName('Derivaciones')) {
    const hder = ss.insertSheet('Derivaciones');
    hder.appendRow(['Fecha','ID_Participante','Nombre','Tipo','Destino','Motivo','Estado','Responsable','Fecha_Seguimiento','Notas']);
    hder.getRange(1,1,1,10).setBackground('#bf360c').setFontColor('#fff').setFontWeight('bold');
    hder.setFrozenRows(1);
    hder.setColumnWidth(3,180).setColumnWidth(5,200).setColumnWidth(6,250);
    creadas.push('Derivaciones');
  }

  // ── Log ──────────────────────────────────────────────────────────────────
  if (!ss.getSheetByName('Log')) {
    const hl = ss.insertSheet('Log');
    hl.appendRow(['Fecha_Hora','Hoja','Fila','Columna','Valor_Anterior','Valor_Nuevo','Usuario']);
    hl.getRange(1,1,1,7).setBackground('#37474f').setFontColor('#fff').setFontWeight('bold');
    hl.setFrozenRows(1);
    creadas.push('Log');
  }

  // ── Dashboard ────────────────────────────────────────────────────────────
  if (!ss.getSheetByName('Dashboard')) {
    const hdb = ss.insertSheet('Dashboard');
    hdb.appendRow(['DASHBOARD — PASO A PASO']);
    hdb.getRange(1,1).setFontSize(16).setFontWeight('bold').setFontColor('#1a237e');
    creadas.push('Dashboard');
  }

  // ── Analytics ────────────────────────────────────────────────────────────
  if (!ss.getSheetByName('Analytics')) {
    const han = ss.insertSheet('Analytics');
    han.appendRow(['ANALYTICS — PASO A PASO']);
    han.getRange(1,1).setFontSize(16).setFontWeight('bold').setFontColor('#6a1b9a');
    han.setFrozenRows(1);
    creadas.push('Analytics');
  }

  // ── Fuentes_Externas ─────────────────────────────────────────────────────
  instalarFuentesExternas(ss);

  // ── Triggers ─────────────────────────────────────────────────────────────
  configurarTriggers();

  // ── Caché limpio ─────────────────────────────────────────────────────────
  try { CacheService.getScriptCache().removeAll(['p_maestro']); } catch(e) {}

  ui.alert('✅ Sistema reparado',
    (creadas.length ? '📋 Hojas creadas: '+creadas.join(', ')+'\n' : '✅ Todas las hojas ya existían\n')+
    '✅ Columna "⚡ Acción Rápida" lista en Maestro\n'+
    '✅ Validaciones aplicadas (Estado, Perfil, Prioridad, Género)\n'+
    '✅ Triggers configurados\n\n'+
    '💡 Usa la columna ⚡ de cada fila para acciones rápidas.',
    ui.ButtonSet.OK);
}

// ============================================================================
// FUENTES EXTERNAS — hoja de configuración de hojas externas
// ============================================================================

function instalarFuentesExternas(ss) {
  ss = ss || SpreadsheetApp.getActive();
  if (ss.getSheetByName('Fuentes_Externas')) return;
  const hf = ss.insertSheet('Fuentes_Externas');
  hf.appendRow(['Nombre','Spreadsheet_ID','Nombre_Hoja','Activo','Última_Sync','Notas']);
  hf.getRange(1,1,1,6).setBackground('#006064').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
  hf.setFrozenRows(1);
  hf.setColumnWidth(1,160).setColumnWidth(2,320).setColumnWidth(3,140).setColumnWidth(4,80).setColumnWidth(5,140).setColumnWidth(6,220);
  // Fila de ejemplo con DP_Empleabilidad
  hf.appendRow(['DP_Empleabilidad', DP_EMPLEABILIDAD.SPREADSHEET_ID, DP_EMPLEABILIDAD.HOJA, 'SÍ', '', 'Fuente principal de derivados']);
  hf.getRange(2,4,500,1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['SÍ','NO'], true).build());
  hf.getRange(2,1,1,6).setBackground('#e0f7fa');
  // Instrucciones en fila 3
  hf.appendRow(['← EJEMPLO — puedes editar','','','','','Agrega una fila por cada fuente externa que quieras importar']);
  hf.getRange(3,1,1,6).setFontColor('#9e9e9e').setFontStyle('italic');
}

function leerFuentesExternas() {
  const ss  = SpreadsheetApp.getActive();
  const hf  = ss.getSheetByName('Fuentes_Externas');
  if (!hf || hf.getLastRow() < 2) return [];
  return hf.getRange(2, 1, hf.getLastRow()-1, 6).getValues()
    .filter(r => String(r[0]).trim() && String(r[1]).trim() && String(r[3]).toUpperCase() === 'SÍ')
    .map(r => ({ nombre: String(r[0]), ssId: String(r[1]), hoja: String(r[2]||''), rowIdx: 0 }));
}

// ============================================================================
// SINCRONIZAR TODO — Kobo + todas las fuentes externas activas
// ============================================================================

function sincronizarTodo() {
  const ui = SpreadsheetApp.getUi();
  const log = [];
  let errores = 0;

  // 1. Kobo
  try {
    sincronizarConValidaciones(true);
    log.push('✅ Kobo sincronizado');
  } catch(e) {
    log.push('⚠️ Kobo: ' + e.message);
    errores++;
  }

  // 2. Derivados (DP_Empleabilidad y fuentes externas activas)
  try {
    importarDPADerivados(true);
    log.push('✅ DP_Empleabilidad → Derivados');
  } catch(e) {
    log.push('⚠️ DP_Empleabilidad: ' + e.message);
    errores++;
  }

  // 3. Otras fuentes externas (Fuentes_Externas sheet)
  const fuentes = leerFuentesExternas();
  const ss = SpreadsheetApp.getActive();
  const hf = ss.getSheetByName('Fuentes_Externas');
  fuentes.forEach(function(f, idx) {
    // Saltar DP_Empleabilidad (ya importada arriba)
    if (String(f.ssId).trim() === String(DP_EMPLEABILIDAD.SPREADSHEET_ID).trim()) return;
    try {
      const extSS = SpreadsheetApp.openById(f.ssId);
      const extH  = f.hoja ? extSS.getSheetByName(f.hoja) : extSS.getSheets()[0];
      if (!extH) throw new Error('Hoja "'+f.hoja+'" no encontrada');
      const filas = extH.getDataRange().getValues();
      log.push('✅ '+f.nombre+': '+( filas.length - 1 )+' registros leídos (integración pendiente de mapeo)');
      // Actualizar timestamp en Fuentes_Externas
      if (hf) hf.getRange(idx+2, 5).setValue(new Date());
    } catch(e) {
      log.push('⚠️ '+f.nombre+': ' + e.message);
      errores++;
    }
  });

  // 4. Limpiar caché
  try { CacheService.getScriptCache().removeAll(['p_maestro']); } catch(e) {}

  // 5. Actualizar PBI si tiene triggers
  try { triggerSilenciosoExportPBI(); } catch(e) {}

  ui.alert(
    errores === 0 ? '✅ Sincronización completa' : '⚠️ Sincronización con advertencias',
    log.join('\n') + '\n\n🕐 ' + new Date().toLocaleString(),
    ui.ButtonSet.OK
  );
}

// ============================================================================
// FIN — v8.7 + Mejoras Inspección
// ============================================================================
