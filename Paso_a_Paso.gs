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
    DIM4:20, DIM5:21, DIM6:22, ESTADO:23, CARPETA_ID:24, DOC_ID:25, DOC_URL:26,
    TIPO_CIERRE:28, FECHA_CIERRE:29, RESULTADO_CIERRE:30
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
      .addItem('📊 Actualizar Dashboard',             'abrirDashboard')
      .addItem('📊 Actualizar Power BI',              'exportarParaPowerBI')
      .addSeparator()
      // ── Automatización ────────────────────────────────────────────────
      .addItem('⏰ Configurar Auto-Sync',              'configurarSyncMatutina')
      .addSeparator()
      // ── Configuración & Setup ─────────────────────────────────────────
      .addItem('📥 Instalar / Reparar Sistema',      'reinstalarCompleto')
      .addItem('🔧 Migrar estructura de datos',       'migrarEstructura')
      .addItem('🧹 Limpiar sesiones duplicadas',      'limpiarSesionesDuplicadas')
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

function actualizarHojaDerivados() {
  const ss = SpreadsheetApp.getActive();
  const hd = ss.getSheetByName('Derivados');
  if (!hd) { SpreadsheetApp.getUi().alert('❌ No existe la hoja Derivados.'); return; }

  const headers = hd.getRange(1, 1, 1, 100).getValues()[0];
  let accionCol = headers.findIndex(h => String(h).includes('Acción'));

  // Si no existe la columna Acción, agregarla al final
  if (accionCol < 0) {
    const lastCol = headers.reduce((acc, v, i) => v ? i + 1 : acc, 0) + 1;
    accionCol = lastCol - 1;
    hd.getRange(1, lastCol).setValue('Acción').setBackground('#880e4f').setFontColor('#fff').setFontWeight('bold');
    hd.setColumnWidth(lastCol, 200);
  }

  // Actualizar dropdown de la columna Acción
  hd.getRange(2, accionCol + 1, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'Enviar formulario Kobo',
        'Recordatorio de sesión agendada',
        'Ya completó el formulario'
      ], true)
      .setAllowInvalid(false)
      .build()
  );

  SpreadsheetApp.getUi().alert('✅ Hoja Derivados actualizada.\n\nEl dropdown "Acción" ahora tiene las 3 opciones:\n• Enviar formulario Kobo\n• Recordatorio de sesión agendada\n• Ya completó el formulario');
}

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
  // DP_Empleabilidad
  try { importarDPADerivados(true); } catch(e) { logError('syncAuto/DP_Emplea', e); }
  // Todas las demás fuentes activas en Fuentes_Externas
  try {
    leerFuentesExternas().forEach(function(f) {
      if (String(f.ssId).trim() === String(DP_EMPLEABILIDAD.SPREADSHEET_ID).trim()) return;
      try { importarFuenteGenerica(f.ssId, f.hoja, f.nombre, true); }
      catch(e) { logError('syncAuto/' + f.nombre, e); }
    });
  } catch(e) { logError('syncAuto/fuentes', e); }
  try { CacheService.getScriptCache().removeAll(['p_maestro']); } catch(e) {}
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
        // 13 cols: Fecha|ID|Nombre|Tipo|Destino|Motivo|Responsable|FechaSeg|Notas|Prioridad|SesionRel|FechaCierre|Estado
        filasDerivaciones.push([
          hoy, id, nombre, '💡 SUGERIDA',
          'Formación Técnica — ' + formacion,
          'Importado de DP_Empleabilidad. Formación: ' + formacion + (cohorte ? ' | Cohorte: ' + cohorte : ''),
          '', '', nota, 'MEDIA', '', '', 'Pendiente'
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
        deriv.getRange(dr, 1, filasDerivaciones.length, 13).setValues(filasDerivaciones);
        deriv.getRange(dr, 1, filasDerivaciones.length, 13).setBackground('#e8f5e9');
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

// ============================================================================
// IMPORTACIÓN GENÉRICA — multi-fuente con auto-detección de columnas
// ============================================================================

/**
 * Normaliza un string para comparación insensible a tildes, mayúsculas y separadores.
 */
function _normCol(s) {
  return String(s||'').toLowerCase()
    .replace(/[áäà]/g,'a').replace(/[éëè]/g,'e').replace(/[íïì]/g,'i')
    .replace(/[óöò]/g,'o').replace(/[úüù]/g,'u')
    .replace(/[\s_\-\.]/g,'');
}

/**
 * Auto-detecta índices de columna (0-based) a partir de los headers de una hoja.
 * Tolerante a tildes, mayúsculas, espacios y variantes de nombre.
 */
function autoDetectarCols(headers) {
  const n = _normCol;
  const find = (...keys) => {
    for (const k of keys) {
      const i = headers.findIndex(h => n(h) === n(k));
      if (i >= 0) return i;
    }
    return -1;
  };
  return {
    ID:        find('ID','CreamosID','Creamos ID','creamos_id','id_participante','codigo','Codigo','code','Carnet'),
    FECHA:     find('Fecha','FechaIngreso','Fecha Ingreso','FechaRegistro','Fecha Registro','fecha_ingreso','Date','Fecha de registro'),
    NOMBRE:    find('Nombre','NombreCompleto','Nombre Completo','Apellidos y Nombre','full name','Participante','nombre_completo'),
    DPI:       find('DPI','NumeroDPI','Numero DPI','numero_dpi','Cedula','Documento','DNI','CUI'),
    EDAD:      find('Edad','Age','Anos','Años','edad_participante'),
    GENERO:    find('Genero','Género','Sexo','Gender','genero_participante'),
    TELEFONO:  find('Telefono','Teléfono','Celular','Movil','Móvil','Phone','Tel','telefono_celular','numero_telefono'),
    EMAIL:     find('Email','Correo','Correo electronico','Correo Electrónico','E-mail','Mail','email_participante'),
    EDUCACION: find('Educacion','Educación','Nivel academico','Nivel Académico','Escolaridad','UltimoGrado','Ultimo Grado','nivel_educativo'),
    FORMACION: find('Formacion','Formación','Programa','TipodeCurso','Tipo de Curso','Curso','Capacitacion','Capacitación','nombre_curso','tipo_formacion'),
    COHORTE:   find('Cohorte','Grupo','Cohort','Generacion','Generación','Ciclo','cohorte_numero','numero_cohorte'),
    NOTA:      find('Notas','Nota','Observaciones','Comentarios','Notes','observacion','nota_adicional'),
    ACTIVO:    find('Activo','Active','Estado','Participando','Status','Inscrito','activo_yn','participando','Vigente'),
  };
}

/**
 * Importa desde cualquier Google Sheet externo → hoja Derivados.
 * Auto-detecta columnas por nombre de encabezado.
 * Retorna {nuevos, omitidos, error}.
 */
function importarFuenteGenerica(ssId, nombreHoja, nombreFuente, silencioso) {
  const ss = SpreadsheetApp.getActive();

  // Garantizar que Derivados existe con headers correctos
  let hDer = ss.getSheetByName('Derivados');
  if (!hDer) {
    importarDPADerivados(true);
    hDer = ss.getSheetByName('Derivados');
    if (!hDer) return { nuevos:0, omitidos:0, error:'No se pudo crear hoja Derivados' };
  }

  // Abrir fuente externa
  let ssExt;
  try { ssExt = SpreadsheetApp.openById(ssId); }
  catch(e) {
    if (!silencioso) SpreadsheetApp.getUi().alert(
      '❌ No se puede abrir "' + nombreFuente + '"\n\n' +
      'ID del Spreadsheet: ' + ssId + '\n\n' +
      'Verifica que tienes acceso a ese Google Sheet.');
    return { nuevos:0, omitidos:0, error: e.message };
  }

  const hSrc = nombreHoja ? ssExt.getSheetByName(nombreHoja) : ssExt.getSheets()[0];
  if (!hSrc) {
    const hojas = ssExt.getSheets().map(h => h.getName()).join(', ');
    if (!silencioso) SpreadsheetApp.getUi().alert(
      '❌ ' + nombreFuente + ': No existe la hoja "' + nombreHoja + '".\n\n' +
      'Hojas disponibles en ese archivo: ' + hojas);
    return { nuevos:0, omitidos:0, error: 'Hoja no encontrada' };
  }
  if (hSrc.getLastRow() < 2) return { nuevos:0, omitidos:0, error: null };

  // Leer todas las filas de una vez
  const lastCol    = Math.max(hSrc.getLastColumn(), 5);
  const todasFilas = hSrc.getRange(1, 1, hSrc.getLastRow(), lastCol).getValues();
  const headers    = todasFilas[0];
  const C = autoDetectarCols(headers);

  if (C.NOMBRE < 0) {
    if (!silencioso) SpreadsheetApp.getUi().alert(
      '❌ ' + nombreFuente + ': No se encontró columna de Nombre.\n\n' +
      'Headers detectados: ' + headers.filter(Boolean).slice(0,12).join(' | ') + '\n\n' +
      'La hoja necesita una columna llamada "Nombre", "Nombre Completo" o similar.');
    return { nuevos:0, omitidos:0, error: 'Sin columna Nombre' };
  }

  // IDs ya existentes (Derivados + Maestro) para evitar duplicados
  const idsExist = new Set();
  const maestro  = ss.getSheetByName(CONFIG.HOJA);
  if (maestro && maestro.getLastRow() > 1)
    maestro.getRange(2, CONFIG.COL.ID, maestro.getLastRow()-1, 1).getValues()
      .forEach(r => { if (r[0]) idsExist.add(String(r[0]).trim()); });
  if (hDer.getLastRow() > 1)
    hDer.getRange(2, COL_DER.ID, hDer.getLastRow()-1, 1).getValues()
      .forEach(r => { if (r[0]) idsExist.add(String(r[0]).trim()); });

  const hoy    = new Date();
  const tz     = Session.getScriptTimeZone();
  const prefix = nombreFuente.replace(/[^A-Za-z0-9]/g,'').substring(0,3).toUpperCase();
  const nuevas = [];

  todasFilas.slice(1).forEach(fila => {
    const nombre = String(C.NOMBRE >= 0 ? fila[C.NOMBRE] : '').trim();
    if (!nombre) return;

    // Usar ID de la fuente o generar uno con prefijo de la fuente
    let id = String(C.ID >= 0 ? fila[C.ID] : '').trim();
    if (!id) id = prefix + nombre.replace(/\s+/g,'').substring(0,4).toUpperCase()
                + Utilities.formatDate(hoy, tz, 'ddMMyyyy');
    if (idsExist.has(id)) return;
    idsExist.add(id);

    const activoRaw = String(C.ACTIVO >= 0 ? fila[C.ACTIVO] : '').toLowerCase().trim();
    const activo    = activoRaw === '' ||
      ['true','si','sí','1','activo','activa','yes','inscrito','inscrita','participando','vigente'].includes(activoRaw);
    const fecha     = C.FECHA >= 0 && fila[C.FECHA] instanceof Date ? fila[C.FECHA] : hoy;

    // Construir fila en el orden de COL_DER (14 columnas)
    nuevas.push([
      id,                                                                     // 1 ID
      fecha,                                                                  // 2 Fecha Orig.
      nombre,                                                                 // 3 Nombre
      String(C.DPI       >= 0 ? fila[C.DPI]       : '').trim(),              // 4 DPI
      C.EDAD      >= 0 ? (fila[C.EDAD] || '') : '',                          // 5 Edad
      String(C.GENERO    >= 0 ? fila[C.GENERO]    : '').trim(),              // 6 Género
      String(C.TELEFONO  >= 0 ? fila[C.TELEFONO]  : '').trim(),              // 7 Teléfono
      String(C.EDUCACION >= 0 ? fila[C.EDUCACION] : '').trim(),              // 8 Educación
      String(C.FORMACION >= 0 ? fila[C.FORMACION] : '').trim(),              // 9 Formación
      String(C.COHORTE   >= 0 ? fila[C.COHORTE]   : '').trim(),              // 10 Cohorte
      String(C.NOTA      >= 0 ? fila[C.NOTA]       : '').trim(),              // 11 Notas
      activo ? 'Pendiente formulario' : 'Inactivo',                          // 12 Estado Derivado
      nombreFuente,                                                           // 13 Fuente
      hoy                                                                     // 14 Fecha Importación
    ]);
  });

  if (nuevas.length) {
    const startRow = hDer.getLastRow() + 1;
    hDer.getRange(startRow, 1, nuevas.length, NCOLS_DER).setValues(nuevas);
    // Color diferente por fuente para identificar origen visualmente
    const coloresFuente = { 'DP_Empleabilidad':'#fff8e1', 'AYB':'#e8f5e9', 'TECH':'#e3f2fd' };
    const bg = coloresFuente[nombreFuente] || '#f3e5f5';
    hDer.getRange(startRow, 1, nuevas.length, NCOLS_DER).setBackground(bg);
    SpreadsheetApp.flush();

    // Registrar timestamp en Fuentes_Externas
    try {
      const hfe = ss.getSheetByName('Fuentes_Externas');
      if (hfe && hfe.getLastRow() > 1) {
        hfe.getRange(2, 1, hfe.getLastRow()-1, 5).getValues().forEach((r, i) => {
          if (String(r[1]).trim() === String(ssId).trim())
            hfe.getRange(i + 2, 5).setValue(new Date());
        });
      }
    } catch(e) {}
  }

  return { nuevos: nuevas.length, omitidos: todasFilas.length - 1 - nuevas.length, error: null };
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
    body.appendParagraph('Creamos ID: ' + id);
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

      const headers = sheet.getRange(1, 1, 1, 100).getValues()[0];
      const accionCol = headers.findIndex(h => String(h).includes('Acción'));
      if (fila < 2 || col !== accionCol + 1 || !valN) return;
      e.range.clearContent();

      const rowData  = sheet.getRange(fila, 1, 1, 100).getValues()[0];
      const nombre   = String(rowData[headers.findIndex(h => String(h).includes('Nombre'))] || '');
      const telRaw   = String(rowData[headers.findIndex(h => String(h).includes('Teléfono'))] || '').replace(/\D/g,'');
      const tel      = telRaw.length === 8 ? '502' + telRaw : telRaw;
      const estadoIdx = headers.findIndex(h => String(h).includes('Estado'));
      const KOBO_FORM = 'https://ee.kobotoolbox.org/x/M9M734If';

      // ── OPCIÓN 1: Mostrar link del formulario (presencial) ──────────────
      if (valN.includes('Enviar formulario')) {
        const html = HtmlService.createHtmlOutput(
          '<div style="font-family:\'Segoe UI\',sans-serif;padding:22px;text-align:center;">' +
          '<div style="font-size:38px;margin-bottom:10px;">📋</div>' +
          '<p style="font-size:13px;color:#666;margin:0 0 4px 0;">Formulario de inscripción para</p>' +
          '<p style="font-size:15px;font-weight:700;color:#1a237e;margin:0 0 18px 0;">' + nombre + '</p>' +
          '<div style="background:#f5f5f5;border-left:3px solid #1a237e;padding:10px 14px;border-radius:0 6px 6px 0;margin-bottom:16px;text-align:left;">' +
          '<p style="margin:0 0 6px 0;font-size:9px;color:#999;font-weight:700;letter-spacing:.5px;text-transform:uppercase;">Link del formulario</p>' +
          '<code style="font-size:10px;color:#1a237e;word-break:break-all;">' + KOBO_FORM + '</code>' +
          '</div>' +
          '<p style="font-size:10px;color:#777;margin:0 0 16px;line-height:1.5;">Muéstrale este link a <strong>' + nombre + '</strong> para que lo llene en tu teléfono o compártelo por WhatsApp/SMS.</p>' +
          '<div style="display:flex;gap:8px;margin-bottom:8px;">' +
          '<button onclick="navigator.clipboard.writeText(\'' + KOBO_FORM + '\').then(()=>{this.textContent=\'✅ Copiado\';}).catch(()=>{})" style="flex:1;padding:9px;background:#1a237e;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:700;font-size:11px;">📋 Copiar link</button>' +
          '<a href="' + KOBO_FORM + '" target="_blank" style="flex:1;padding:9px;background:#1e88e5;color:#fff;border-radius:6px;text-decoration:none;font-weight:700;font-size:11px;display:flex;align-items:center;justify-content:center;">🌐 Abrir form.</a>' +
          '</div>' +
          (tel ? '<a href="https://wa.me/' + tel + '?text=' + encodeURIComponent('Hola ' + nombre + ', te comparto el formulario de Paso a Paso para que lo llenes: ' + KOBO_FORM) + '" target="_blank" style="display:block;padding:9px;background:#25d366;color:#fff;border-radius:6px;text-decoration:none;font-weight:700;font-size:11px;text-align:center;">💬 Enviar por WhatsApp</a>' : '') +
          '</div>'
        ).setWidth(340).setHeight(380).setTitle('Formulario Kobo — ' + nombre);
        SpreadsheetApp.getUi().showModelessDialog(html, 'Formulario Kobo');
        if (estadoIdx >= 0) sheet.getRange(fila, estadoIdx + 1).setValue('Formulario enviado');

      // ── OPCIÓN 2: Recordatorio por WhatsApp ────────────────────────────
      } else if (valN.includes('Recordatorio') || valN.includes('sesión')) {
        if (!tel) {
          SpreadsheetApp.getUi().alert('⚠️ Sin teléfono\n\n' + nombre + ' no tiene número registrado.\nAbre el formulario con la opción "Enviar formulario Kobo" para compartirlo presencialmente.');
          return;
        }
        const msg = encodeURIComponent(
          'Hola ' + nombre + ' 👋\n\n' +
          '📅 *Recordatorio de sesión agendada*\n\n' +
          'Te escribimos del equipo de Paso a Paso - Creamos Guatemala. Queremos recordarte que tienes una sesión agendada con nosotros.\n\n' +
          '✅ *Por favor, confirma tu asistencia*\n\n' +
          'Si tienes alguna duda o necesitas cambiar la hora, no dudes en escribirnos.\n\n' +
          '¡Nos vemos pronto!'
        );
        const html = HtmlService.createHtmlOutput(
          '<div style="font-family:\'Segoe UI\',sans-serif;padding:22px;text-align:center;">' +
          '<div style="font-size:38px;margin-bottom:10px;">📅</div>' +
          '<p style="font-size:13px;color:#666;margin:0 0 4px;">Recordatorio de sesión para</p>' +
          '<p style="font-size:15px;font-weight:700;color:#1a237e;margin:0 0 6px;">' + nombre + '</p>' +
          '<p style="font-size:11px;color:#888;margin:0 0 18px;">📱 ' + telRaw + '</p>' +
          '<p style="font-size:10px;color:#555;margin:0 0 16px;line-height:1.6;background:#e3f2fd;padding:10px;border-radius:6px;text-align:left;border-left:3px solid #1976d2;">' +
          '<strong>Mensaje:</strong><br><br>' +
          '"Hola ' + nombre + ' 👋<br><br>' +
          '📅 <strong>Recordatorio de sesión agendada</strong><br><br>' +
          'Tienes una sesión programada con nosotros.<br><br>' +
          '✅ Por favor, confirma tu asistencia."</p>' +
          '<a href="https://wa.me/' + tel + '?text=' + msg + '" target="_blank" style="display:block;padding:11px;background:#25d366;color:#fff;border-radius:6px;text-decoration:none;font-weight:700;font-size:13px;margin-bottom:8px;">💬 Enviar recordatorio WhatsApp</a>' +
          '<p style="font-size:9px;color:#aaa;">Cierra esta ventana después de enviar.</p>' +
          '</div>'
        ).setWidth(320).setHeight(360).setTitle('Recordatorio sesión — ' + nombre);
        SpreadsheetApp.getUi().showModelessDialog(html, 'Recordatorio WhatsApp');

      // ── OPCIÓN 3: Marcar como completado ───────────────────────────────
      } else if (valN.includes('complet')) {
        if (estadoIdx >= 0) {
          sheet.getRange(fila, estadoIdx + 1).setValue('Completó formulario');
          sheet.getRange(fila, estadoIdx + 1).setBackground('#d9ead3').setFontColor('#274e13');
        }
        SpreadsheetApp.getUi().alert('✅ ' + nombre + ' marcada como completado.\n\nCuando sus datos lleguen desde Kobo, aparecerán automáticamente en la hoja Maestro.');
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
    SpreadsheetApp.getActiveSpreadsheet().toast('Actualizando Dashboard...', '📊', 3);
    actualizarDashboard(ss);
    ss.setActiveSheet(ss.getSheetByName('Dashboard'));
    SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard actualizado ✅', '📊', 4);
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error al actualizar Dashboard: ' + e); }
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

/** Muestra diálogo para configurar la sync automática */
function configurarSyncMatutina() {
  const props       = PropertiesService.getScriptProperties();
  const horaActual  = props.getProperty(PROP_SYNC_HORA) || '7';
  const freqActual  = props.getProperty('sync_frecuencia') || 'diaria';
  const ultimaRun   = props.getProperty('sync_last_run');
  const ultimaTxt   = ultimaRun
    ? Utilities.formatDate(new Date(ultimaRun), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    : 'Nunca';
  const triggers    = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === TRIGGER_SYNC_FN);
  const activo      = triggers.length > 0;

  const frecDescMap = {
    'cada1h':   '⚡ Cada 1 hora',
    'cada2h':   '🔄 Cada 2 horas',
    'cada6h':   '🕕 Cada 6 horas',
    'cada12h':  '🌗 Cada 12 horas',
    'diaria':   '📅 Una vez al día'
  };
  const frecDesc = frecDescMap[freqActual] || '—';

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head><meta charset="UTF-8">
    <style>
      *{box-sizing:border-box;margin:0;padding:0}
      body{font-family:'Segoe UI',Arial,sans-serif;font-size:13px;background:#f5f7ff;color:#333}
      .hdr{background:linear-gradient(135deg,#1a237e,#283593);color:#fff;padding:16px 20px}
      .hdr h2{font-size:15px;font-weight:700}
      .hdr p{font-size:10px;opacity:.8;margin-top:3px}
      .body{padding:18px}
      .card{background:#fff;border-radius:8px;padding:12px 14px;margin-bottom:14px;border:1px solid #e8eaf6}
      .status{border-left:4px solid ${activo?'#43a047':'#fb8c00'};background:${activo?'#e8f5e9':'#fff8e1'};border-radius:0 6px 6px 0}
      .status .big{font-size:13px;font-weight:700;color:${activo?'#2e7d32':'#e65100'};}
      .status .sub{font-size:10px;color:#777;margin-top:3px}
      .sec-title{font-size:9px;font-weight:700;color:#9e9e9e;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
      .freq-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:4px}
      .fq{background:#fff;border:2px solid #e0e0e0;border-radius:7px;padding:10px 8px;cursor:pointer;text-align:center}
      .fq:hover{border-color:#1a237e;background:#f0f2ff}
      .fq.sel{border-color:#1a237e;background:#e8eaf6}
      .fq strong{display:block;font-size:13px;color:#1a237e;margin-bottom:2px}
      .fq span{font-size:9px;color:#777}
      .hora-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:5px}
      .hq{background:#fff;border:2px solid #e0e0e0;border-radius:6px;padding:7px 4px;cursor:pointer;text-align:center;font-size:11px;font-weight:700;color:#555}
      .hq:hover,.hq.sel{border-color:#1a237e;background:#e8eaf6;color:#1a237e}
      .steps{font-size:10px;line-height:1.7;color:#555;background:#e8eaf6;padding:10px 12px;border-radius:6px;margin-bottom:14px}
      .btn{width:100%;padding:11px;background:#1a237e;color:#fff;border:none;border-radius:7px;font-size:13px;font-weight:700;cursor:pointer;margin-top:6px}
      .btn:hover{background:#283593}
      .btn.now{background:#1e88e5;margin-top:8px}
      .btn.danger{background:#e53935;margin-top:6px}
      #horaRow{display:none;margin-top:10px}
    </style></head><body>
    <div class="hdr">
      <h2>⏰ Auto-Sync de Datos</h2>
      <p>Importa datos nuevos automáticamente. Solo trae registros que aún no están.</p>
    </div>
    <div class="body">
      <div class="card status">
        <div class="big">${activo ? '✅ ACTIVO — '+frecDesc : '🔴 INACTIVO'}</div>
        <div class="sub">Última sync: <strong>${ultimaTxt}</strong>${activo && freqActual==='diaria' ? ' · Hora programada: '+horaActual+':00' : ''}</div>
      </div>

      <div class="steps">
        Cada vez que se ejecuta:<br>
        1️⃣ Descarga solo los <strong>registros nuevos</strong> de Kobo<br>
        2️⃣ Importa nuevos de DP_Empleabilidad → Derivados<br>
        3️⃣ Actualiza PBI y Dashboard
      </div>

      <div class="sec-title">Frecuencia de revisión</div>
      <div class="freq-grid">
        <div class="fq ${freqActual==='cada1h'?'sel':''}" onclick="selFreq(this,'cada1h')"><strong>Cada hora</strong><span>Más frecuente</span></div>
        <div class="fq ${freqActual==='cada2h'?'sel':''}" onclick="selFreq(this,'cada2h')"><strong>Cada 2h</strong><span>Balanceado</span></div>
        <div class="fq ${freqActual==='cada6h'?'sel':''}" onclick="selFreq(this,'cada6h')"><strong>Cada 6h</strong><span>✅ Recomen.</span></div>
        <div class="fq ${freqActual==='diaria'?'sel':''}" onclick="selFreq(this,'diaria')"><strong>1× al día</strong><span>Ahorra cuota</span></div>
      </div>

      <div id="horaRow">
        <div class="sec-title" style="margin-top:10px">Hora de ejecución (solo para opción diaria)</div>
        <div class="hora-grid">
          <div class="hq ${horaActual==='6'?'sel':''}" onclick="selHora(this,'6')">6:00</div>
          <div class="hq ${horaActual==='7'?'sel':''}" onclick="selHora(this,'7')">7:00 ✓</div>
          <div class="hq ${horaActual==='8'?'sel':''}" onclick="selHora(this,'8')">8:00</div>
          <div class="hq ${horaActual==='12'?'sel':''}" onclick="selHora(this,'12')">12:00</div>
        </div>
      </div>

      <button class="btn" onclick="activar()">✅ Activar Auto-Sync</button>
      <button class="btn now" onclick="ahora()">🔄 Sincronizar Ahora (manual)</button>
      ${activo ? '<button class="btn danger" onclick="desactivar()">🔴 Desactivar</button>' : ''}
    </div>
    <script>
      var freq='${freqActual}', hora='${horaActual}';
      function selFreq(el,f){
        document.querySelectorAll('.fq').forEach(x=>x.classList.remove('sel')); el.classList.add('sel'); freq=f;
        document.getElementById('horaRow').style.display=f==='diaria'?'block':'none';
      }
      function selHora(el,h){
        document.querySelectorAll('.hq').forEach(x=>x.classList.remove('sel')); el.classList.add('sel'); hora=h;
      }
      function activar(){
        document.querySelector('.btn').disabled=true; document.querySelector('.btn').textContent='Activando…';
        google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).activarAutoSync(freq,hora);
      }
      function ahora(){
        document.querySelectorAll('.btn').forEach(b=>b.disabled=true);
        document.querySelector('.now').textContent='Sincronizando…';
        google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).sincronizarTodo();
      }
      function desactivar(){
        if(confirm('¿Desactivar la sync automática?'))
          google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).desactivarSyncDiaria();
      }
      // Mostrar selector de hora si ya es diaria
      if(freq==='diaria') document.getElementById('horaRow').style.display='block';
    </script></body></html>
  `).setWidth(420).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, '⏰ Auto-Sync de Datos');
}

function activarAutoSync(frecuencia, hora) {
  try {
    // Eliminar triggers anteriores
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === TRIGGER_SYNC_FN)
      .forEach(t => ScriptApp.deleteTrigger(t));

    const props = PropertiesService.getScriptProperties();
    props.setProperty('sync_frecuencia', frecuencia);
    props.setProperty(PROP_SYNC_HORA, String(hora));

    const h = parseInt(hora) || 7;
    const builder = ScriptApp.newTrigger(TRIGGER_SYNC_FN).timeBased();

    if (frecuencia === 'cada1h') {
      builder.everyHours(1).create();
    } else if (frecuencia === 'cada2h') {
      builder.everyHours(2).create();
    } else if (frecuencia === 'cada6h') {
      builder.everyHours(6).create();
    } else {
      builder.everyDays(1).atHour(h).create();
    }

    const descMap = {cada1h:'cada hora',cada2h:'cada 2 horas',cada6h:'cada 6 horas',diaria:'diariamente a las '+h+':00'};
    return '✅ Auto-Sync activada — se ejecutará ' + (descMap[frecuencia]||'diariamente') + '.\n\nSolo importa registros nuevos que aún no estén en la hoja.';
  } catch(e) {
    logError('activarAutoSync', e);
    return '❌ Error: ' + e.message;
  }
}

function activarSyncDiaria(hora) {
  return activarAutoSync('diaria', hora);
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
      const dd = deriv.getRange(2, 1, deriv.getLastRow()-1, 13).getValues().filter(r => r[0]);
      const urgentes    = dd.filter(r => r[9] === 'URGENTE' || String(r[3]).includes('URGENTE')).length;
      const pendientes  = dd.filter(r => r[12] === 'Pendiente').length;
      const completadas = dd.filter(r => r[12] === 'Completado').length;
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

    '<div style="padding:0 16px;border-bottom:1px solid #e8eaf6;display:flex;gap:0;flex-shrink:0;" id="tabBtns">' +
      '<button onclick="tabPerfil()" id="btn0" style="flex:1;padding:9px 4px;background:transparent;border:none;border-bottom:2px solid #1a237e;color:#1a237e;font-size:10px;font-weight:700;cursor:pointer;text-transform:uppercase;letter-spacing:.3px;">Perfil</button>' +
      '<button onclick="tabDimensiones()" id="btn1" style="flex:1;padding:9px 4px;background:transparent;border:none;border-bottom:2px solid transparent;color:#9e9e9e;font-size:10px;font-weight:700;cursor:pointer;text-transform:uppercase;letter-spacing:.3px;">Dimensiones</button>' +
      '<button onclick="tabHistorial()" id="btn2" style="flex:1;padding:9px 4px;background:transparent;border:none;border-bottom:2px solid transparent;color:#9e9e9e;font-size:10px;font-weight:700;cursor:pointer;text-transform:uppercase;letter-spacing:.3px;">Historial</button>' +
    '</div>' +

    '<div id="tabPerfil" style="padding:12px 16px;overflow-y:auto;max-height:300px;">' +
      '<div style="font-size:9px;font-weight:700;color:#9e9e9e;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px;">📋 Datos Personales</div>' +
      fila('DPI', p.dpi) +
      fila('Edad', p.edad ? p.edad+' años' : '') +
      fila('Género', p.genero) +
      fila('Teléfono', p.telefono) +
      fila('Zona', p.zona) +

      '<div style="font-size:9px;font-weight:700;color:#9e9e9e;text-transform:uppercase;letter-spacing:.5px;margin:10px 0 6px;">📚 Perfil Socioeconómico</div>' +
      fila('Educación', p.educacion) +
      fila('Situación Laboral', p.laboral) +
      fila('Fortalezas', p.fortalezas) +
      fila('Objetivo', p.objetivo) +
    '</div>' +

    '<div id="tabDimensiones" style="display:none;padding:14px 16px;overflow-y:auto;max-height:300px;">' +
      '<div style="position:relative;width:100%;height:240px;">' +
        '<canvas id="radarChart" style="position:absolute;top:0;left:0;width:100%!important;height:240px!important;"></canvas>' +
      '</div>' +
      '<div id="dimList" style="margin-top:10px;display:grid;grid-template-columns:1fr 1fr;gap:4px;"></div>' +
    '</div>' +

    '<div id="tabHistorial" style="display:none;padding:12px 16px;overflow-y:auto;max-height:300px;">' +
      '<div id="historialCont" style="text-align:center;padding:24px;color:#9e9e9e;font-size:11px;">⏳ Cargando historial…</div>' +
    '</div>' +

    '<div style="padding:10px 16px;border-top:1px solid #e8eaf6;background:#f9f9f9;display:flex;gap:6px;flex-shrink:0;">' +
      (waUrl ? '<a href="'+waUrl+'" target="_blank" style="flex:1;padding:8px;background:#25d366;color:#fff;border-radius:6px;cursor:pointer;font-weight:700;font-size:10px;text-align:center;text-decoration:none;display:block;line-height:1.3;">💬<br>WhatsApp</a>' : '') +
      (p.docUrl ? '<a href="'+p.docUrl+'" target="_blank" style="flex:1;padding:8px;background:#e8eaf6;color:#1a237e;border-radius:6px;cursor:pointer;font-weight:700;font-size:10px;text-align:center;text-decoration:none;display:block;line-height:1.3;">📄<br>Expediente</a>' : '') +
      '<button onclick="cerrarPerfil()" style="flex:1;padding:8px;background:#f5f5f5;color:#555;border:none;border-radius:6px;cursor:pointer;font-weight:600;font-size:10px;line-height:1.3;">✕<br>Cerrar</button>' +
    '</div>';

  contenido.innerHTML = html;

  function setTab(activeIdx) {
    ['tabPerfil','tabDimensiones','tabHistorial'].forEach(function(id,i){
      document.getElementById(id).style.display = i === activeIdx ? 'block' : 'none';
    });
    ['btn0','btn1','btn2'].forEach(function(id,i){
      var btn = document.getElementById(id);
      btn.style.borderBottomColor = i === activeIdx ? '#1a237e' : 'transparent';
      btn.style.color = i === activeIdx ? '#1a237e' : '#9e9e9e';
    });
  }

  window.tabPerfil = function(){ setTab(0); };
  window.tabDimensiones = function(){ setTab(1); setTimeout(function(){ dibujarRadar(p); }, 100); };
  window.tabHistorial = function(){
    setTab(2);
    var cont = document.getElementById('historialCont');
    if (cont.dataset.loaded) return;
    cont.dataset.loaded = '1';
    google.script.run
      .withSuccessHandler(function(h){ renderHistorial(h, cont); })
      .withFailureHandler(function(){ cont.innerHTML = '<p style="color:#e53935;font-size:11px;text-align:center;">Error al cargar historial</p>'; })
      .obtenerHistorialParticipante(p.id);
  };

  window.renderHistorial = function(h, cont) {
    if (!h || (h.sesiones.length === 0 && h.derivaciones.length === 0)) {
      cont.innerHTML = '<div style="text-align:center;padding:24px;color:#bdbdbd;font-size:11px;">Sin registros aún</div>';
      return;
    }
    function item(fecha, icono, titulo, sub, color) {
      return '<div style="display:flex;gap:10px;margin-bottom:12px;"><div style="width:28px;height:28px;border-radius:50%;background:'+color+';color:#fff;display:flex;align-items:center;justify-content:center;font-size:13px;flex-shrink:0;">'+icono+'</div><div style="flex:1;border-bottom:1px solid #f5f5f5;padding-bottom:10px;"><div style="font-size:9px;color:#9e9e9e;margin-bottom:2px;">'+fecha+'</div><div style="font-size:11px;font-weight:600;color:#212121;">'+titulo+'</div>'+(sub?'<div style="font-size:10px;color:#757575;margin-top:2px;">'+sub+'</div>':'')+'</div></div>';
    }
    var html = '';
    if (h.sesiones.length > 0) {
      html += '<div style="font-size:9px;font-weight:700;color:#9e9e9e;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;">📅 Sesiones ('+h.sesiones.length+')</div>';
      h.sesiones.forEach(function(s){ html += item(s.fecha,'📅',s.tipo,s.notas,'#1a237e'); });
    }
    if (h.derivaciones.length > 0) {
      html += '<div style="font-size:9px;font-weight:700;color:#9e9e9e;text-transform:uppercase;letter-spacing:.5px;margin:'+(h.sesiones.length?'14px':'0')+'  0 8px;">🔗 Derivaciones ('+h.derivaciones.length+')</div>';
      var dColors = {URGENTE:'#c62828',ALTA:'#e65100',MEDIA:'#f57f17',SUGERIDA:'#1565c0'};
      h.derivaciones.forEach(function(d){ html += item(d.fecha,'🔗',d.destino,d.motivo+(d.estado?' · '+d.estado:''),dColors[d.prioridad]||'#546e7a'); });
    }
    cont.innerHTML = html;
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

    // COL_DER (1-based): ID:1, FECHA:2, NOMBRE:3, DPI:4, EDAD:5, GENERO:6, TELEFONO:7,
    //                    EDUCACION:8, FORMACION:9, COHORTE:10, NOTA:11, ESTADO:12, ORIGEN:13, FECHA_IMP:14
    return hDeriv.getRange(2, 1, hDeriv.getLastRow()-1, 14).getValues()
      .filter(r => r[COL_DER.NOMBRE-1]) // tiene nombre
      .map(r => {
        const tel = String(r[COL_DER.TELEFONO-1]||'').replace(/\D/g,'');
        const telN = tel.length === 8 ? '502'+tel : tel;
        const enM = idsEnMaestro.has(String(r[COL_DER.ID-1]).trim()) || (tel && telEnMaestro.has(tel)) || (telN && telEnMaestro.has(telN));
        return {
          nombre:      String(r[COL_DER.NOMBRE    -1]||''),
          telefono:    String(r[COL_DER.TELEFONO  -1]||''),
          genero:      String(r[COL_DER.GENERO    -1]||''),
          educacion:   String(r[COL_DER.EDUCACION -1]||''),
          formacion:   String(r[COL_DER.FORMACION -1]||''),
          estadoDeriv: String(r[COL_DER.ESTADO    -1]||'Pendiente'),
          enMaestro:   enM
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
        <div class="dmi" onclick="run('abrirFormCierre')"><span class="dmi-ico">🏁</span><div class="dmi-t"><span class="dmi-l">Paso a Paso de Cierre</span><span class="dmi-s">Registrar cierre del participante</span></div></div>
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
  // Deshabilitar botón inmediatamente para evitar doble envío
  var btn=document.querySelector('.ses-ok');
  if(btn){btn.disabled=true;btn.textContent='⏳ Guardando…';}
  var ov=document.getElementById('sesOv');
  google.script.run
    .withSuccessHandler(function(res){
      if(ov)ov.remove();
      if(res&&res.duplicado){alert('⚠️ Esta sesión ya estaba registrada (duplicado ignorado).');}
    })
    .withFailureHandler(function(e){
      if(btn){btn.disabled=false;btn.textContent='✅ Guardar Sesión';}
      alert('❌ Error: '+e.message);
    })
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

    // Verificar duplicado: misma ID + misma fecha + mismo tipo registrado en los últimos 60 segundos
    if (hSes.getLastRow() > 1) {
      const ahora    = Date.now();
      const fechaNorm = String(fecha).trim();
      const recientes = hSes.getRange(2, 1, hSes.getLastRow()-1, 7).getValues()
        .filter(r => {
          const mismoId   = String(r[1]).trim() === String(pid).trim();
          const mismoTipo = String(r[3]).trim() === String(tipo).trim();
          const mismaFecha = r[0] instanceof Date
            ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd') === fechaNorm
            : String(r[0]).slice(0,10) === fechaNorm;
          const reciente  = r[6] instanceof Date && (ahora - r[6].getTime()) < 60000;
          return mismoId && mismoTipo && mismaFecha && reciente;
        });
      if (recientes.length > 0) return { duplicado: true };
    }

    const orientador = Session.getEffectiveUser().getEmail();
    hSes.appendRow([new Date(fecha), pid, pnom, tipo, notas||'', orientador, new Date()]);
    CacheService.getScriptCache().remove('p_maestro');
    return { duplicado: false };
  } catch(e) {
    logError('guardarSesionInterna', e);
    throw e;
  }
}

// Retorna sesiones + derivaciones de un participante para el Historial del perfil
function obtenerHistorialParticipante(id) {
  const ss = SpreadsheetApp.getActive();
  const resultado = { sesiones: [], derivaciones: [] };
  try {
    const hSes = ss.getSheetByName('Sesiones');
    if (hSes && hSes.getLastRow() > 1) {
      resultado.sesiones = hSes.getRange(2, 1, hSes.getLastRow()-1, 7).getValues()
        .filter(r => String(r[1]).trim() === String(id).trim())
        .map(r => ({
          fecha:  r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '—',
          tipo:   String(r[3] || ''),
          notas:  String(r[4] || ''),
          user:   String(r[5] || '')
        }))
        .sort(function(a,b){ return b.fecha.localeCompare(a.fecha); });
    }
  } catch(e) {}
  try {
    const hDer = ss.getSheetByName('Derivaciones');
    if (hDer && hDer.getLastRow() > 1) {
      resultado.derivaciones = hDer.getRange(2, 1, hDer.getLastRow()-1, 13).getValues()
        .filter(r => String(r[1]).trim() === String(id).trim())
        .map(r => ({
          fecha:    r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '—',
          tipo:     String(r[3] || ''),
          destino:  String(r[4] || ''),
          motivo:   String(r[5] || ''),
          prioridad:String(r[9]  || ''),
          estado:   String(r[12] || '')
        }))
        .sort(function(a,b){ return b.fecha.localeCompare(a.fecha); });
    }
  } catch(e) {}
  return resultado;
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
      '<div class="info"><strong>Participante:</strong> '+escaparHtml(nom)+'<br><strong>Creamos ID:</strong> '+escaparHtml(id)+'</div>' +
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
// PASO A PASO DE CIERRE
// ============================================================================

function abrirFormCierre() {
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
    const fila  = rango.getRow();
    const id    = datos[C.ID-1]     || '';
    const nom   = datos[C.NOMBRE-1] || '';
    const est   = datos[C.ESTADO-1] || '';
    if (!id) { SpreadsheetApp.getUi().alert('⚠️ Participante sin ID.'); return; }

    const estadosFinales = ['Cierre', 'Completado'];
    if (estadosFinales.includes(est)) {
      SpreadsheetApp.getUi().alert('ℹ️ Este participante ya tiene cierre registrado.\nEstado actual: ' + est);
      return;
    }

    const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const html = HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
      '*{box-sizing:border-box;margin:0;padding:0}' +
      'body{font-family:Arial,sans-serif;background:#f5f7ff;font-size:12px;padding:0}' +
      '.header{background:linear-gradient(135deg,#1a237e,#283593);color:#fff;padding:16px;text-align:center}' +
      '.header h3{font-size:15px;margin-bottom:4px}' +
      '.header span{font-size:11px;opacity:.85}' +
      '.steps{display:flex;justify-content:center;gap:0;padding:12px 16px 8px;background:#fff;border-bottom:1px solid #e8eaf6}' +
      '.step{display:flex;align-items:center;gap:4px;font-size:10px;color:#9e9e9e}' +
      '.step.active{color:#1a237e;font-weight:700}' +
      '.step.done{color:#2e7d32}' +
      '.step-num{width:20px;height:20px;border-radius:50%;border:2px solid currentColor;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;flex-shrink:0}' +
      '.step-sep{width:24px;height:2px;background:#e0e0e0;margin:0 4px}' +
      '.step.done .step-sep,.step.active .step-sep{background:#1a237e}' +
      '.panel{display:none;padding:16px;animation:fd .25s}' +
      '.panel.on{display:block}' +
      '@keyframes fd{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}' +
      '.info-box{background:#e8eaf6;border-left:4px solid #1a237e;padding:10px 12px;border-radius:4px;margin-bottom:14px}' +
      '.info-box strong{color:#1a237e;font-size:12px}' +
      '.info-box p{color:#444;margin-top:4px;font-size:11px}' +
      '.tipo-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:14px}' +
      '.tipo-card{border:2px solid #e0e0e0;border-radius:8px;padding:12px 8px;text-align:center;cursor:pointer;transition:.2s;background:#fff}' +
      '.tipo-card:hover{border-color:#3949ab;background:#e8eaf6}' +
      '.tipo-card.sel{border-color:#1a237e;background:#e8eaf6}' +
      '.tipo-card .ico{font-size:22px;margin-bottom:4px}' +
      '.tipo-card .lbl{font-size:10px;font-weight:700;color:#333;line-height:1.3}' +
      'label{display:block;font-weight:700;margin:10px 0 4px;color:#333;font-size:11px}' +
      'input,select,textarea{width:100%;padding:8px;border:1px solid #c5cae9;border-radius:4px;font-size:12px;background:#fff}' +
      'textarea{height:70px;resize:vertical}' +
      '.btn-sig{background:#1a237e;color:#fff;padding:10px;border:none;border-radius:5px;cursor:pointer;width:100%;font-weight:700;font-size:12px;margin-top:10px}' +
      '.btn-sig:hover{background:#283593}' +
      '.btn-back{background:#e0e0e0;color:#333;padding:8px;border:none;border-radius:5px;cursor:pointer;width:100%;font-size:11px;margin-top:6px}' +
      '.btn-back:hover{background:#bdbdbd}' +
      '.estado-chip{display:inline-block;padding:2px 10px;border-radius:12px;font-size:10px;font-weight:700;background:#fff3e0;color:#e65100;margin-left:6px}' +
      '</style></head><body>' +

      '<div class="header">' +
      '<h3>🏁 Paso a Paso de Cierre</h3>' +
      '<span>' + escaparHtml(nom) + ' · ' + escaparHtml(id) + '</span>' +
      '</div>' +

      '<div class="steps">' +
      '<div class="step active" id="s1"><div class="step-num">1</div><span>Tipo</span></div>' +
      '<div class="step-sep"></div>' +
      '<div class="step" id="s2"><div class="step-num">2</div><span>Detalles</span></div>' +
      '<div class="step-sep"></div>' +
      '<div class="step" id="s3"><div class="step-num">3</div><span>Confirmar</span></div>' +
      '</div>' +

      /* ── PASO 1: tipo de cierre ── */
      '<div class="panel on" id="p1">' +
      '<div class="info-box"><strong>Participante</strong>' +
      '<p>' + escaparHtml(nom) + ' — Estado actual: <strong>' + escaparHtml(est || 'Sin estado') + '</strong></p></div>' +
      '<label>¿Cómo finaliza el proceso?</label>' +
      '<div class="tipo-grid">' +
      '<div class="tipo-card" id="tc-completado" onclick="selTipo(\'✅ Completado\')"><div class="ico">✅</div><div class="lbl">Completado<br>(Graduado)</div></div>' +
      '<div class="tipo-card" id="tc-no" onclick="selTipo(\'❌ No completado\')"><div class="ico">❌</div><div class="lbl">No completado</div></div>' +
      '<div class="tipo-card" id="tc-abandono" onclick="selTipo(\'🚪 Abandonó\')"><div class="ico">🚪</div><div class="lbl">Abandonó</div></div>' +
      '<div class="tipo-card" id="tc-derivado" onclick="selTipo(\'➡️ Derivado a cierre\')"><div class="ico">➡️</div><div class="lbl">Derivado a cierre</div></div>' +
      '</div>' +
      '<button class="btn-sig" onclick="irPaso2()">Siguiente →</button>' +
      '</div>' +

      /* ── PASO 2: fecha y resultado ── */
      '<div class="panel" id="p2">' +
      '<div class="info-box"><strong>Tipo de cierre seleccionado:</strong><span class="estado-chip" id="tipo-sel-lbl">—</span></div>' +
      '<label>Fecha de cierre *</label>' +
      '<input type="date" id="fecha" value="' + hoy + '">' +
      '<label>Resultado / Observaciones *</label>' +
      '<textarea id="resultado" placeholder="Describe el resultado del proceso, logros alcanzados, situación final del participante…"></textarea>' +
      '<button class="btn-sig" onclick="irPaso3()">Siguiente →</button>' +
      '<button class="btn-back" onclick="irPaso(1)">← Atrás</button>' +
      '</div>' +

      /* ── PASO 3: confirmación ── */
      '<div class="panel" id="p3">' +
      '<div class="info-box"><strong>Resumen del cierre</strong><p id="resumen-txt" style="margin-top:6px;line-height:1.6"></p></div>' +
      '<button class="btn-sig" onclick="confirmarCierre()" id="btn-confirmar">🏁 Registrar Cierre</button>' +
      '<button class="btn-back" onclick="irPaso(2)">← Atrás</button>' +
      '</div>' +

      '<script>' +
      'var tipoCierre="";' +
      'var ID_MAP={"✅ Completado":"tc-completado","❌ No completado":"tc-no","🚪 Abandonó":"tc-abandono","➡️ Derivado a cierre":"tc-derivado"};' +
      'function selTipo(t){' +
      '  tipoCierre=t;' +
      '  Object.keys(ID_MAP).forEach(function(k){document.getElementById(ID_MAP[k]).classList.remove("sel");});' +
      '  document.getElementById(ID_MAP[t]).classList.add("sel");' +
      '}' +
      'function irPaso(n){' +
      '  [1,2,3].forEach(function(i){' +
      '    document.getElementById("p"+i).classList.remove("on");' +
      '    var s=document.getElementById("s"+i);' +
      '    s.classList.remove("active","done");' +
      '    if(i<n)s.classList.add("done");' +
      '    else if(i===n)s.classList.add("active");' +
      '  });' +
      '  document.getElementById("p"+n).classList.add("on");' +
      '}' +
      'function irPaso2(){' +
      '  if(!tipoCierre){alert("Selecciona un tipo de cierre.");return;}' +
      '  document.getElementById("tipo-sel-lbl").textContent=tipoCierre;' +
      '  irPaso(2);' +
      '}' +
      'function irPaso3(){' +
      '  var fecha=document.getElementById("fecha").value;' +
      '  var res=document.getElementById("resultado").value.trim();' +
      '  if(!fecha){alert("Ingresa la fecha de cierre.");return;}' +
      '  if(!res){alert("Ingresa el resultado/observaciones.");return;}' +
      '  var fechaFmt=fecha.split("-").reverse().join("/");' +
      '  document.getElementById("resumen-txt").innerHTML=' +
      '    "<b>Participante:</b> ' + escaparHtml(nom) + '<br>"' +
      '   +"<b>ID:</b> ' + escaparHtml(id) + '<br>"' +
      '   +"<b>Tipo:</b> "+tipoCierre+"<br>"' +
      '   +"<b>Fecha:</b> "+fechaFmt+"<br>"' +
      '   +"<b>Resultado:</b> "+res;' +
      '  irPaso(3);' +
      '}' +
      'function confirmarCierre(){' +
      '  var fecha=document.getElementById("fecha").value;' +
      '  var res=document.getElementById("resultado").value.trim();' +
      '  var btn=document.getElementById("btn-confirmar");' +
      '  btn.disabled=true;btn.textContent="⏳ Guardando…";' +
      '  google.script.run' +
      '    .withSuccessHandler(function(){alert("✅ Cierre registrado correctamente.");google.script.host.close();})' +
      '    .withFailureHandler(function(e){alert("❌ Error: "+e.message);btn.disabled=false;btn.textContent="🏁 Registrar Cierre";})' +
      '    .guardarCierre(' + fila + ',"' + escaparHtml(id) + '","' + escaparHtml(nom) + '",tipoCierre,fecha,res);' +
      '}' +
      '</script></body></html>'
    ).setTitle('🏁 Cierre — ' + nom).setWidth(360);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e); }
}

function guardarCierre(fila, id, nombre, tipoCierre, fechaCierreStr, resultado) {
  try {
    const ss   = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) throw new Error('Hoja Maestro no encontrada.');
    const C = CONFIG.COL;

    const nuevoEstado = tipoCierre.includes('Completado') ? 'Completado' : 'Cierre';
    const fechaCierre = fechaCierreStr ? new Date(fechaCierreStr) : new Date();

    hoja.getRange(fila, C.ESTADO).setValue(nuevoEstado);
    hoja.getRange(fila, C.TIPO_CIERRE).setValue(tipoCierre);
    hoja.getRange(fila, C.FECHA_CIERRE).setValue(fechaCierre);
    hoja.getRange(fila, C.RESULTADO_CIERRE).setValue(resultado);

    // Colorear la fila de cierre en verde oscuro
    hoja.getRange(fila, 1, 1, 30).setBackground(nuevoEstado === 'Completado' ? '#d9ead3' : '#f3f3f3');

    // Registrar en hoja Cierres
    let hCierres = ss.getSheetByName('Cierres');
    if (!hCierres) {
      hCierres = ss.insertSheet('Cierres');
      hCierres.appendRow(['Fecha_Cierre','ID_Creamos','Nombre','Tipo_Cierre','Resultado','Encargado','Estado_Final','Timestamp']);
      hCierres.getRange(1,1,1,8).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#fff');
      hCierres.setFrozenRows(1);
    }
    const encargado = Session.getEffectiveUser().getEmail();
    hCierres.appendRow([fechaCierre, id, nombre, tipoCierre, resultado, encargado, nuevoEstado, new Date()]);

    SpreadsheetApp.flush();

    // Evento de calendario post-cierre si aplica
    try { crearEventoCalendario(nombre, nuevoEstado); } catch(e2) {}
  } catch(e) {
    logError('guardarCierre', e);
    throw e;
  }
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
      ? sheet.getRange(2, 13, sheet.getLastRow()-1, 1).getValues().filter(r => r[0]==='Pendiente').length
      : 0;
    if (pend > 0) SpreadsheetApp.getUi().alert('📋 Derivaciones\n\n⚠️ Tienes ' + pend + ' derivación(es) PENDIENTE(S).\n\nCambia el Estado a "Completado" cuando las gestiones.');
  } catch(e) { SpreadsheetApp.getUi().alert('❌ Error: ' + e); }
}

function guardarDerivacion(id, nombre, tipo, destino, motivo, notas) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Derivaciones');
    if (!sheet) return;
    // 13 cols: Fecha|ID|Nombre|Tipo|Destino|Motivo|Responsable|FechaSeg|Notas|Prioridad|SesionRel|FechaCierre|Estado
    sheet.appendRow([new Date(), id, nombre, tipo||'Manual', destino, motivo,
                     Session.getEffectiveUser().getEmail(), '', notas||'', '', '', '', 'Pendiente']);
    const n     = sheet.getLastRow();
    const color = tipo && tipo.includes('URGENTE') ? '#ffcdd2' : tipo && tipo.includes('ALERTA') ? '#fff9c4' : '#e8f5e9';
    sheet.getRange(n, 1, 1, 13).setBackground(color);
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

    const add = (prioNivel, tipo, dest, motivo) => {
      // 13 cols: Fecha|ID|Nombre|Tipo|Destino|Motivo|Responsable|FechaSeg|Notas|Prioridad|SesionRel|FechaCierre|Estado
      deriv.appendRow([hoy, id, nombre, tipo, dest, motivo, '', '', '', prioNivel, '', '', 'Pendiente']);
      const c = prioNivel === 'URGENTE' ? '#ffcdd2' : prioNivel === 'ALTA' ? '#ffe0b2' : '#fff9c4';
      deriv.getRange(deriv.getLastRow(), 1, 1, 13).setBackground(c);
    };

    if (perfil==='Perfil D' || prioridad==='CRÍTICO') {
      add('URGENTE','Apoyo Emocional','Creamos Voces','Barreras críticas — Perfil ' + perfil);
      if (dim5 <= 4) add('URGENTE','Servicios Profesionales','Servicios Profesionales','Barreras estructurales: ' + dim5 + '/10');
      enviarAlertaCritica(id, nombre, perfil, dim5);
    }
    if (perfil === 'Perfil B') add('MEDIA','Orientación Vocacional','Mentoría Vocacional','Requiere orientación vocacional — 3-6 sesiones');
    if (perfil === 'Perfil C') {
      if (dim1 <= 4) add('MEDIA','Educación','Educación de Adultos','Capital educativo bajo: ' + dim1 + '/10');
      if (dim3 <= 4) add('MEDIA','Digital','Alfabetización Digital','Habilidades digitales bajas: ' + dim3 + '/10');
    }
    if (dim5 <= 3 && perfil !== 'Perfil D') add('ALTA','Apoyo Emocional','Creamos Voces','Barreras muy bajas: ' + dim5 + '/10');
    if (dim6 <= 3) add('SUGERIDA','Red de Apoyo','Grupos Comunitarios','Red de apoyo débil: ' + dim6 + '/10');
    if (perfil === 'Perfil A') add('SUGERIDA','Empleabilidad','Intermediación Laboral','Lista para empleo — ' + Number(fila[M.PUNTAJE-1]) + '/60 pts');
  } catch(e) {}
}

function enviarAlertaCritica(id, nombre, perfil, dim5) {
  try {
    MailApp.sendEmail(
      getAdminEmail(),
      '🚨 URGENTE: Barreras críticas — ' + nombre,
      '🚨 CASO URGENTE\n\nParticipante: ' + nombre + '\nCreamos ID: ' + id +
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
    '<button id="btnG" onclick="guardar()">✅ Guardar Sesión</button>'+
    '<script>function guardar(){'+
    'var f=document.getElementById("sf").value;'+
    'var t=document.getElementById("st").value;'+
    'var n=document.getElementById("sn").value;'+
    'if(!f||!t){alert("Completa la fecha y tipo");return;}'+
    'var b=document.getElementById("btnG");b.disabled=true;b.textContent="⏳ Guardando…";'+
    'google.script.run.withSuccessHandler(function(r){if(r&&r.duplicado){alert("⚠️ Sesión ya registrada (duplicado ignorado).");}google.script.host.close();})'+
    '.withFailureHandler(function(e){b.disabled=false;b.textContent="✅ Guardar Sesión";alert("Error: "+e.message);})'+
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
  // Columnas 28-30 — Datos de Cierre
  const cierreHeaders = [['Tipo_Cierre','Fecha_Cierre','Resultado_Cierre']];
  maestro.getRange(1, 28, 1, 3).setValues(cierreHeaders)
    .setBackground('#1b5e20').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
  maestro.setColumnWidth(28, 140).setColumnWidth(29, 110).setColumnWidth(30, 220);
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
      // Headers en el mismo orden que COL_DER + columna Acción al final
      hd.appendRow(['ID','Fecha Orig.','Nombre','DPI','Edad','Género','Teléfono',
                    'Educación','Formación','Cohorte','Notas','Estado Derivado','Fuente','Fecha Importación','⚡ Acción']);
      hd.getRange(1,1,1,15).setBackground('#880e4f').setFontColor('#fff').setFontWeight('bold');
      hd.setFrozenRows(1);
      creadas.push('Derivados');
    } else {
      // Asegurar columna Acción si no existe
      if (!String(hd.getRange(1,15).getValue()).includes('Acción')) {
        hd.getRange(1,15).setValue('⚡ Acción').setBackground('#880e4f').setFontColor('#fff').setFontWeight('bold');
      }
    }
    // Validaciones usando COL_DER para que siempre apunten a la columna correcta
    hd.getRange(2, COL_DER.ESTADO, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente formulario','Formulario enviado','Completó formulario','Rechazado'], true).build());
    hd.getRange(2, 15, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Enviar formulario Kobo','Recordatorio de sesión agendada','Ya completó el formulario'], true).build());
    hd.setColumnWidth(COL_DER.NOMBRE, 180).setColumnWidth(15, 200);
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
  {
    let hder = ss.getSheetByName('Derivaciones');
    if (!hder) {
      hder = ss.insertSheet('Derivaciones');
      // 13 cols: Fecha|ID|Nombre|Tipo|Destino|Motivo|Responsable|FechaSeg|Notas|Prioridad|SesionRel|FechaCierre|Estado
      hder.appendRow(['Fecha','Creamos_ID','Nombre','Tipo','Destino','Motivo',
                      'Responsable','Fecha_Seguimiento','Notas','Prioridad','Sesion_Relacionada','Fecha_Cierre','Estado']);
      hder.getRange(1,1,1,13).setBackground('#bf360c').setFontColor('#fff').setFontWeight('bold');
      hder.setFrozenRows(1);
      hder.setColumnWidth(3,180).setColumnWidth(5,200).setColumnWidth(6,250).setColumnWidth(9,200);
      creadas.push('Derivaciones');
    } else {
      // Si el Estado estaba en col 7 (estructura antigua), migrar automáticamente
      const hdrs = hder.getRange(1,1,1,14).getValues()[0];
      if (String(hdrs[6]).includes('Estado') && hder.getLastRow() > 1) {
        const datos = hder.getRange(1, 1, hder.getLastRow(), 14).getValues();
        const nuevos = datos.map(function(r, i) {
          if (i === 0) return ['Fecha','Creamos_ID','Nombre','Tipo','Destino','Motivo','Responsable','Fecha_Seguimiento','Notas','Prioridad','Sesion_Relacionada','Fecha_Cierre','Estado'];
          return [r[0],r[1],r[2],r[3],r[4],r[5],r[7],r[8],r[9],r[10],r[11],r[12],r[6]];
        });
        hder.getRange(1, 1, nuevos.length, 13).setValues(nuevos);
        // Limpiar columna 14 (Resultado antigua) sin borrar datos reales
        if (hder.getLastRow() > 1) hder.getRange(2, 14, hder.getLastRow()-1, 1).clearContent();
        hder.getRange(1, 14).clearContent();
      }
    }
    // Validaciones: Estado en col 13, Prioridad en col 10
    hder.getRange(2, 13, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente','En gestión','Contactado','En proceso','Completado','No interesado','Cancelado'], true).build());
    hder.getRange(2, 10, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['URGENTE','ALTA','MEDIA','SUGERIDA'], true).build());
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
  hf.setColumnWidth(1,160).setColumnWidth(2,380).setColumnWidth(3,150).setColumnWidth(4,80).setColumnWidth(5,160).setColumnWidth(6,280);

  // Fila 2: DP_Empleabilidad (fuente principal)
  hf.appendRow(['DP_Empleabilidad', DP_EMPLEABILIDAD.SPREADSHEET_ID, DP_EMPLEABILIDAD.HOJA, 'SÍ', '', 'Fuente principal de derivados']);
  hf.getRange(2,1,1,6).setBackground('#e0f7fa');

  // Fila 3 y 4: plantillas para AYB y TECH — el usuario solo pone el Spreadsheet_ID
  hf.appendRow(['AYB','← PEGA AQUÍ EL ID DEL SPREADSHEET DE AYB','Paso a paso','NO','','Programa AYB — cambia NO a SÍ cuando tengas el ID']);
  hf.getRange(3,1,1,6).setBackground('#fff8e1').setFontColor('#555');

  hf.appendRow(['TECH','← PEGA AQUÍ EL ID DEL SPREADSHEET DE TECH','Paso a paso','NO','','Programa TECH — cambia NO a SÍ cuando tengas el ID']);
  hf.getRange(4,1,1,6).setBackground('#e8f5e9').setFontColor('#555');

  // Fila 5: instrucción general
  hf.appendRow(['','','','','','⬆ Para activar: reemplaza el texto de la col B con el ID real del Google Sheet y cambia Activo a SÍ']);
  hf.getRange(5,1,1,6).setFontColor('#9e9e9e').setFontStyle('italic');

  // Validación dropdown Activo
  hf.getRange(2,4,500,1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['SÍ','NO'], true).build());
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

  // 3. Todas las demás fuentes activas en Fuentes_Externas
  leerFuentesExternas().forEach(function(f) {
    if (String(f.ssId).trim() === String(DP_EMPLEABILIDAD.SPREADSHEET_ID).trim()) return;
    try {
      const res = importarFuenteGenerica(f.ssId, f.hoja, f.nombre, true);
      if (res.error) {
        log.push('⚠️ ' + f.nombre + ': ' + res.error);
        errores++;
      } else {
        log.push('✅ ' + f.nombre + ': ' + res.nuevos + ' nuevo(s), ' + res.omitidos + ' ya existían');
      }
    } catch(e) {
      log.push('⚠️ ' + f.nombre + ': ' + e.message);
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
// LIMPIAR SESIONES DUPLICADAS
// ============================================================================

/**
 * Elimina filas duplicadas de la hoja Sesiones.
 * Criterio de duplicado: mismo ID_Participante + misma Fecha + mismo Tipo_Sesion.
 * Conserva la primera ocurrencia y borra las repetidas.
 */
function limpiarSesionesDuplicadas() {
  const ui  = SpreadsheetApp.getUi();
  const ss  = SpreadsheetApp.getActive();
  const hSes = ss.getSheetByName('Sesiones');
  if (!hSes || hSes.getLastRow() < 3) {
    ui.alert('ℹ️ No hay suficientes filas en Sesiones para revisar duplicados.');
    return;
  }

  const datos = hSes.getRange(2, 1, hSes.getLastRow()-1, 7).getValues();
  const tz    = Session.getScriptTimeZone();
  const vistos = new Set();
  const filasAEliminar = []; // índices 0-based sobre el array datos

  datos.forEach(function(r, i) {
    const fechaStr = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')
      : String(r[0]).slice(0,10);
    const clave = String(r[1]).trim() + '|' + fechaStr + '|' + String(r[3]).trim();
    if (vistos.has(clave)) {
      filasAEliminar.push(i + 2); // fila real en la hoja (1-based, +1 por header)
    } else {
      vistos.add(clave);
    }
  });

  if (filasAEliminar.length === 0) {
    ui.alert('✅ No se encontraron duplicados en Sesiones.');
    return;
  }

  const conf = ui.alert(
    '⚠️ Duplicados encontrados',
    'Se encontraron ' + filasAEliminar.length + ' fila(s) duplicada(s) en Sesiones.\n\n' +
    '¿Eliminarlas? (se conserva la primera ocurrencia de cada sesión)',
    ui.ButtonSet.YES_NO
  );
  if (conf !== ui.Button.YES) return;

  // Eliminar de abajo hacia arriba para no desplazar índices
  filasAEliminar.reverse().forEach(function(n) { hSes.deleteRow(n); });

  ui.alert('✅ ' + filasAEliminar.length + ' duplicado(s) eliminado(s) de Sesiones.');
}

// ============================================================================
// MIGRACIÓN DE ESTRUCTURA — arregla datos existentes sin borrar filas
// ============================================================================

/**
 * Corre una sola vez para:
 * 1. Corregir headers de Derivados (no toca los datos, solo la fila 1)
 * 2. Mover la columna Estado de Derivaciones al final (col 13)
 *    y eliminar Resultado (col 14) — sin borrar datos de otras columnas
 */
function migrarEstructura() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const log = [];

  // ── 1. Derivados: corregir headers ─────────────────────────────────────────
  const hDer = ss.getSheetByName('Derivados');
  if (hDer) {
    const h1 = String(hDer.getRange(1,1).getValue());
    if (h1 === 'Fecha_Import' || h1 === 'ID') {
      // Reescribir la fila de headers para que coincida con COL_DER + Acción
      hDer.getRange(1, 1, 1, 14).setValues([[
        'ID','Fecha Orig.','Nombre','DPI','Edad','Género','Teléfono',
        'Educación','Formación','Cohorte','Notas','Estado Derivado','Fuente','Fecha Importación'
      ]]);
      if (!String(hDer.getRange(1,15).getValue()).includes('Acción'))
        hDer.getRange(1,15).setValue('⚡ Acción');
      hDer.getRange(1,1,1,15).setBackground('#880e4f').setFontColor('#fff').setFontWeight('bold');
      log.push('✅ Derivados: headers corregidos (ID en col A)');
    } else {
      log.push('ℹ️ Derivados: headers ya están correctos');
    }
  } else {
    log.push('⚠️ Derivados: hoja no encontrada');
  }

  // ── 2. Derivaciones: mover Estado al final, quitar Resultado ───────────────
  const hDeriv = ss.getSheetByName('Derivaciones');
  if (hDeriv && hDeriv.getLastRow() >= 1) {
    const nCols  = Math.max(hDeriv.getLastColumn(), 14);
    const allRows = hDeriv.getRange(1, 1, hDeriv.getLastRow(), nCols).getValues();
    const hdrs   = allRows[0];

    // Verificar si Estado está en col 7 (índice 6) — estructura antigua
    const estadoEnCol7 = String(hdrs[6]).toLowerCase().includes('estado');

    if (estadoEnCol7) {
      const nuevos = allRows.map(function(r, i) {
        if (i === 0) {
          return ['Fecha','Creamos_ID','Nombre','Tipo','Destino','Motivo',
                  'Responsable','Fecha_Seguimiento','Notas','Prioridad',
                  'Sesion_Relacionada','Fecha_Cierre','Estado'];
        }
        // Reordenar: mover Estado (col 7, índice 6) al final
        // Responsable era col 8 (índice 7), etc.
        return [r[0],r[1],r[2],r[3],r[4],r[5],r[7],r[8],r[9],r[10],r[11],r[12],r[6]];
      });

      // Limpiar el rango y reescribir con nueva estructura (13 cols)
      hDeriv.getRange(1, 1, hDeriv.getLastRow(), 14).clearContent();
      hDeriv.getRange(1, 1, nuevos.length, 13).setValues(nuevos);

      // Restaurar estilo del header
      hDeriv.getRange(1,1,1,13).setBackground('#bf360c').setFontColor('#fff').setFontWeight('bold');

      // Aplicar validaciones con nueva posición
      hDeriv.getRange(2, 13, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(['Pendiente','En gestión','Contactado','En proceso','Completado','No interesado','Cancelado'], true).build());
      hDeriv.getRange(2, 10, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(['URGENTE','ALTA','MEDIA','SUGERIDA'], true).build());

      log.push('✅ Derivaciones: Estado movido a col 13, Resultado eliminado (' + (nuevos.length-1) + ' filas migradas)');
    } else if (String(hdrs[12]).toLowerCase().includes('estado')) {
      log.push('ℹ️ Derivaciones: ya tiene la estructura nueva (Estado en col 13)');
    } else {
      log.push('⚠️ Derivaciones: no se reconoció la estructura actual — revisa manualmente');
    }
  } else {
    log.push('⚠️ Derivaciones: hoja no encontrada o vacía');
  }

  ui.alert('🔧 Migración completada', log.join('\n\n') +
    '\n\n✅ No se borraron datos.\nSi algo se ve raro, avísale al equipo técnico.',
    ui.ButtonSet.OK);
}

// ============================================================================
// FIN — v8.7 + Mejoras Inspección
// ============================================================================
