// ============================================================================
// SISTEMA PASO A PASO - FITO v5.0
// ============================================================================

// ============================================================================
// CONFIGURACIÓN — EDITA SOLO ESTA SECCIÓN
// ============================================================================

const CONFIG = {
  // Google
  SPREADSHEET_ID:        SpreadsheetApp.getActive().getId(),
  FOLDER_PARTICIPANTES_ID: "1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU",

  // Kobo — ya configurado
  KOBO_API_KEY:  "64cc018b88067397addd36b09288be8b6539cf39",
  KOBO_ASSET_ID: "abHRWdRnPhKwzPQBajc7RZ",
  KOBO_URL:      "https://kf.kobotoolbox.org/api/v2",

  // Notificaciones
  ADMIN_EMAIL: "adrian@creamosguatemla.org",

  // Nombre de la hoja principal
  HOJA: "Maestro",

  // Colores por perfil
  COLORES: {
    "Perfil A": { fondo: "#d9ead3", texto: "#274e13" },
    "Perfil B": { fondo: "#cfe2f3", texto: "#1c4587" },
    "Perfil C": { fondo: "#fff2cc", texto: "#7f6000" },
    "Perfil D": { fondo: "#f4cccc", texto: "#660000" }
  },

  // Columnas (posición en el sheet, empieza en 1)
  COL: {
    ID:         1,   // A
    FECHA:      2,   // B
    NOMBRE:     3,   // C
    DPI:        4,   // D
    EDAD:       5,   // E
    GENERO:     6,   // F
    TELEFONO:   7,   // G
    ZONA:       8,   // H
    EMAIL:      9,   // I
    EDUCACION:  10,  // J
    LABORAL:    11,  // K
    FORTALEZAS: 12,  // L
    OBJETIVO:   13,  // M
    PERFIL:     14,  // N
    PRIORIDAD:  15,  // O
    PUNTAJE:    16,  // P
    DIM1:       17,  // Q
    DIM2:       18,  // R
    DIM3:       19,  // S
    DIM4:       20,  // T
    DIM5:       21,  // U
    DIM6:       22,  // V
    ESTADO:     23,  // W
    CARPETA_ID: 24,  // X
    DOC_ID:     25,  // Y
    DOC_URL:    26   // Z
  }
};

// ============================================================================
// MENÚ
// ============================================================================

function onOpen() {
  cargarApiKey();
  SpreadsheetApp.getUi()
    .createMenu('📊 PASO A PASO')
    .addItem('📥 INSTALAR SISTEMA', 'instalar')
    .addSeparator()
    .addItem('🔄 Sincronizar Kobo', 'sincronizar')
    .addItem('🎨 Colorear Tabla',   'colorearTabla')
    .addSeparator()
    .addItem('⚙️ Configuración',    'abrirConfiguracion')
    .addItem('🧪 Probar Kobo',      'probarKobo')
    .addToUi();
}

// Abre el sidebar con la ficha al seleccionar una fila
function onSelectionChange(e) {
  try {
    const hoja = e.range.getSheet();
    if (hoja.getName() !== CONFIG.HOJA) return;
    const fila = e.range.getRow();
    if (fila < 2) return;
    const datos = hoja.getRange(fila, 1, 1, 26).getValues()[0];
    if (!datos[0]) return;
    abrirFichaParticipante(datos);
  } catch(_) {}
}

// ============================================================================
// INSTALAR
// ============================================================================

function instalar() {
  const ss = SpreadsheetApp.getActive();

  // Crear hoja Maestro si no existe
  let hoja = ss.getSheetByName(CONFIG.HOJA);
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG.HOJA);
  }

  // Poner encabezados solo si la hoja está vacía
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
  }

  SpreadsheetApp.getUi().alert('✅ Sistema instalado\n\nAhora usa "🔄 Sincronizar Kobo" para traer participantes.');
}

// ============================================================================
// CONFIGURACIÓN
// ============================================================================

function cargarApiKey() {
  const guardada = PropertiesService.getUserProperties().getProperty('KOBO_API_KEY');
  if (guardada) CONFIG.KOBO_API_KEY = guardada;
}

function abrirConfiguracion() {
  const html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;padding:20px}label{display:block;font-weight:bold;margin:12px 0 4px}'
    + 'input,textarea{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;font-family:monospace}'
    + 'textarea{height:80px}button{margin-top:14px;background:#1f73e6;color:#fff;padding:10px 20px;border:none;border-radius:4px;cursor:pointer;width:100%}'
    + '.ok{background:#c8e6c9;padding:10px;border-radius:4px;color:#2e7d32;margin-bottom:10px}</style>'
    + '<div class="ok">✅ API Key configurada en el código</div>'
    + '<label>ID Carpeta Drive</label>'
    + '<input id="cid" value="' + CONFIG.FOLDER_PARTICIPANTES_ID + '">'
    + '<label>API Key Kobo</label>'
    + '<textarea id="key">' + CONFIG.KOBO_API_KEY + '</textarea>'
    + '<button onclick="var c=document.getElementById(\'cid\').value,k=document.getElementById(\'key\').value;'
    + 'google.script.run.guardarConfiguracion(c,k);alert(\'✅ Guardado\')">💾 Guardar</button>'
  );
  SpreadsheetApp.getUi().showSidebar(html);
}

function guardarConfiguracion(carpetaId, apiKey) {
  PropertiesService.getUserProperties()
    .setProperty('KOBO_API_KEY', apiKey)
    .setProperty('FOLDER_ID', carpetaId);
  CONFIG.KOBO_API_KEY = apiKey;
  CONFIG.FOLDER_PARTICIPANTES_ID = carpetaId;
}

// ============================================================================
// PROBAR KOBO
// ============================================================================

function probarKobo() {
  cargarApiKey();
  const resp = UrlFetchApp.fetch(
    CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/',
    { headers: { Authorization: 'Token ' + CONFIG.KOBO_API_KEY }, muteHttpExceptions: true }
  );
  const code = resp.getResponseCode();
  if (code === 200) {
    const n = JSON.parse(resp.getContentText()).deployment__submission_count || 0;
    SpreadsheetApp.getUi().alert('✅ Conexión exitosa\n\n' + n + ' respuestas en Kobo');
  } else {
    SpreadsheetApp.getUi().alert('❌ Error ' + code + '\n\nVerifica la API Key en ⚙️ Configuración');
  }
}

// ============================================================================
// SINCRONIZAR KOBO
// ============================================================================

function sincronizar() {
  cargarApiKey();
  const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
  if (!hoja) {
    SpreadsheetApp.getUi().alert('❌ Primero instala el sistema (📥 INSTALAR)');
    return;
  }

  // Descargar datos de Kobo
  const resp = UrlFetchApp.fetch(
    CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json',
    { headers: { Authorization: 'Token ' + CONFIG.KOBO_API_KEY }, muteHttpExceptions: true }
  );

  if (resp.getResponseCode() !== 200) {
    SpreadsheetApp.getUi().alert('❌ Error Kobo ' + resp.getResponseCode() + '\n\nVerifica tu API Key');
    return;
  }

  const registros = JSON.parse(resp.getContentText()).results || [];
  if (registros.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No hay respuestas en Kobo');
    return;
  }

  // IDs ya existentes
  const existentes = new Set();
  if (hoja.getLastRow() > 1) {
    hoja.getRange(2, CONFIG.COL.ID, hoja.getLastRow() - 1, 1)
      .getValues().forEach(r => { if (r[0]) existentes.add(String(r[0])); });
  }

  let agregados = 0;
  registros.forEach(r => {
    const id = r['creamos_id'] || r['_id'] || '';
    if (!id || existentes.has(String(id))) return;

    const fila = mapearFila(r);
    hoja.appendRow(fila);
    colorearFila(hoja, hoja.getLastRow(), fila[CONFIG.COL.PERFIL - 1]);
    agregados++;
  });

  SpreadsheetApp.getUi().alert('✅ Sincronización completada\n\n' + agregados + ' participantes nuevos de ' + registros.length + ' en Kobo');
}

function mapearFila(r) {
  return [
    r['creamos_id']                                          || '',  // A ID
    new Date(),                                                       // B Fecha
    r['nombre_completo_del_la_participante']                 || '',  // C Nombre
    r['numero_de_dpi_opcional']                              || '',  // D DPI
    r['edad']                                                || '',  // E Edad
    r['genero']                                              || '',  // F Género
    r['numero_de_telefono']                                  || '',  // G Teléfono
    r['lugar_de_residencia']                                 || '',  // H Zona
    r['correo_electronico_opcional']                         || '',  // I Email
    r['cual_es_el_ultimo_grado_que_completaste']             || '',  // J Educación
    r['cual_es_tu_situacion_laboral_actual']                 || '',  // K Laboral
    r['que_sabes_hacer_bien']                                || '',  // L Fortalezas
    r['que_tipo_de_empleo_estas_buscando_especificamente']   || '',  // M Objetivo
    r['perfil_asignado']                                     || '',  // N Perfil
    r['prioridad_caso']                                      || '',  // O Prioridad
    r['puntaje_total_60']                                    || 0,   // P Puntaje
    r['dimension_1_capital_educativo']                       || 0,   // Q Dim1
    r['dimension_2_capital_laboral']                         || 0,   // R Dim2
    r['dimension_3_habilidades_digitales']                   || 0,   // S Dim3
    r['dimension_4_claridad_vocacional']                     || 0,   // T Dim4
    r['dimension_5_barreras_estructurales']                  || 0,   // U Dim5
    r['dimension_6_red_apoyo']                               || 0,   // V Dim6
    'Orientación',                                                    // W Estado
    '',                                                               // X Carpeta ID
    '',                                                               // Y Doc ID
    ''                                                                // Z Doc URL
  ];
}

// ============================================================================
// COLORES
// ============================================================================

function colorearTabla() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
  if (!hoja || hoja.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('⚠️ No hay datos para colorear');
    return;
  }
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 26).getValues();
  datos.forEach((fila, i) => {
    colorearFila(hoja, i + 2, fila[CONFIG.COL.PERFIL - 1]);
  });
  SpreadsheetApp.getUi().alert('✅ Tabla coloreada\n🟢 A  🔵 B  🟡 C  🔴 D');
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
// FICHA PARTICIPANTE — SIDEBAR (NO ventana flotante)
// ============================================================================

function abrirFichaParticipante(datos) {
  const nombre    = datos[CONFIG.COL.NOMBRE    - 1] || '—';
  const dpi       = datos[CONFIG.COL.DPI       - 1] || '—';
  const edad      = datos[CONFIG.COL.EDAD      - 1] || '—';
  const genero    = datos[CONFIG.COL.GENERO    - 1] || '—';
  const telefono  = datos[CONFIG.COL.TELEFONO  - 1] || '—';
  const zona      = datos[CONFIG.COL.ZONA      - 1] || '—';
  const perfil    = datos[CONFIG.COL.PERFIL    - 1] || '—';
  const prioridad = datos[CONFIG.COL.PRIORIDAD - 1] || '—';
  const puntaje   = datos[CONFIG.COL.PUNTAJE   - 1] || 0;
  const estado    = datos[CONFIG.COL.ESTADO    - 1] || 'Orientación';
  const objetivo  = datos[CONFIG.COL.OBJETIVO  - 1] || '—';
  const fortalezas = datos[CONFIG.COL.FORTALEZAS - 1] || '—';
  const d1 = Number(datos[CONFIG.COL.DIM1 - 1]) || 0;
  const d2 = Number(datos[CONFIG.COL.DIM2 - 1]) || 0;
  const d3 = Number(datos[CONFIG.COL.DIM3 - 1]) || 0;
  const d4 = Number(datos[CONFIG.COL.DIM4 - 1]) || 0;
  const d5 = Number(datos[CONFIG.COL.DIM5 - 1]) || 0;
  const d6 = Number(datos[CONFIG.COL.DIM6 - 1]) || 0;

  const c = CONFIG.COLORES[perfil] || { fondo: '#f5f5f5', texto: '#333' };
  const cPrioridad = prioridad === 'CRÍTICO' ? '#c62828'
    : prioridad === 'ALTO'  ? '#e65100'
    : prioridad === 'MEDIO' ? '#f9a825' : '#388e3c';

  const iniciales = nombre.split(' ').slice(0, 2).map(p => p[0]).join('').toUpperCase();
  const radar = svgRadar([d1, d2, d3, d4, d5, d6], c.texto);
  const barras = svgBarras([d1, d2, d3, d4, d5, d6], c.texto);

  const html = HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:Arial,sans-serif;font-size:13px;background:#f8f9fa;color:#333}

  .avatar{width:64px;height:64px;border-radius:50%;background:` + c.texto + `;
    color:#fff;font-size:22px;font-weight:bold;display:flex;align-items:center;
    justify-content:center;flex-shrink:0}

  .header{background:` + c.fondo + `;padding:14px;display:flex;gap:12px;align-items:center;
    border-bottom:3px solid ` + c.texto + `}
  .header-info{flex:1;min-width:0}
  .nombre{font-size:15px;font-weight:bold;color:` + c.texto + `;
    white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .badges{display:flex;gap:6px;margin-top:6px;flex-wrap:wrap}
  .badge{padding:2px 8px;border-radius:12px;font-size:11px;font-weight:bold;
    border:1.5px solid ` + c.texto + `;color:` + c.texto + `}
  .badge-prio{background:` + cPrioridad + `;color:#fff;border-color:` + cPrioridad + `}
  .badge-estado{background:#fff;border-color:#999;color:#555}

  .puntaje{text-align:center;padding:10px;background:#fff;border-bottom:1px solid #eee}
  .puntaje-num{font-size:32px;font-weight:bold;color:` + c.texto + `}
  .puntaje-sub{font-size:11px;color:#999}

  .sec{padding:10px 14px;border-bottom:1px solid #eee;background:#fff}
  .sec+.sec{margin-top:6px}
  .sec-titulo{font-size:10px;font-weight:bold;color:#999;text-transform:uppercase;
    letter-spacing:.8px;margin-bottom:8px}
  .fila{display:flex;justify-content:space-between;padding:3px 0}
  .fila-k{color:#777}
  .fila-v{font-weight:bold;text-align:right;max-width:55%;word-break:break-word}

  .radar-wrap{display:flex;justify-content:center;padding:8px 0}
  .gap{margin-top:6px}
</style></head><body>

<div class="header">
  <div class="avatar">` + iniciales + `</div>
  <div class="header-info">
    <div class="nombre">` + nombre + `</div>
    <div class="badges">
      <span class="badge">` + perfil + `</span>
      <span class="badge badge-prio">` + prioridad + `</span>
      <span class="badge badge-estado">` + estado + `</span>
    </div>
  </div>
</div>

<div class="puntaje">
  <div class="puntaje-num">` + puntaje + `<span style="font-size:14px;color:#aaa">/60</span></div>
  <div class="puntaje-sub">Puntaje diagnóstico</div>
</div>

<div class="gap"></div>

<div class="sec">
  <div class="sec-titulo">Datos personales</div>
  <div class="fila"><span class="fila-k">DPI</span><span class="fila-v">` + dpi + `</span></div>
  <div class="fila"><span class="fila-k">Edad</span><span class="fila-v">` + edad + `</span></div>
  <div class="fila"><span class="fila-k">Género</span><span class="fila-v">` + genero + `</span></div>
  <div class="fila"><span class="fila-k">Teléfono</span><span class="fila-v">` + telefono + `</span></div>
  <div class="fila"><span class="fila-k">Zona</span><span class="fila-v">` + zona + `</span></div>
</div>

<div class="sec">
  <div class="sec-titulo">Objetivo laboral</div>
  <div style="color:#444;line-height:1.5">` + objetivo + `</div>
</div>

<div class="sec">
  <div class="sec-titulo">Fortalezas</div>
  <div style="color:#444;line-height:1.5">` + fortalezas + `</div>
</div>

<div class="sec">
  <div class="sec-titulo">Perfil por dimensiones</div>
  <div class="radar-wrap">` + radar + `</div>
  ` + barras + `
</div>

</body></html>`);

  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================================
// SVG RADAR + BARRAS
// ============================================================================

function svgRadar(dims, color) {
  const cx = 85, cy = 85, r = 65;
  const angulos = [-90, -30, 30, 90, 150, 210];
  const labels  = ['Educ','Labor','Digital','Vocal','Barr','Apoyo'];

  let fondos = '';
  [.33, .66, 1].forEach(function(n) {
    const pts = angulos.map(function(a) {
      const rad = a * Math.PI / 180;
      return (cx + r * n * Math.cos(rad)).toFixed(1) + ',' + (cy + r * n * Math.sin(rad)).toFixed(1);
    }).join(' ');
    fondos += '<polygon points="' + pts + '" fill="none" stroke="#e0e0e0" stroke-width="1"/>';
  });

  let ejes = angulos.map(function(a) {
    const rad = a * Math.PI / 180;
    return '<line x1="' + cx + '" y1="' + cy + '" x2="' + (cx + r * Math.cos(rad)).toFixed(1)
      + '" y2="' + (cy + r * Math.sin(rad)).toFixed(1) + '" stroke="#e0e0e0" stroke-width="1"/>';
  }).join('');

  const dataPts = angulos.map(function(a, i) {
    const rad = a * Math.PI / 180;
    const esc = (dims[i] / 10);
    return (cx + r * esc * Math.cos(rad)).toFixed(1) + ',' + (cy + r * esc * Math.sin(rad)).toFixed(1);
  }).join(' ');

  let lbls = angulos.map(function(a, i) {
    const rad = a * Math.PI / 180;
    const x = (cx + (r + 14) * Math.cos(rad)).toFixed(1);
    const y = (cy + (r + 14) * Math.sin(rad)).toFixed(1);
    return '<text x="' + x + '" y="' + y + '" text-anchor="middle" font-size="9" fill="#777">' + labels[i] + '</text>';
  }).join('');

  return '<svg width="170" height="170" viewBox="0 0 170 170">'
    + fondos + ejes
    + '<polygon points="' + dataPts + '" fill="' + color + '33" stroke="' + color + '" stroke-width="2" stroke-linejoin="round"/>'
    + lbls + '</svg>';
}

function svgBarras(dims, color) {
  const labels = ['Educativo','Laboral','Digital','Vocacional','Barreras','Red Apoyo'];
  return dims.map(function(v, i) {
    const pct = Math.round((v / 10) * 100);
    return '<div style="margin:5px 0">'
      + '<div style="display:flex;justify-content:space-between;font-size:11px;color:#666;margin-bottom:2px">'
      + '<span>' + labels[i] + '</span><span>' + v + '/10</span></div>'
      + '<div style="background:#eee;border-radius:4px;height:7px">'
      + '<div style="width:' + pct + '%;background:' + color + ';height:7px;border-radius:4px"></div>'
      + '</div></div>';
  }).join('');
}

// ============================================================================
// FIN
// ============================================================================
