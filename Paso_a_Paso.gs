// ============================================================================
// SISTEMA PASO A PASO - FITO v5.1
// TOP: Automatizado + Fácil + Sin errores
// ============================================================================

const CONFIG = {
  // Google
  SPREADSHEET_ID:        SpreadsheetApp.getActive().getId(),
  FOLDER_PARTICIPANTES_ID: "1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU",

  // Kobo - YA CONFIGURADO
  KOBO_API_KEY:  "64cc018b88067397addd36b09288be8b6539cf39",
  KOBO_ASSET_ID: "abHRWdRnPhKwzPQBajc7RZ",
  KOBO_URL:      "https://kf.kobotoolbox.org/api/v2",

  HOJA: "Maestro",

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
  SpreadsheetApp.getUi()
    .createMenu('📊 PASO A PASO')
    .addItem('📥 Instalar', 'instalar')
    .addSeparator()
    .addItem('🔄 Sincronizar Kobo', 'sincronizar')
    .addItem('🎨 Colorear Tabla', 'colorearTabla')
    .addItem('📋 Ver Ficha', 'verFicha')
    .addSeparator()
    .addItem('🧪 Probar Kobo', 'probarKobo')
    .addItem('⚙️ Configuración', 'abrirConfiguracion')
    .addToUi();
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

    // Crear encabezados si está vacía
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

    SpreadsheetApp.getUi().alert('✅ Sistema instalado\n\nAhora:\n1. 🔄 Sincronizar Kobo\n2. 🎨 Colorear Tabla\n3. 📋 Ver Ficha');
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================================================
// SINCRONIZAR KOBO
// ============================================================================

function sincronizar() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja) {
      SpreadsheetApp.getUi().alert('❌ Instala primero: 📥 Instalar');
      return;
    }

    // Obtener datos de Kobo
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json',
      {
        headers: { Authorization: 'Token ' + CONFIG.KOBO_API_KEY },
        muteHttpExceptions: true
      }
    );

    if (resp.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert('❌ Error Kobo ' + resp.getResponseCode());
      return;
    }

    const registros = JSON.parse(resp.getContentText()).results || [];
    if (!registros.length) {
      SpreadsheetApp.getUi().alert('⚠️ Sin respuestas en Kobo');
      return;
    }

    // IDs existentes
    const existentes = new Set();
    if (hoja.getLastRow() > 1) {
      hoja.getRange(2, CONFIG.COL.ID, hoja.getLastRow() - 1, 1)
        .getValues().forEach(r => { if (r[0]) existentes.add(String(r[0])); });
    }

    let agregados = 0;
    registros.forEach(r => {
      const id = r['creamos_id'] || r['_id'] || '';
      if (!id || existentes.has(String(id))) return;

      const fila = [
        id || '', new Date(),
        r['nombre_completo_del_la_participante'] || '',
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
        r['puntaje_total_60'] || 0,
        r['dimension_1_capital_educativo'] || 0,
        r['dimension_2_capital_laboral'] || 0,
        r['dimension_3_habilidades_digitales'] || 0,
        r['dimension_4_claridad_vocacional'] || 0,
        r['dimension_5_barreras_estructurales'] || 0,
        r['dimension_6_red_apoyo'] || 0,
        'Orientación', '', '', ''
      ];

      hoja.appendRow(fila);
      colorearFila(hoja, hoja.getLastRow(), fila[CONFIG.COL.PERFIL - 1]);
      agregados++;
    });

    SpreadsheetApp.getUi().alert('✅ Sincronizado\n\n' + agregados + ' participantes nuevos\nde ' + registros.length + ' en Kobo');
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================================================
// COLOREAR TABLA
// ============================================================================

function colorearTabla() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA);
    if (!hoja || hoja.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Sin datos');
      return;
    }

    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 26).getValues();
    datos.forEach((fila, i) => {
      colorearFila(hoja, i + 2, fila[CONFIG.COL.PERFIL - 1]);
    });

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
    const ss = SpreadsheetApp.getActive();
    const rango = ss.getActiveRange();
    const hoja = rango.getSheet();

    if (hoja.getName() !== CONFIG.HOJA) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila en la hoja Maestro');
      return;
    }

    const fila = rango.getRow();
    if (fila < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila de datos (no encabezado)');
      return;
    }

    const datos = hoja.getRange(fila, 1, 1, 26).getValues()[0];
    if (!datos[CONFIG.COL.ID - 1]) {
      SpreadsheetApp.getUi().alert('⚠️ Fila vacía');
      return;
    }

    mostrarFicha(datos);
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

function mostrarFicha(datos) {
  const nombre    = datos[CONFIG.COL.NOMBRE - 1]     || '—';
  const dpi       = datos[CONFIG.COL.DPI - 1]        || '—';
  const edad      = datos[CONFIG.COL.EDAD - 1]       || '—';
  const genero    = datos[CONFIG.COL.GENERO - 1]     || '—';
  const telefono  = datos[CONFIG.COL.TELEFONO - 1]   || '—';
  const zona      = datos[CONFIG.COL.ZONA - 1]       || '—';
  const perfil    = datos[CONFIG.COL.PERFIL - 1]     || '—';
  const prioridad = datos[CONFIG.COL.PRIORIDAD - 1]  || '—';
  const puntaje   = datos[CONFIG.COL.PUNTAJE - 1]    || 0;
  const estado    = datos[CONFIG.COL.ESTADO - 1]     || 'Orientación';
  const objetivo  = datos[CONFIG.COL.OBJETIVO - 1]   || '—';
  const fortalezas = datos[CONFIG.COL.FORTALEZAS - 1] || '—';
  const d1 = Number(datos[CONFIG.COL.DIM1 - 1]) || 0;
  const d2 = Number(datos[CONFIG.COL.DIM2 - 1]) || 0;
  const d3 = Number(datos[CONFIG.COL.DIM3 - 1]) || 0;
  const d4 = Number(datos[CONFIG.COL.DIM4 - 1]) || 0;
  const d5 = Number(datos[CONFIG.COL.DIM5 - 1]) || 0;
  const d6 = Number(datos[CONFIG.COL.DIM6 - 1]) || 0;

  const c = CONFIG.COLORES[perfil] || { fondo: '#f5f5f5', texto: '#333' };
  const cPrio = prioridad === 'CRÍTICO' ? '#c62828'
    : prioridad === 'ALTO'  ? '#e65100'
    : prioridad === 'MEDIO' ? '#f9a825' : '#388e3c';

  const iniciales = nombre.split(' ').slice(0, 2).map(p => p.charAt(0)).join('').toUpperCase();
  const radar = svgRadar([d1, d2, d3, d4, d5, d6], c.texto);
  const barras = svgBarras([d1, d2, d3, d4, d5, d6], c.texto);

  const html = HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html><head><meta charset="UTF-8"><style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:Arial;font-size:13px;background:#f8f9fa;color:#333;overflow-y:auto}
.header{background:` + c.fondo + `;padding:16px;border-bottom:3px solid ` + c.texto + `;display:flex;gap:12px;align-items:center}
.avatar{width:60px;height:60px;border-radius:50%;background:` + c.texto + `;color:#fff;font-size:20px;font-weight:bold;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.header-info{flex:1}
.nombre{font-size:15px;font-weight:bold;color:` + c.texto + `}
.badges{display:flex;gap:6px;margin-top:6px;flex-wrap:wrap}
.badge{padding:3px 10px;border-radius:12px;font-size:11px;font-weight:bold;border:1.5px solid ` + c.texto + `;color:` + c.texto + `}
.badge-prio{background:` + cPrio + `;border-color:` + cPrio + `;color:#fff}
.badge-estado{background:#fff;color:#666}
.puntaje{text-align:center;padding:12px;background:#fff;border-bottom:1px solid #eee}
.puntaje-num{font-size:32px;font-weight:bold;color:` + c.texto + `}
.puntaje-sub{font-size:10px;color:#999;margin-top:2px}
.gap{margin:6px 0}
.sec{padding:12px 14px;background:#fff;border-bottom:1px solid #eee}
.sec-titulo{font-size:10px;font-weight:bold;color:#999;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px}
.fila{display:flex;justify-content:space-between;padding:4px 0;font-size:13px}
.fila-k{color:#777;flex:1}
.fila-v{font-weight:bold;text-align:right;max-width:50%;word-break:break-word}
.radar-wrap{display:flex;justify-content:center;padding:8px 0}
.desc{color:#555;line-height:1.5;font-size:13px}
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
  <div class="puntaje-sub">Diagnóstico</div>
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
  <div class="desc">` + objetivo + `</div>
</div>

<div class="sec">
  <div class="sec-titulo">Fortalezas</div>
  <div class="desc">` + fortalezas + `</div>
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
// CONFIGURACIÓN
// ============================================================================

function abrirConfiguracion() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial;padding:16px;background:#f5f5f5}
      label{display:block;font-weight:bold;margin:12px 0 4px;color:#333}
      input,textarea{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;font-family:monospace;font-size:12px}
      textarea{height:70px;resize:vertical}
      button{margin-top:14px;background:#1f73e6;color:#fff;padding:10px;border:none;border-radius:4px;cursor:pointer;width:100%;font-weight:bold}
      button:hover{background:#1557b0}
      .info{background:#c8e6c9;color:#2e7d32;padding:8px;border-radius:4px;margin-bottom:10px;font-size:12px}
    </style>
    <div class="info">✅ API Key ya configurada en el código</div>
    <label>ID Carpeta Drive:</label>
    <input id="cid" value="1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU">
    <label>API Key (si necesitas cambiar):</label>
    <textarea id="key">64cc018b88067397addd36b09288be8b6539cf39</textarea>
    <button onclick="var c=document.getElementById('cid').value,k=document.getElementById('key').value;google.script.run.guardarConfig(c,k);alert('✅ Guardado')">Guardar</button>
  `);
  SpreadsheetApp.getUi().showSidebar(html);
}

function guardarConfig(carpeta, apiKey) {
  PropertiesService.getUserProperties()
    .setProperty('FOLDER_ID', carpeta)
    .setProperty('KOBO_API_KEY', apiKey);
  CONFIG.FOLDER_PARTICIPANTES_ID = carpeta;
  CONFIG.KOBO_API_KEY = apiKey;
}

// ============================================================================
// PROBAR KOBO
// ============================================================================

function probarKobo() {
  try {
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/',
      { headers: { Authorization: 'Token ' + CONFIG.KOBO_API_KEY }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() === 200) {
      const n = JSON.parse(resp.getContentText()).deployment__submission_count || 0;
      SpreadsheetApp.getUi().alert('✅ Conexión OK\n\n📊 ' + n + ' respuestas en Kobo');
    } else {
      SpreadsheetApp.getUi().alert('❌ Error ' + resp.getResponseCode() + '\n\nVerifica API Key en ⚙️ Configuración');
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================================================
// SVG RADAR Y BARRAS
// ============================================================================

function svgRadar(dims, color) {
  const cx = 85, cy = 85, r = 65;
  const angulos = [-90, -30, 30, 90, 150, 210];
  const labels = ['Educ', 'Labor', 'Digit', 'Vocal', 'Barr', 'Apoyo'];

  let fondos = '';
  [.33, .66, 1].forEach(n => {
    const pts = angulos.map(a => {
      const rad = a * Math.PI / 180;
      return (cx + r * n * Math.cos(rad)).toFixed(0) + ',' + (cy + r * n * Math.sin(rad)).toFixed(0);
    }).join(' ');
    fondos += '<polygon points="' + pts + '" fill="none" stroke="#e0e0e0" stroke-width="1"/>';
  });

  let ejes = angulos.map(a => {
    const rad = a * Math.PI / 180;
    const x2 = (cx + r * Math.cos(rad)).toFixed(0);
    const y2 = (cy + r * Math.sin(rad)).toFixed(0);
    return '<line x1="' + cx + '" y1="' + cy + '" x2="' + x2 + '" y2="' + y2 + '" stroke="#e0e0e0" stroke-width="1"/>';
  }).join('');

  const dataPts = angulos.map((a, i) => {
    const rad = a * Math.PI / 180;
    const esc = dims[i] / 10;
    return (cx + r * esc * Math.cos(rad)).toFixed(0) + ',' + (cy + r * esc * Math.sin(rad)).toFixed(0);
  }).join(' ');

  let lbls = angulos.map((a, i) => {
    const rad = a * Math.PI / 180;
    const x = (cx + (r + 14) * Math.cos(rad)).toFixed(0);
    const y = (cy + (r + 14) * Math.sin(rad)).toFixed(0);
    return '<text x="' + x + '" y="' + y + '" text-anchor="middle" font-size="8" fill="#777">' + labels[i] + '</text>';
  }).join('');

  return '<svg width="170" height="170" viewBox="0 0 170 170">' + fondos + ejes
    + '<polygon points="' + dataPts + '" fill="' + color + '33" stroke="' + color + '" stroke-width="2"/>'
    + lbls + '</svg>';
}

function svgBarras(dims, color) {
  const labels = ['Educativo', 'Laboral', 'Digital', 'Vocacional', 'Barreras', 'Red Apoyo'];
  return dims.map((v, i) => {
    const pct = Math.round((v / 10) * 100);
    return '<div style="margin:5px 0">'
      + '<div style="display:flex;justify-content:space-between;font-size:11px;color:#666;margin-bottom:2px">'
      + '<span>' + labels[i] + '</span><span>' + v + '/10</span></div>'
      + '<div style="background:#eee;border-radius:3px;height:6px;overflow:hidden">'
      + '<div style="width:' + pct + '%;background:' + color + ';height:6px"></div></div></div>';
  }).join('');
}

// ============================================================================
// FIN
// ============================================================================
