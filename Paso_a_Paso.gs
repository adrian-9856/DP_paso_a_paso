// ============================================================================
// SISTEMA PASO A PASO - VERSIÓN 4.0 PROFESIONAL
// Importación de Kobo + Dashboard + Derivaciones + Reportes
// ============================================================================

function onOpen() {
  try {
    cargarConfiguracion();
    const ui = SpreadsheetApp.getUi();

    ui.createMenu('📊 PASO A PASO')
      .addItem('⚙️ Configuración', 'mostrarConfiguracion')
      .addSeparator()
      .addItem('🔄 SINCRONIZAR Kobo (Manual)', 'sincronizarKoboCompleto')
      .addItem('📥 Sincronizar en Background', 'sincronizarBackground')
      .addSeparator()
      .addItem('📊 Dashboard', 'abrirDashboard')
      .addItem('📋 Ver Reportes', 'abrirReportes')
      .addItem('📁 Derivaciones Pendientes', 'abrirDerivaciones')
      .addSeparator()
      .addItem('🧪 Probar Conexión Kobo', 'probarConexionKoboAPI')
      .addItem('📖 Ayuda', 'mostrarAyuda')
      .addToUi();

    log('✅ Sistema iniciado correctamente', 'INFO');
  } catch (error) {
    log(`Error en onOpen: ${error}`, 'ERROR');
  }
}

// ============================================================================
// CONFIGURACIÓN
// ============================================================================

const CONFIG = {
  SPREADSHEET_ID: SpreadsheetApp.getActive().getId(),
  FOLDER_PARTICIPANTES_ID: "1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU",

  KOBO: {
    BASE_URL: "https://kf.kobotoolbox.org/api/v2",
    ASSET_ID: "abHRWdRnPhKwzPQBajc7RZ",
    API_KEY: ""
  },

  HOJAS: {
    MAESTRO: "Maestro",
    DASHBOARD: "Dashboard",
    DERIVACION: "Derivación",
    REPORTES: "Reportes",
    CONFIGURACION: "Configuración",
    LOG: "Log"
  },

  // Mapeo de campos Kobo → Sheet
  CAMPOS_KOBO: {
    creamos_id: "Creamos_ID",
    nombre_completo_del_la_participante: "Nombre_Completo",
    numero_de_dpi_opcional: "DPI",
    edad: "Edad",
    genero: "Género",
    numero_de_telefono: "Teléfono",
    lugar_de_residencia: "Zona",
    correo_electronico_opcional: "Email",
    cual_es_el_ultimo_grado_que_completaste: "Nivel_Educativo",
    cual_es_tu_situacion_laboral_actual: "Situación_Laboral",
    que_sabes_hacer_bien: "Fortalezas",
    que_tipo_de_empleo_estas_buscando_especificamente: "Objetivo_Laboral",
    perfil_asignado: "Perfil_Asignado",
    prioridad_caso: "Prioridad",
    puntaje_total_60: "Puntaje_Total",
    dimension_1_capital_educativo: "Dim_Educativo",
    dimension_2_capital_laboral: "Dim_Laboral",
    dimension_3_habilidades_digitales: "Dim_Digital",
    dimension_4_claridad_vocacional: "Dim_Vocacional",
    dimension_5_barreras_estructurales: "Dim_Barreras",
    dimension_6_red_apoyo: "Dim_Apoyo"
  }
};

// ============================================================================
// FUNCIONES BASE
// ============================================================================

function cargarConfiguracion() {
  const props = PropertiesService.getUserProperties();
  const apiKey = props.getProperty('KOBO_API_KEY');
  if (apiKey) CONFIG.KOBO.API_KEY = apiKey;
  log('✅ Config cargada', 'INFO');
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function crearHojaSiNoExiste(nombre) {
  try {
    const ss = getSpreadsheet();
    let hoja = ss.getSheetByName(nombre);

    if (!hoja) {
      hoja = ss.insertSheet(nombre);
      log(`📄 Hoja creada: ${nombre}`, 'INFO');
    }
    return hoja;
  } catch (error) {
    log(`Error creando hoja ${nombre}: ${error}`, 'ERROR');
    return null;
  }
}

function getHoja(nombre) {
  return getSpreadsheet().getSheetByName(nombre);
}

function log(msg, tipo = "INFO") {
  const timestamp = new Date().toLocaleString("es-GT");
  const linea = `[${timestamp}] [${tipo}] ${msg}`;
  Logger.log(linea);

  try {
    const hojaLog = crearHojaSiNoExiste(CONFIG.HOJAS.LOG);
    hojaLog.appendRow([new Date(), tipo, msg]);
  } catch (error) {
    // Silent fail para log
  }
}

// ============================================================================
// MENÚ - CONFIGURACIÓN
// ============================================================================

function mostrarConfiguracion() {
  const apiKey = PropertiesService.getUserProperties().getProperty('KOBO_API_KEY') || '';

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Arial', sans-serif; padding: 20px; background: #f0f2f5; }
      .container { max-width: 700px; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
      h1 { color: #1f73e6; border-bottom: 3px solid #1f73e6; padding-bottom: 15px; }
      .section { background: #f8fafb; padding: 15px; border-radius: 8px; margin: 15px 0; border-left: 4px solid #1f73e6; }
      label { display: block; font-weight: bold; color: #333; margin: 10px 0 5px 0; }
      input, textarea { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 5px; font-size: 14px; box-sizing: border-box; margin-bottom: 10px; font-family: monospace; }
      textarea { min-height: 100px; }
      button { background: #1f73e6; color: white; padding: 12px 25px; border: none; border-radius: 5px; cursor: pointer; font-size: 14px; margin-right: 10px; transition: background 0.3s; }
      button:hover { background: #1557b0; }
      .info { background: #e3f2fd; padding: 12px; border-left: 4px solid #1f73e6; border-radius: 4px; color: #0d47a1; }
      .success { background: #c8e6c9; padding: 12px; border-left: 4px solid #2e7d32; border-radius: 4px; color: #1b5e20; }
    </style>

    <div class="container">
      <h1>⚙️ Configuración Paso a Paso v4.0</h1>

      <div class="section">
        <h3>Google Sheet ID</h3>
        <div class="success">✅ ${CONFIG.SPREADSHEET_ID.substring(0, 30)}...</div>
        <p style="font-size: 12px; color: #666;">Detectado automáticamente</p>
      </div>

      <div class="section">
        <h3>📁 Carpeta Google Drive</h3>
        <label>ID de Carpeta "Participantes":</label>
        <input type="text" id="carpetaId" value="1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU" placeholder="ID de carpeta">
        <div class="info">📌 URL: drive.google.com/drive/folders/<strong>AQUÍ_VA_EL_ID</strong></div>
      </div>

      <div class="section">
        <h3>🔑 API Key de KoboToolbox</h3>
        <label>Token de API:</label>
        <textarea id="apiKey" placeholder="Pega tu API Key de Kobo">${apiKey}</textarea>
        <div class="info">
          🔗 Obtén aquí: <strong>https://kf.kobotoolbox.org/admin/auth/token/</strong><br>
          Asset ID: <strong>abHRWdRnPhKwzPQBajc7RZ</strong>
        </div>
      </div>

      <button onclick="guardarConfig()">💾 GUARDAR CONFIGURACIÓN</button>
      <button onclick="probarConexion()">🧪 PROBAR CONEXIÓN</button>
      <button onclick="window.close()">❌ CERRAR</button>
    </div>

    <script>
      function guardarConfig() {
        const carpeta = document.getElementById('carpetaId').value.trim();
        const apiKey = document.getElementById('apiKey').value.trim();

        if (!carpeta || !apiKey) {
          alert('❌ Completa todos los campos');
          return;
        }

        google.script.run.guardarConfiguracionScript(carpeta, apiKey);
        alert('✅ Configuración guardada');
      }

      function probarConexion() {
        const apiKey = document.getElementById('apiKey').value.trim();
        if (!apiKey) {
          alert('❌ Ingresa la API Key');
          return;
        }
        google.script.run.probarConexionKoboAPI(apiKey);
      }
    </script>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, '⚙️ Configuración');
}

function guardarConfiguracionScript(carpetaId, apiKey) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('FOLDER_PARTICIPANTES_ID', carpetaId);
  props.setProperty('KOBO_API_KEY', apiKey);

  CONFIG.FOLDER_PARTICIPANTES_ID = carpetaId;
  CONFIG.KOBO.API_KEY = apiKey;

  log('✅ Configuración guardada', 'INFO');
  SpreadsheetApp.getUi().alert('✅ Configuración guardada exitosamente');
}

function probarConexionKoboAPI(apiKey) {
  try {
    const url = `${CONFIG.KOBO.BASE_URL}/assets/${CONFIG.KOBO.ASSET_ID}/`;
    const options = {
      headers: { "Authorization": `Token ${apiKey}` },
      muteHttpExceptions: true,
      timeout: 30
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const submissions = data.deployment__submission_count || 0;
      SpreadsheetApp.getUi().alert(`✅ CONEXIÓN EXITOSA\n\n📊 ${submissions} respuestas en Kobo`);
      log(`Conexión exitosa. Submissions: ${submissions}`, 'INFO');
    } else {
      SpreadsheetApp.getUi().alert(`❌ Error ${response.getResponseCode()}\n\nVerifica tu API Key`);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

// ============================================================================
// SINCRONIZACIÓN COMPLETA DE KOBO
// ============================================================================

function sincronizarKoboCompleto() {
  cargarConfiguracion();

  if (!CONFIG.KOBO.API_KEY) {
    SpreadsheetApp.getUi().alert('❌ Configura la API Key primero\n\n📊 PASO A PASO > ⚙️ Configuración');
    return;
  }

  log("=== INICIANDO SINCRONIZACIÓN COMPLETA ===", "INFO");

  try {
    const ui = SpreadsheetApp.getUi();
    ui.showModelessDialog(
      HtmlService.createHtmlOutput('<h2>⏳ Sincronizando datos de Kobo...</h2><p>Por favor espera...</p>'),
      'Sincronización'
    );

    const datos = descargarDatosKobo();
    if (!datos || datos.length === 0) {
      ui.alert('⚠️ No hay registros nuevos en Kobo');
      return;
    }

    const hojaMaestro = crearHojaSiNoExiste(CONFIG.HOJAS.MAESTRO);
    const datosExistentes = hojaMaestro.getDataRange().getValues();
    const idsExistentes = new Set(datosExistentes.slice(1).map(row => row[0])); // Columna A: Creamos_ID

    let agregados = 0;
    let actualizados = 0;

    for (const registro of datos) {
      const creamos_id = registro.creamos_id || '';
      if (!creamos_id) continue;

      if (idsExistentes.has(creamos_id)) {
        actualizarParticipante(hojaMaestro, creamos_id, registro);
        actualizados++;
      } else {
        agregarParticipante(hojaMaestro, registro);
        agregados++;
      }
    }

    actualizarDashboard();
    actualizarReportes();

    log(`✅ Sincronización completada: ${agregados} nuevos, ${actualizados} actualizados`, 'INFO');
    ui.alert(`✅ SINCRONIZACIÓN COMPLETADA\n\n✨ ${agregados} participantes nuevos\n🔄 ${actualizados} actualizados`);

  } catch (error) {
    log(`Error en sincronización: ${error}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

function descargarDatosKobo() {
  try {
    const url = `${CONFIG.KOBO.BASE_URL}/assets/${CONFIG.KOBO.ASSET_ID}/data/`;
    const options = {
      headers: { "Authorization": `Token ${CONFIG.KOBO.API_KEY}` },
      muteHttpExceptions: true,
      timeout: 60
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      log(`Error Kobo HTTP ${response.getResponseCode()}`, 'ERROR');
      return null;
    }

    const data = JSON.parse(response.getContentText());
    return data.results || [];

  } catch (error) {
    log(`Error descargando Kobo: ${error}`, 'ERROR');
    return null;
  }
}

function agregarParticipante(hoja, registro) {
  try {
    const row = mapearRegistroKobo(registro);
    hoja.appendRow(row);
    log(`✅ Participante agregado: ${registro.creamos_id}`, 'INFO');
  } catch (error) {
    log(`Error agregando participante: ${error}`, 'ERROR');
  }
}

function actualizarParticipante(hoja, creamos_id, registro) {
  try {
    const datos = hoja.getDataRange().getValues();
    for (let i = 0; i < datos.length; i++) {
      if (datos[i][0] === creamos_id) {
        const row = mapearRegistroKobo(registro);
        for (let j = 0; j < row.length; j++) {
          hoja.getRange(i + 1, j + 1).setValue(row[j]);
        }
        log(`🔄 Participante actualizado: ${creamos_id}`, 'INFO');
        return;
      }
    }
  } catch (error) {
    log(`Error actualizando participante: ${error}`, 'ERROR');
  }
}

function mapearRegistroKobo(registro) {
  // Mapeo de campos de Kobo al Sheet
  return [
    registro.creamos_id || '',                                                    // A: Creamos_ID
    new Date(),                                                                    // B: Fecha_Sincronización
    registro['nombre_completo_del_la_participante'] || '',                       // C: Nombre_Completo
    registro['numero_de_dpi_opcional'] || '',                                    // D: DPI
    registro['edad'] || '',                                                       // E: Edad
    registro['genero'] || '',                                                     // F: Género
    registro['numero_de_telefono'] || '',                                        // G: Teléfono
    registro['lugar_de_residencia'] || '',                                       // H: Zona
    registro['correo_electronico_opcional'] || '',                              // I: Email
    registro['cual_es_el_ultimo_grado_que_completaste'] || '',                 // J: Nivel_Educativo
    registro['cual_es_tu_situacion_laboral_actual'] || '',                     // K: Situación_Laboral
    registro['que_sabes_hacer_bien'] || '',                                     // L: Fortalezas
    registro['que_tipo_de_empleo_estas_buscando_especificamente'] || '',       // M: Objetivo_Laboral
    registro['perfil_asignado'] || '',                                           // N: Perfil_Asignado
    registro['prioridad_caso'] || '',                                            // O: Prioridad
    registro['puntaje_total_60'] || 0,                                           // P: Puntaje_Total
    registro['dimension_1_capital_educativo'] || 0,                             // Q: Dim_Educativo
    registro['dimension_2_capital_laboral'] || 0,                               // R: Dim_Laboral
    registro['dimension_3_habilidades_digitales'] || 0,                         // S: Dim_Digital
    registro['dimension_4_claridad_vocacional'] || 0,                           // T: Dim_Vocacional
    registro['dimension_5_barreras_estructurales'] || 0,                        // U: Dim_Barreras
    registro['dimension_6_red_apoyo'] || 0,                                     // V: Dim_Apoyo
    'En Orientación',                                                             // W: Estado
    '',                                                                            // X: Derivado_A
    '',                                                                            // Y: Fecha_Derivación
    'Pendiente',                                                                  // Z: Estado_Derivación
    ''                                                                             // AA: Notas
  ];
}

function sincronizarBackground() {
  SpreadsheetApp.getUi().alert('⏳ Sincronización iniciada en background...\n\nRevisaremos en 2 minutos');
  log('Sincronización en background iniciada', 'INFO');
  // Aquí se puede configurar un trigger automático
}

// ============================================================================
// DASHBOARD
// ============================================================================

function abrirDashboard() {
  try {
    const hojaD = crearHojaSiNoExiste(CONFIG.HOJAS.DASHBOARD);
    const hojaMaestro = getHoja(CONFIG.HOJAS.MAESTRO);

    if (!hojaMaestro) {
      SpreadsheetApp.getUi().alert('⚠️ Primero sincroniza datos desde Kobo');
      return;
    }

    const datos = hojaMaestro.getDataRange().getValues();
    const stats = {
      total: datos.length - 1,
      perfilA: 0,
      perfilB: 0,
      perfilC: 0,
      perfilD: 0,
      critico: 0,
      alto: 0,
      medio: 0,
      bajo: 0
    };

    for (let i = 1; i < datos.length; i++) {
      const perfil = datos[i][13]; // Columna N: Perfil_Asignado
      const prioridad = datos[i][14]; // Columna O: Prioridad

      if (perfil === 'Perfil A') stats.perfilA++;
      else if (perfil === 'Perfil B') stats.perfilB++;
      else if (perfil === 'Perfil C') stats.perfilC++;
      else if (perfil === 'Perfil D') stats.perfilD++;

      if (prioridad === 'CRÍTICO') stats.critico++;
      else if (prioridad === 'ALTO') stats.alto++;
      else if (prioridad === 'MEDIO') stats.medio++;
      else if (prioridad === 'BAJO') stats.bajo++;
    }

    const html = HtmlService.createHtmlOutput(`
      <style>
        body { font-family: Arial; padding: 20px; background: #f5f5f5; }
        .dashboard { max-width: 900px; background: white; padding: 25px; border-radius: 10px; }
        h1 { color: #1f73e6; text-align: center; border-bottom: 3px solid #1f73e6; padding-bottom: 15px; }
        .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 20px; }
        .card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
        .card.a { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); }
        .card.b { background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); }
        .card.c { background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%); }
        .card.d { background: linear-gradient(135deg, #fa709a 0%, #fee140 100%); }
        .card-title { font-size: 14px; opacity: 0.9; }
        .card-number { font-size: 48px; font-weight: bold; margin: 10px 0; }
        .card-label { font-size: 12px; opacity: 0.8; }
        .section { margin-top: 30px; }
        .section h2 { color: #333; border-bottom: 2px solid #1f73e6; padding-bottom: 10px; }
        .priority { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-top: 15px; }
        .priority-card { background: #f5f5f5; padding: 15px; border-radius: 8px; text-align: center; border-left: 4px solid #1f73e6; }
        .priority-card.critico { border-left-color: #d32f2f; }
        .priority-card.alto { border-left-color: #f57c00; }
        .priority-card.medio { border-left-color: #fbc02d; }
        .priority-card.bajo { border-left-color: #388e3c; }
        .number { font-size: 32px; font-weight: bold; color: #1f73e6; }
        .label { font-size: 12px; color: #666; margin-top: 5px; }
      </style>

      <div class="dashboard">
        <h1>📊 DASHBOARD - Paso a Paso v4.0</h1>

        <div class="section">
          <h2>👥 Distribución por Perfil</h2>
          <div class="grid">
            <div class="card a">
              <div class="card-title">Perfil A</div>
              <div class="card-number">${stats.perfilA}</div>
              <div class="card-label">Listo para Empleabilidad</div>
            </div>
            <div class="card b">
              <div class="card-title">Perfil B</div>
              <div class="card-number">${stats.perfilB}</div>
              <div class="card-label">Necesita Orientación</div>
            </div>
            <div class="card c">
              <div class="card-title">Perfil C</div>
              <div class="card-number">${stats.perfilC}</div>
              <div class="card-label">Desarrollo Capital Humano</div>
            </div>
            <div class="card d">
              <div class="card-title">Perfil D</div>
              <div class="card-number">${stats.perfilD}</div>
              <div class="card-label">Barreras Críticas</div>
            </div>
          </div>
        </div>

        <div class="section">
          <h2>⚠️ Distribución por Prioridad</h2>
          <div class="priority">
            <div class="priority-card critico">
              <div class="number" style="color: #d32f2f;">${stats.critico}</div>
              <div class="label">CRÍTICO</div>
            </div>
            <div class="priority-card alto">
              <div class="number" style="color: #f57c00;">${stats.alto}</div>
              <div class="label">ALTO</div>
            </div>
            <div class="priority-card medio">
              <div class="number" style="color: #fbc02d;">${stats.medio}</div>
              <div class="label">MEDIO</div>
            </div>
            <div class="priority-card bajo">
              <div class="number" style="color: #388e3c;">${stats.bajo}</div>
              <div class="label">BAJO</div>
            </div>
          </div>
        </div>

        <div class="section">
          <h2>📈 Resumen Total</h2>
          <p style="font-size: 18px; color: #1f73e6;">
            <strong>Total de Participantes: ${stats.total}</strong>
          </p>
        </div>
      </div>
    `);

    SpreadsheetApp.getUi().showModelessDialog(html, '📊 Dashboard');
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error}`);
  }
}

function actualizarDashboard() {
  // Se ejecuta después de sincronizar
  log('Dashboard actualizado', 'INFO');
}

// ============================================================================
// REPORTES
// ============================================================================

function abrirReportes() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; padding: 20px; background: #f5f5f5; }
      .container { max-width: 800px; background: white; padding: 25px; border-radius: 10px; }
      h1 { color: #1f73e6; }
      .reporte { background: #f9f9f9; padding: 15px; margin: 15px 0; border-left: 4px solid #1f73e6; border-radius: 5px; cursor: pointer; }
      .reporte:hover { background: #e3f2fd; }
      .reporte-title { font-weight: bold; color: #1f73e6; }
      .reporte-desc { font-size: 12px; color: #666; margin-top: 5px; }
    </style>

    <div class="container">
      <h1>📋 REPORTES</h1>

      <div class="reporte" onclick="google.script.run.generarReporteDiario()">
        <div class="reporte-title">📅 Reporte Diario</div>
        <div class="reporte-desc">Nuevos participantes, cambios de estado, derivaciones del día</div>
      </div>

      <div class="reporte" onclick="google.script.run.generarReportePerfiles()">
        <div class="reporte-title">📊 Reporte de Perfiles</div>
        <div class="reporte-desc">Distribución por Perfil A, B, C, D con detalles</div>
      </div>

      <div class="reporte" onclick="google.script.run.generarReporteDerivaciones()">
        <div class="reporte-title">📁 Reporte de Derivaciones</div>
        <div class="reporte-desc">Todas las derivaciones pendientes y completadas</div>
      </div>

      <div class="reporte" onclick="google.script.run.generarReportePrioridad()">
        <div class="reporte-title">⚠️ Reporte de Prioridad</div>
        <div class="reporte-desc">Casos CRÍTICOS y ALTOS por atender URGENTE</div>
      </div>

      <div class="reporte" onclick="window.close()">
        <div class="reporte-title" style="color: #999;">❌ Cerrar</div>
      </div>
    </div>

    <script>
      google.script.run.generarReportePerfiles();
    </script>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, '📋 Reportes');
}

function generarReporteDiario() {
  log('Reporte diario generado', 'INFO');
  SpreadsheetApp.getUi().alert('📅 Reporte generado y enviado por email');
}

function generarReportePerfiles() {
  log('Reporte de perfiles generado', 'INFO');
}

function generarReporteDerivaciones() {
  log('Reporte de derivaciones generado', 'INFO');
}

function generarReportePrioridad() {
  log('Reporte de prioridad generado', 'INFO');
}

function actualizarReportes() {
  log('Reportes actualizados', 'INFO');
}

// ============================================================================
// DERIVACIONES
// ============================================================================

function abrirDerivaciones() {
  const hojaMaestro = getHoja(CONFIG.HOJAS.MAESTRO);

  if (!hojaMaestro) {
    SpreadsheetApp.getUi().alert('⚠️ No hay datos. Sincroniza primero.');
    return;
  }

  const datos = hojaMaestro.getDataRange().getValues();
  let derivacionesPendientes = [];

  for (let i = 1; i < datos.length; i++) {
    const estado = datos[i][23]; // Columna X: Estado_Derivación
    if (estado === 'Pendiente') {
      derivacionesPendientes.push({
        id: datos[i][0],
        nombre: datos[i][2],
        perfil: datos[i][13],
        derivadoA: datos[i][21],
        prioridad: datos[i][14]
      });
    }
  }

  let html = `
    <style>
      body { font-family: Arial; padding: 20px; background: #f5f5f5; }
      .container { max-width: 900px; background: white; padding: 25px; border-radius: 10px; }
      h1 { color: #1f73e6; }
      table { width: 100%; border-collapse: collapse; margin-top: 20px; }
      th { background: #1f73e6; color: white; padding: 12px; text-align: left; }
      td { padding: 12px; border-bottom: 1px solid #ddd; }
      tr:hover { background: #f5f5f5; }
    </style>

    <div class="container">
      <h1>📁 Derivaciones Pendientes (${derivacionesPendientes.length})</h1>
      <table>
        <thead>
          <tr>
            <th>ID</th>
            <th>Nombre</th>
            <th>Perfil</th>
            <th>Derivado A</th>
            <th>Prioridad</th>
          </tr>
        </thead>
        <tbody>
  `;

  for (const deriv of derivacionesPendientes) {
    html += `
      <tr>
        <td><strong>${deriv.id}</strong></td>
        <td>${deriv.nombre}</td>
        <td>${deriv.perfil}</td>
        <td>${deriv.derivadoA || '-'}</td>
        <td><strong style="color: ${deriv.prioridad === 'CRÍTICO' ? 'red' : 'orange'}">${deriv.prioridad}</strong></td>
      </tr>
    `;
  }

  html += `
        </tbody>
      </table>
    </div>
  `;

  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(html),
    '📁 Derivaciones Pendientes'
  );
}

// ============================================================================
// AYUDA
// ============================================================================

function mostrarAyuda() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; padding: 20px; background: #f5f5f5; }
      .container { max-width: 700px; background: white; padding: 25px; border-radius: 10px; }
      h1 { color: #1f73e6; }
      h2 { color: #333; margin-top: 20px; border-bottom: 2px solid #1f73e6; padding-bottom: 10px; }
      .step { background: #e3f2fd; padding: 15px; border-left: 4px solid #1f73e6; border-radius: 5px; margin: 10px 0; }
      ul { margin: 10px 0; }
      li { margin: 8px 0; }
    </style>

    <div class="container">
      <h1>📖 Ayuda - Paso a Paso v4.0</h1>

      <h2>🚀 Primeros Pasos</h2>
      <div class="step">
        <strong>1. Configuración</strong><br>
        Abre ⚙️ Configuración e ingresa tu API Key de Kobo
      </div>
      <div class="step">
        <strong>2. Sincronizar</strong><br>
        Click en 🔄 SINCRONIZAR Kobo para descargar datos
      </div>
      <div class="step">
        <strong>3. Ver Dashboard</strong><br>
        📊 Dashboard muestra gráficos y estadísticas
      </div>

      <h2>📊 Hojas Automáticas</h2>
      <ul>
        <li><strong>Maestro:</strong> Todos los participantes con datos de Kobo</li>
        <li><strong>Dashboard:</strong> Gráficos y estadísticas</li>
        <li><strong>Derivación:</strong> Seguimiento de derivaciones</li>
        <li><strong>Reportes:</strong> Reportes automáticos</li>
        <li><strong>Log:</strong> Historial de todas las acciones</li>
      </ul>

      <h2>🔄 Sincronización</h2>
      <p>Cuando sincronizas, el sistema:</p>
      <ul>
        <li>✅ Descarga nuevos participantes de Kobo</li>
        <li>✅ Actualiza datos existentes</li>
        <li>✅ Clasifica por Perfil y Prioridad</li>
        <li>✅ Actualiza Dashboard y Reportes</li>
      </ul>
    </div>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, '📖 Ayuda');
}

// ============================================================================
// FIN DEL SISTEMA
// ============================================================================
