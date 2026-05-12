// ============================================================================
// SISTEMA PASO A PASO - FITO (VERSIÓN 3.0 - COMPLETA Y AUTOMATIZADA)
// ============================================================================
// Gestión integral de participantes con integración KoboToolbox
// Automatización: Carpetas, Documentos, Fotos, Reportes
// ============================================================================

// ============================================================================
// MENÚ PRINCIPAL
// ============================================================================

function onOpen() {
  cargarConfiguracion();
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 PASO A PASO')
    .addItem('⚙️ Configuración', 'mostrarConfiguracion')
    .addSeparator()
    .addItem('📥 Sincronizar Kobo (Manual)', 'sincronizarKoboManual')
    .addItem('📊 Ver Estadísticas', 'mostrarEstadisticas')
    .addItem('🔔 Revisar Alertas', 'revisarAlertasMentoria')
    .addItem('📁 Crear Carpeta Participante', 'crearCarpetaParticipante')
    .addSeparator()
    .addItem('📖 Ayuda', 'mostrarAyuda')
    .addToUi();
}

// ============================================================================
// CONFIGURACIÓN DEL SISTEMA
// ============================================================================

const CONFIG = {
  // IDs Google (AUTOMÁTICO)
  SPREADSHEET_ID: SpreadsheetApp.getActive().getId(),
  SHEET_NAME: "Maestro",

  // Kobo - Usa API Key en lugar de usuario/contraseña
  KOBO_API_KEY: "",
  KOBO_ASSET_ID: "abHRWdRnPhKwzPQBajc7RZ",
  KOBO_EXPORT_ID: "esxXsxLthEAgNHbXQnV8qYV",
  KOBO_BASE_URL: "https://kf.kobotoolbox.org/api/v2",

  // Google Drive
  FOLDER_PARTICIPANTES_ID: "1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU",

  // Email
  ADMIN_EMAIL: "adrian@creamosguatemla.org",

  // Columnas del Sheet (A=1, B=2, etc.)
  COLUMNAS: {
    ID_CREAMOS: 1,
    FECHA_REGISTRO: 2,
    NOMBRE_COMPLETO: 3,
    DPI: 4,
    EDAD: 5,
    GENERO: 6,
    TELEFONO: 7,
    ZONA: 8,
    NIVEL_EDUCATIVO: 9,
    PROBLEMAS_VITALES: 10,
    CONCIENCIA_PARTICIPANTE: 11,
    PERFIL_RESULTANTE: 12,
    ESTADO_ACTUAL: 13,
    FECHA_ULTIMO_CAMBIO: 14,
    RESPONSABLE_ACTUAL: 15,
    INICIO_MENTORIA: 16,
    SESIONES_COMPLETADAS: 17,
    ULTIMA_SESION: 18,
    DURACION_MESES: 19,
    ALERTA_MENTORIA: 20,
    PROXIMA_SESION: 21,
    CARPETA_DRIVE_ID: 22,
    EXPEDIENTE_ID: 23,
    EXPEDIENTE_URL: 24,
    FOTO_URL: 25,
    FUENTE: 26,
    PROGRAMA: 27,
    TIPO_DERIVACION: 28,
    DERIVADO_A: 29,
    TIPO_CIERRE: 30,
    FECHA_CIERRE: 31,
    RESULTADO_FINAL: 32,
    ENCARGADO_CIERRE: 33
  },

  ENUM: {
    ESTADO: ["Orientación", "Mentoría", "Formación", "Cierre", "Derivación", "Activo", "Inactivo"],
    GENERO: ["Masculino", "Femenino", "Otro"],
    NIVEL_EDUCATIVO: ["Sin educación", "Primaria", "Secundaria", "Técnico", "Universitario"],
    PROGRAMA: ["Empleabilidad", "Tech", "AYB", "General"]
  },

  ALERTAS: {
    DURACION_MENTORIA_MESES_MAX: 6,
    SESIONES_MENTORIA_MAX: 3
  }
};

// ============================================================================
// CARGAR CONFIGURACIÓN
// ============================================================================

function cargarConfiguracion() {
  const props = PropertiesService.getUserProperties();

  const apiKey = props.getProperty('KOBO_API_KEY');
  const carpetaId = props.getProperty('FOLDER_PARTICIPANTES_ID');

  if (apiKey) CONFIG.KOBO_API_KEY = apiKey;
  if (carpetaId) CONFIG.FOLDER_PARTICIPANTES_ID = carpetaId;

  log('✅ Configuración cargada', 'INFO');
}

// ============================================================================
// FUNCIONES BASE
// ============================================================================

function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getHojaMatriz() {
  return getSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
}

function getCarpetaRaiz() {
  return DriveApp.getFolderById(CONFIG.FOLDER_PARTICIPANTES_ID);
}

function log(msg, tipo = "INFO") {
  const timestamp = new Date().toLocaleString("es-GT");
  Logger.log(`[${timestamp}] [${tipo}] ${msg}`);
}

// ============================================================================
// MENÚ - CONFIGURACIÓN
// ============================================================================

function mostrarConfiguracion() {
  const apiKey = PropertiesService.getUserProperties().getProperty('KOBO_API_KEY') || '';

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
      .container { max-width: 600px; background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
      h2 { color: #1f73e6; margin-top: 0; border-bottom: 2px solid #1f73e6; padding-bottom: 10px; }
      .section { margin-bottom: 20px; padding: 15px; background: #f9f9f9; border-radius: 5px; }
      .section h3 { color: #333; margin-top: 0; }
      label { display: block; margin: 10px 0 5px 0; font-weight: bold; color: #333; }
      input, textarea { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; box-sizing: border-box; font-family: monospace; }
      textarea { min-height: 80px; }
      button { background: #1f73e6; color: white; padding: 12px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin: 10px 5px 10px 0; }
      button:hover { background: #1557b0; }
      .info { background: #e3f2fd; padding: 12px; border-left: 4px solid #1f73e6; border-radius: 4px; margin: 10px 0; }
      .success { background: #c8e6c9; padding: 12px; color: #2e7d32; border-radius: 4px; border-left: 4px solid #2e7d32; }
      .error { background: #ffebee; padding: 12px; color: #c62828; border-radius: 4px; border-left: 4px solid #c62828; }
    </style>

    <div class="container">
      <h2>⚙️ Configuración del Sistema Paso a Paso</h2>

      <div class="section">
        <h3>1️⃣ ID de Google Drive</h3>
        <div class="success">✅ AUTOMÁTICO: ${CONFIG.SPREADSHEET_ID.substring(0, 20)}...</div>
        <div class="info">📌 Tu Sheet está detectado automáticamente</div>
      </div>

      <div class="section">
        <h3>2️⃣ Carpeta de Participantes</h3>
        <label>ID de la Carpeta (Google Drive):</label>
        <input type="text" id="carpetaId" value="1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU" placeholder="Pega el ID de tu carpeta">
        <div class="info">📌 URL: drive.google.com/drive/folders/<strong>AQUÍ_VA_EL_ID</strong></div>
      </div>

      <div class="section">
        <h3>3️⃣ Kobo API Key (IMPORTANTE)</h3>
        <label>Clave API de KoboToolbox:</label>
        <textarea id="apiKey" placeholder="Pega tu clave API de Kobo">${apiKey}</textarea>
        <div class="info">🔑 Obtén tu clave en: https://kf.kobotoolbox.org/admin/auth/token/</div>
      </div>

      <button onclick="guardarConfig()">💾 GUARDAR</button>
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
        google.script.run.probarConexionKobo(apiKey);
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
  CONFIG.KOBO_API_KEY = apiKey;

  log('✅ Configuración guardada', 'INFO');
}

function probarConexionKobo(apiKey) {
  try {
    const url = `${CONFIG.KOBO_BASE_URL}/assets/${CONFIG.KOBO_ASSET_ID}/`;
    const options = {
      headers: { "Authorization": `Token ${apiKey}` },
      muteHttpExceptions: true,
      timeout: 30
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const submissions = data.deployment__submission_count || 0;
      SpreadsheetApp.getUi().alert(`✅ Conexión EXITOSA\n\n📊 ${submissions} respuestas encontradas`);
      log(`Conexión exitosa. Submissions: ${submissions}`, 'INFO');
    } else {
      SpreadsheetApp.getUi().alert(`❌ Error ${response.getResponseCode()}\n\nVerifica tu API Key`);
      log(`Error HTTP ${response.getResponseCode()}`, 'ERROR');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
    log(`Error en prueba: ${error}`, 'ERROR');
  }
}

// ============================================================================
// MENÚ - SINCRONIZACIÓN KOBO
// ============================================================================

function sincronizarKoboManual() {
  cargarConfiguracion();

  if (!CONFIG.KOBO_API_KEY) {
    SpreadsheetApp.getUi().alert('❌ Configura la API Key primero\n\n📊 PASO A PASO > ⚙️ Configuración');
    return;
  }

  log("=== INICIANDO SINCRONIZACIÓN ===", "INFO");

  try {
    SpreadsheetApp.getUi().showModelessDialog(
      HtmlService.createHtmlOutput('<p>⏳ Sincronizando...</p>'),
      'Sincronización'
    );

    const registros = descargarRegistrosKobo();
    if (!registros || registros.length === 0) {
      SpreadsheetApp.getUi().alert('⚠️ No hay registros nuevos');
      return;
    }

    const hoja = getHojaMatriz();
    const datosExistentes = hoja.getDataRange().getValues();
    const dpisExistentes = new Set(datosExistentes.slice(1).map(row => row[CONFIG.COLUMNAS.DPI - 1]));

    let agregados = 0;
    for (const registro of registros) {
      const dpi = registro.DPI || '';
      if (dpi && !dpisExistentes.has(dpi)) {
        agregarParticipante(registro);
        agregados++;
      }
    }

    log(`✅ Sincronización completada: ${agregados} participantes agregados`, 'INFO');
    SpreadsheetApp.getUi().alert(`✅ Sincronización completada\n\n${agregados} participantes nuevos`);

  } catch (error) {
    log(`Error en sincronización: ${error}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

function descargarRegistrosKobo() {
  try {
    const url = `${CONFIG.KOBO_BASE_URL}/assets/${CONFIG.KOBO_ASSET_ID}/data/`;
    const options = {
      headers: { "Authorization": `Token ${CONFIG.KOBO_API_KEY}` },
      muteHttpExceptions: true,
      timeout: 60
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      log(`Error Kobo: ${response.getResponseCode()}`, 'ERROR');
      return null;
    }

    const data = JSON.parse(response.getContentText());
    const resultados = data.results || [];

    log(`Registros descargados: ${resultados.length}`, 'INFO');

    return resultados.map(r => ({
      DPI: r.DPI || '',
      NOMBRE_COMPLETO: r['nombre_completo_segun_dpi'] || r.nombre || '',
      EDAD: r.edad || '',
      GENERO: r.genero || '',
      TELEFONO: r.telefono || '',
      ZONA: r.zona || '',
      NIVEL_EDUCATIVO: r.nivel_educativo || '',
      PROBLEMAS_VITALES: r['en_que_area_esta_interesado'] || '',
      PROGRAMA: r.programa || 'General'
    }));

  } catch (error) {
    log(`Error descargando Kobo: ${error}`, 'ERROR');
    return null;
  }
}

function agregarParticipante(datosKobo) {
  try {
    const hoja = getHojaMatriz();
    const idCreamos = generarIdCreamos(datosKobo);

    // Crear carpeta del participante
    const carpetaParticipante = crearCarpetaParticipanteInterno(idCreamos, datosKobo);

    // Crear documento de perfil
    const docPerfil = crearDocumentoPerfilParticipante(idCreamos, datosKobo, carpetaParticipante);

    const nuevaFila = [
      idCreamos,                                    // A: ID_CREAMOS
      new Date(),                                   // B: FECHA_REGISTRO
      datosKobo.NOMBRE_COMPLETO,                  // C: NOMBRE_COMPLETO
      datosKobo.DPI,                              // D: DPI
      datosKobo.EDAD || '',                       // E: EDAD
      datosKobo.GENERO || '',                     // F: GENERO
      datosKobo.TELEFONO || '',                  // G: TELEFONO
      datosKobo.ZONA || '',                       // H: ZONA
      datosKobo.NIVEL_EDUCATIVO || '',           // I: NIVEL_EDUCATIVO
      datosKobo.PROBLEMAS_VITALES || '',         // J: PROBLEMAS_VITALES
      '',                                          // K: CONCIENCIA_PARTICIPANTE
      '',                                          // L: PERFIL_RESULTANTE
      'Orientación',                              // M: ESTADO_ACTUAL
      new Date(),                                 // N: FECHA_ULTIMO_CAMBIO
      '',                                         // O: RESPONSABLE_ACTUAL
      '',                                         // P: INICIO_MENTORIA
      0,                                          // Q: SESIONES_COMPLETADAS
      '',                                         // R: ULTIMA_SESION
      0,                                          // S: DURACION_MESES
      '',                                         // T: ALERTA_MENTORIA
      '',                                         // U: PROXIMA_SESION
      carpetaParticipante.getId(),               // V: CARPETA_DRIVE_ID
      docPerfil.getId(),                         // W: EXPEDIENTE_ID
      docPerfil.getUrl(),                        // X: EXPEDIENTE_URL
      '',                                         // Y: FOTO_URL
      'KoboToolbox',                             // Z: FUENTE
      datosKobo.PROGRAMA || 'General',           // AA: PROGRAMA
      '',                                         // AB: TIPO_DERIVACION
      '',                                         // AC: DERIVADO_A
      '',                                         // AD: TIPO_CIERRE
      '',                                         // AE: FECHA_CIERRE
      '',                                         // AF: RESULTADO_FINAL
      ''                                          // AG: ENCARGADO_CIERRE
    ];

    hoja.appendRow(nuevaFila);
    log(`✅ Participante agregado: ${idCreamos}`, 'INFO');

  } catch (error) {
    log(`Error agregando participante: ${error}`, 'ERROR');
  }
}

function crearCarpetaParticipanteInterno(idCreamos, datosKobo) {
  try {
    const carpetaRaiz = getCarpetaRaiz();
    const nombreCarpeta = `${idCreamos} - ${datosKobo.NOMBRE_COMPLETO}`;
    const carpeta = carpetaRaiz.createFolder(nombreCarpeta);

    log(`📁 Carpeta creada: ${nombreCarpeta}`, 'INFO');
    return carpeta;

  } catch (error) {
    log(`Error creando carpeta: ${error}`, 'ERROR');
    throw error;
  }
}

function crearDocumentoPerfilParticipante(idCreamos, datosKobo, carpetaParticipante) {
  try {
    const docName = `Perfil - ${datosKobo.NOMBRE_COMPLETO}`;

    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    // Encabezado
    body.appendParagraph(`PERFIL DEL PARTICIPANTE`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`Sistema: Paso a Paso - Fito`).setItalic(true);

    body.appendParagraph('');

    // Información personal
    body.appendParagraph('INFORMACIÓN PERSONAL').setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const table = body.appendTable([
      ['Campo', 'Valor'],
      ['ID Creamos', idCreamos],
      ['Nombre Completo', datosKobo.NOMBRE_COMPLETO],
      ['DPI', datosKobo.DPI],
      ['Edad', datosKobo.EDAD || '-'],
      ['Género', datosKobo.GENERO || '-'],
      ['Teléfono', datosKobo.TELEFONO || '-'],
      ['Zona', datosKobo.ZONA || '-'],
      ['Nivel Educativo', datosKobo.NIVEL_EDUCATIVO || '-'],
      ['Programa', datosKobo.PROGRAMA || 'General']
    ]);

    // Problemas vitales
    body.appendParagraph('');
    body.appendParagraph('ÁREAS DE INTERÉS / PROBLEMAS VITALES').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(datosKobo.PROBLEMAS_VITALES || '-');

    // Secciones para discusión de caso
    body.appendParagraph('');
    body.appendParagraph('DISCUSIÓN DE CASO').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Análisis y observaciones:');
    body.appendParagraph('');

    body.appendParagraph('Notas de sesiones:');
    body.appendParagraph('');

    body.appendParagraph('Plan de acción:');
    body.appendParagraph('');

    // Footer
    body.appendParagraph('');
    body.appendParagraph(`Creado: ${new Date().toLocaleString('es-GT')}`).setItalic(true);

    doc.saveAndClose();

    // Mover el documento a la carpeta del participante
    const file = DriveApp.getFileById(doc.getId());
    carpetaParticipante.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    log(`📄 Documento de perfil creado: ${docName}`, 'INFO');
    return doc;

  } catch (error) {
    log(`Error creando documento: ${error}`, 'ERROR');
    throw error;
  }
}

function crearCarpetaParticipante() {
  cargarConfiguracion();

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; padding: 20px; }
      input { width: 100%; padding: 10px; margin: 10px 0; border: 1px solid #ddd; border-radius: 4px; }
      button { background: #1f73e6; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
    </style>
    <h2>📁 Crear Carpeta para Participante</h2>
    <label>ID Creamos (ej: A1234):</label>
    <input type="text" id="idCreamos" placeholder="Ej: A1234">
    <label>Nombre del Participante:</label>
    <input type="text" id="nombre" placeholder="Nombre completo">
    <button onclick="crearCarpetaBtn()">Crear Carpeta</button>
    <script>
      function crearCarpetaBtn() {
        const id = document.getElementById('idCreamos').value;
        const nombre = document.getElementById('nombre').value;
        if (id && nombre) {
          google.script.run.crearCarpetaParticipanteManual(id, nombre);
        } else {
          alert('❌ Completa los campos');
        }
      }
    </script>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, 'Crear Carpeta');
}

function crearCarpetaParticipanteManual(idCreamos, nombre) {
  try {
    const carpetaRaiz = getCarpetaRaiz();
    const nombreCarpeta = `${idCreamos} - ${nombre}`;
    const carpeta = carpetaRaiz.createFolder(nombreCarpeta);

    SpreadsheetApp.getUi().alert(`✅ Carpeta creada: ${nombreCarpeta}`);
    log(`📁 Carpeta creada: ${nombreCarpeta}`, 'INFO');

  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

function generarIdCreamos(datos) {
  const nombre = datos.NOMBRE_COMPLETO || 'X';
  const dpi = datos.DPI || '000000';
  const primerLetra = nombre.charAt(0).toUpperCase();
  const ultimosDpi = dpi.slice(-6).padStart(6, '0').substring(0, 3);
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMM');

  return `${primerLetra}${ultimosDpi}${fecha}`;
}

// ============================================================================
// ALERTAS DE MENTORÍA
// ============================================================================

function revisarAlertasMentoria() {
  cargarConfiguracion();
  log("=== REVISIÓN DE ALERTAS ===", "INFO");

  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    let alertas = 0;
    const ahora = new Date();

    for (let i = 1; i < valores.length; i++) {
      const estado = valores[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1];

      if (estado !== "Mentoría") continue;

      const inicioMentoria = new Date(valores[i][CONFIG.COLUMNAS.INICIO_MENTORIA - 1]);
      const sesiones = parseInt(valores[i][CONFIG.COLUMNAS.SESIONES_COMPLETADAS - 1]) || 0;

      const meses = Math.floor((ahora - inicioMentoria) / (1000 * 60 * 60 * 24 * 30.44));

      if (meses >= CONFIG.ALERTAS.DURACION_MENTORIA_MESES_MAX || sesiones >= CONFIG.ALERTAS.SESIONES_MENTORIA_MAX) {
        hoja.getRange(i + 1, CONFIG.COLUMNAS.ALERTA_MENTORIA)
          .setValue("🔴 ALERTA")
          .setBackground("#ff6666")
          .setFontColor("#ffffff");
        alertas++;
      }
    }

    SpreadsheetApp.getUi().alert(`✅ Revisión completada\n${alertas} alertas encontradas`);
    log(`Alertas detectadas: ${alertas}`, 'INFO');

  } catch (error) {
    log(`Error en alertas: ${error}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

// ============================================================================
// ESTADÍSTICAS
// ============================================================================

function mostrarEstadisticas() {
  cargarConfiguracion();
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    let stats = {
      total: valores.length - 1,
      orientacion: 0,
      mentoria: 0,
      formacion: 0,
      cierre: 0,
      alertas: 0
    };

    for (let i = 1; i < valores.length; i++) {
      const estado = valores[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1];
      const alerta = valores[i][CONFIG.COLUMNAS.ALERTA_MENTORIA - 1];

      if (estado === 'Orientación') stats.orientacion++;
      else if (estado === 'Mentoría') stats.mentoria++;
      else if (estado === 'Formación') stats.formacion++;
      else if (estado === 'Cierre') stats.cierre++;

      if (alerta && alerta.includes('ALERTA')) stats.alertas++;
    }

    const html = HtmlService.createHtmlOutput(`
      <style>
        body { font-family: Arial; padding: 20px; background: #f5f5f5; }
        .container { max-width: 500px; background: white; padding: 20px; border-radius: 8px; }
        .stat { display: flex; justify-content: space-between; padding: 10px; border-bottom: 1px solid #eee; }
        .stat-label { font-weight: bold; }
        .stat-value { color: #1f73e6; font-weight: bold; font-size: 18px; }
        .alert { background: #ffebee; padding: 10px; color: #c62828; border-radius: 4px; margin: 10px 0; }
      </style>
      <div class="container">
        <h2>📊 Estadísticas</h2>
        <div class="stat">
          <span class="stat-label">Total Participantes:</span>
          <span class="stat-value">${stats.total}</span>
        </div>
        <div class="stat">
          <span class="stat-label">En Orientación:</span>
          <span class="stat-value">${stats.orientacion}</span>
        </div>
        <div class="stat">
          <span class="stat-label">En Mentoría:</span>
          <span class="stat-value">${stats.mentoria}</span>
        </div>
        <div class="stat">
          <span class="stat-label">En Formación:</span>
          <span class="stat-value">${stats.formacion}</span>
        </div>
        <div class="stat">
          <span class="stat-label">Cerrados:</span>
          <span class="stat-value">${stats.cierre}</span>
        </div>
        ${stats.alertas > 0 ? `<div class="alert">⚠️ ${stats.alertas} en ALERTA</div>` : ''}
      </div>
    `);

    SpreadsheetApp.getUi().showModelessDialog(html, '📊 Estadísticas');
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error}`);
  }
}

// ============================================================================
// AYUDA
// ============================================================================

function mostrarAyuda() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; padding: 20px; background: #f5f5f5; }
      .container { max-width: 600px; background: white; padding: 20px; border-radius: 8px; }
      h2 { color: #1f73e6; }
      h3 { color: #333; margin-top: 20px; }
      .step { background: #e3f2fd; padding: 15px; border-left: 4px solid #1f73e6; border-radius: 4px; margin: 10px 0; }
    </style>
    <div class="container">
      <h2>📖 Ayuda - Paso a Paso</h2>

      <h3>🚀 Primeros Pasos</h3>
      <div class="step">
        <strong>1. Configurar</strong><br>
        📊 PASO A PASO > ⚙️ Configuración
      </div>
      <div class="step">
        <strong>2. Sincronizar</strong><br>
        📥 Sincronizar Kobo (Manual)
      </div>
      <div class="step">
        <strong>3. Ver resultados</strong><br>
        Los participantes aparecerán en el Sheet y se crearán sus carpetas automáticamente
      </div>

      <h3>📁 Carpetas y Documentos</h3>
      <p>Cada participante tendrá automáticamente:</p>
      <ul>
        <li>📁 Carpeta con su ID + Nombre</li>
        <li>📄 Documento de Perfil (para discusión de caso)</li>
        <li>📸 Espacio para agregar foto</li>
      </ul>

      <h3>🔔 Alertas</h3>
      <p>Se activan cuando:</p>
      <ul>
        <li>Mentoría > 6 meses</li>
        <li>Sesiones > 3</li>
      </ul>
    </div>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, '📖 Ayuda');
}

// ============================================================================
// FIN DEL SISTEMA
// ============================================================================
