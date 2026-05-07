// ============================================================================
// SISTEMA PASO A PASO - FITO (VERSIÓN COMPLETA CON MENÚ)
// ============================================================================
// Gestión integral de participantes con automatización en Google Sheets
// Todas las funciones: Importación, Documentos, Alertas, Estados
// ============================================================================

// ============================================================================
// MENÚ PRINCIPAL - Se ejecuta al abrir la hoja
// ============================================================================

function onOpen() {
  cargarConfiguracion();
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 PASO A PASO')
    .addItem('⚙️ Configuración Inicial', 'mostrarConfiguracion')
    .addSeparator()
    .addItem('📥 Importar de KoboToolbox', 'importarDatosKobo')
    .addItem('🔔 Revisar Alertas de Mentoría', 'revisarAlertasMentoria')
    .addItem('📈 Ver Estadísticas', 'mostrarEstadisticas')
    .addSeparator()
    .addItem('📖 Ayuda y Documentación', 'mostrarAyuda')
    .addToUi();
}

// ============================================================================
// CONFIGURACIÓN CENTRALIZADA DEL SISTEMA PASO A PASO
// ============================================================================

const CONFIG = {
  // IDs de Google Sheets y Drives (AUTOMÁTICO)
  SPREADSHEET_ID: SpreadsheetApp.getActive().getId(),
  SHEET_NAME: "Maestro",
  FOLDER_PARTICIPANTES_ID: "",

  // URLs y Credenciales Kobo
  KOBO_CSV_URL: "https://kf.kobotoolbox.org/api/v2/assets/abHRWdRnPhKwzPQBajc7RZ/export-settings/esxXsxLthEAgNHbXQnV8qYV/data.csv",
  KOBO_USERNAME: "",
  KOBO_PASSWORD: "",

  // Emails para notificaciones
  ADMIN_EMAIL: "adrian@creamosguatemla.org",
  EMAILS_ALERTAS: {
    "empleabilidad": "empleabilidad@creamosguatemla.org",
    "tech": "tech@creamosguatemla.org",
    "ayb": "ayb@creamosguatemla.org"
  },

  // Configuración de columnas (A=1, B=2, etc.)
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
    FUENTE: 24,
    PROGRAMA: 25,
    TIPO_DERIVACION: 26,
    DERIVADO_A: 27,
    TIPO_CIERRE: 28,
    FECHA_CIERRE: 29,
    RESULTADO_FINAL: 30,
    ENCARGADO_CIERRE: 31
  },

  // Valores permitidos
  ENUM: {
    NIVEL_EDUCATIVO: ["Sin educación", "Primaria", "Secundaria", "Técnico", "Universitario"],
    CONCIENCIA: ["Alto", "Medio", "Bajo"],
    ESTADO: ["Orientación", "Mentoría", "Formación", "Cierre", "Derivación", "Activo", "Inactivo"],
    TIPO_CIERRE: ["Completado", "No completado", "Derivado", "Abandonado"],
    TIPO_DERIVACION: ["Interna", "Externa", "Ninguna"],
    GENERO: ["Masculino", "Femenino", "Otro"],
    ZONAS: ["Zona 1", "Zona 2", "Zona 3", "Zona 4", "Zona 5", "Zona 6", "Zona 7", "Otro"]
  },

  // Configuración de alertas
  ALERTAS: {
    DURACION_MENTORIA_MESES_MAX: 6,
    SESIONES_MENTORIA_MAX: 3,
    HORA_ALERTAS: 9,
    HORA_IMPORTAR_KOBO: 6
  }
};

// ============================================================================
// CARGAR CONFIGURACIÓN GUARDADA
// ============================================================================

function cargarConfiguracion() {
  const props = PropertiesService.getUserProperties();

  const carpetaId = props.getProperty('FOLDER_PARTICIPANTES_ID');
  const koboUser = props.getProperty('KOBO_USERNAME');
  const koboPass = props.getProperty('KOBO_PASSWORD');
  const koboUrl = props.getProperty('KOBO_CSV_URL');

  if (carpetaId) CONFIG.FOLDER_PARTICIPANTES_ID = carpetaId;
  if (koboUser) CONFIG.KOBO_USERNAME = koboUser;
  if (koboPass) CONFIG.KOBO_PASSWORD = koboPass;
  if (koboUrl) CONFIG.KOBO_CSV_URL = koboUrl;

  log('📂 Configuración cargada desde almacenamiento', 'INFO');
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

function log(mensaje, tipo = "INFO") {
  const timestamp = new Date().toLocaleString("es-GT");
  const linea = `[${timestamp}] [${tipo}] ${mensaje}`;
  Logger.log(linea);
}

// ============================================================================
// MENÚ - CONFIGURACIÓN
// ============================================================================

function mostrarConfiguracion() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
      .container { max-width: 500px; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
      h2 { color: #1f73e6; margin-top: 0; }
      .section { margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px; }
      label { display: block; margin: 10px 0 5px 0; font-weight: bold; color: #333; }
      input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; box-sizing: border-box; }
      button { background: #1f73e6; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin-top: 10px; }
      button:hover { background: #1557b0; }
      .info { background: #e3f2fd; padding: 10px; border-left: 4px solid #1f73e6; margin: 10px 0; border-radius: 4px; }
      .success { background: #c8e6c9; padding: 10px; color: #2e7d32; border-radius: 4px; }
    </style>

    <div class="container">
      <h2>⚙️ Configuración del Sistema</h2>

      <div class="section">
        <h3>1️⃣ ID del Google Sheet</h3>
        <div class="success">✅ AUTOMÁTICO: ${CONFIG.SPREADSHEET_ID}</div>
      </div>

      <div class="section">
        <h3>2️⃣ Carpeta en Google Drive</h3>
        <label>ID de la Carpeta "Participantes":</label>
        <input type="text" id="carpetaId" placeholder="Ej: 1A2B3C4D5E6F7G8H9I0J">
        <div class="info">📌 Copia el ID de tu carpeta desde la URL: drive.google.com/drive/folders/<strong>AQUÍ_VA_EL_ID</strong></div>
      </div>

      <div class="section">
        <h3>3️⃣ KoboToolbox</h3>
        <label>Usuario de KoboToolbox:</label>
        <input type="text" id="koboUser" placeholder="usuario@email.com">

        <label>Contraseña de KoboToolbox:</label>
        <input type="password" id="koboPass" placeholder="tu_contraseña">

        <label>URL CSV de KoboToolbox (Opcional):</label>
        <input type="text" id="koboUrl" placeholder="https://kf.kobotoolbox.org/api/v2/assets/...">
        <div class="info">📌 Si tienes una URL específica de tu formulario Kobo, pégala aquí. Si no, usa la configuración por defecto.</div>
      </div>

      <button onclick="guardarConfiguracion()">💾 GUARDAR CONFIGURACIÓN</button>
      <button onclick="probarConexion()">🧪 PROBAR CONEXIÓN</button>
    </div>

    <script>
      function guardarConfiguracion() {
        const carpeta = document.getElementById('carpetaId').value;
        const user = document.getElementById('koboUser').value;
        const pass = document.getElementById('koboPass').value;
        const url = document.getElementById('koboUrl').value;

        if (!carpeta) {
          alert('❌ Debes ingresar el ID de la carpeta');
          return;
        }

        if (!user || !pass) {
          alert('⚠️ Se recomienda ingresar usuario y contraseña de Kobo');
        }

        google.script.run.guardarConfiguracionScript(carpeta, user, pass, url);
        alert('✅ Configuración guardada exitosamente');
      }

      function probarConexion() {
        const user = document.getElementById('koboUser').value;
        const pass = document.getElementById('koboPass').value;

        if (!user || !pass) {
          alert('❌ Debes ingresar usuario y contraseña de Kobo para probar');
          return;
        }

        alert('🧪 Probando conexión...');
        google.script.run.probarConexionKobo();
      }
    </script>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, '⚙️ Configuración');
}

function guardarConfiguracionScript(carpetaId, koboUser, koboPass, koboUrl) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('FOLDER_PARTICIPANTES_ID', carpetaId);
  props.setProperty('KOBO_USERNAME', koboUser);
  props.setProperty('KOBO_PASSWORD', koboPass);
  if (koboUrl) props.setProperty('KOBO_CSV_URL', koboUrl);

  CONFIG.FOLDER_PARTICIPANTES_ID = carpetaId;
  CONFIG.KOBO_USERNAME = koboUser;
  CONFIG.KOBO_PASSWORD = koboPass;
  if (koboUrl) CONFIG.KOBO_CSV_URL = koboUrl;

  log('✅ Configuración guardada', 'INFO');
}

function probarConexionKobo() {
  try {
    if (!CONFIG.KOBO_USERNAME || !CONFIG.KOBO_PASSWORD) {
      SpreadsheetApp.getUi().alert('❌ Credenciales de Kobo no configuradas');
      return;
    }

    const auth = Utilities.base64Encode(`${CONFIG.KOBO_USERNAME}:${CONFIG.KOBO_PASSWORD}`);
    const options = {
      headers: { "Authorization": `Basic ${auth}` },
      muteHttpExceptions: true,
      timeout: 30
    };

    log(`Probando conexión Kobo con URL: ${CONFIG.KOBO_CSV_URL}`, 'INFO');
    const response = UrlFetchApp.fetch(CONFIG.KOBO_CSV_URL, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 200) {
      const content = response.getContentText();
      const lineas = content.split('\n').length;
      SpreadsheetApp.getUi().alert(`✅ Conexión EXITOSA\n\n📊 ${lineas - 1} registros encontrados`);
      log(`Conexión exitosa. Registros: ${lineas - 1}`, 'INFO');
    } else if (statusCode === 401) {
      SpreadsheetApp.getUi().alert(`❌ Error 401: Credenciales inválidas\n\nVerifica tu usuario y contraseña de KoboToolbox`);
      log(`Error 401: Credenciales inválidas. User: ${CONFIG.KOBO_USERNAME}`, 'ERROR');
    } else {
      const errorMsg = response.getContentText();
      SpreadsheetApp.getUi().alert(`❌ Error ${statusCode}\n\n${errorMsg.substring(0, 100)}`);
      log(`Error HTTP ${statusCode}: ${errorMsg}`, 'ERROR');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error de conexión: ${error}`);
    log(`Error en probarConexionKobo: ${error}`, 'ERROR');
  }
}

// ============================================================================
// MENÚ - ESTADÍSTICAS
// ============================================================================

function mostrarEstadisticas() {
  cargarConfiguracion();
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
    const values = sheet.getDataRange().getValues();

    let stats = {
      total: values.length - 1,
      orientacion: 0,
      mentoria: 0,
      formacion: 0,
      cierre: 0,
      alertas: 0
    };

    for (let i = 1; i < values.length; i++) {
      const estado = values[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1];
      const alerta = values[i][CONFIG.COLUMNAS.ALERTA_MENTORIA - 1];

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
        h2 { color: #1f73e6; }
        .stat { display: flex; justify-content: space-between; padding: 10px; border-bottom: 1px solid #eee; }
        .stat-label { font-weight: bold; }
        .stat-value { font-size: 18px; color: #1f73e6; font-weight: bold; }
        .alert { background: #ffebee; padding: 10px; color: #d32f2f; border-radius: 4px; margin: 10px 0; }
      </style>

      <div class="container">
        <h2>📊 Estadísticas del Sistema</h2>
        <div class="stat">
          <span class="stat-label">Total de Participantes:</span>
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
        ${stats.alertas > 0 ? `<div class="alert">⚠️ ${stats.alertas} participantes en ALERTA</div>` : ''}
      </div>
    `);

    SpreadsheetApp.getUi().showModelessDialog(html, '📊 Estadísticas');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error);
  }
}

// ============================================================================
// MENÚ - AYUDA
// ============================================================================

function mostrarAyuda() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; padding: 20px; background: #f5f5f5; }
      .container { max-width: 600px; background: white; padding: 20px; border-radius: 8px; }
      h2 { color: #1f73e6; }
      h3 { color: #333; margin-top: 15px; }
      ul { margin: 10px 0; }
      li { margin: 5px 0; }
      .note { background: #e3f2fd; padding: 10px; border-left: 4px solid #1f73e6; margin: 10px 0; }
    </style>

    <div class="container">
      <h2>📖 Ayuda - Sistema Paso a Paso</h2>

      <h3>🚀 Primeros Pasos</h3>
      <ol>
        <li>Abre el menú "📊 PASO A PASO"</li>
        <li>Haz click en "⚙️ Configuración Inicial"</li>
        <li>Completa los datos requeridos</li>
        <li>Haz click en "💾 GUARDAR CONFIGURACIÓN"</li>
      </ol>

      <h3>📥 Importar Datos</h3>
      <p>Usa el menú "📥 Importar de KoboToolbox" para descargar los datos automáticamente</p>

      <h3>🔔 Alertas de Mentoría</h3>
      <p>Se activan si:</p>
      <ul>
        <li>La mentoría dura más de 6 meses</li>
        <li>Se completaron más de 3 sesiones</li>
      </ul>

      <h3>📋 Cambios de Estado</h3>
      <p>Simplemente edita la columna "Estado_Actual" en el Sheet y el sistema se actualizará automáticamente</p>

      <div class="note">
        <strong>💡 Tip:</strong> Usa los filtros para ver solo participantes activos, en alerta, etc.
      </div>
    </div>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, '📖 Ayuda');
}

// ============================================================================
// IMPORTACIÓN DE KOBO (FUNCIONES)
// ============================================================================

function importarDatosKobo() {
  cargarConfiguracion();

  if (!CONFIG.KOBO_USERNAME || !CONFIG.KOBO_PASSWORD) {
    SpreadsheetApp.getUi().alert('❌ Debes configurar KoboToolbox primero\n\nAbre: 📊 PASO A PASO > ⚙️ Configuración Inicial');
    return;
  }

  if (!CONFIG.FOLDER_PARTICIPANTES_ID) {
    SpreadsheetApp.getUi().alert('❌ Debes configurar la carpeta de participantes\n\nAbre: 📊 PASO A PASO > ⚙️ Configuración Inicial');
    return;
  }

  log("=== INICIANDO IMPORTACIÓN KOBO ===", "INFO");

  try {
    const csvData = descargarCSVKobo();
    if (!csvData) {
      log("No se pudo descargar CSV de Kobo", "ERROR");
      SpreadsheetApp.getUi().alert('❌ No se pudo descargar datos de Kobo.\n\nVerifica:\n1. Usuario y contraseña\n2. URL del formulario\n3. Tu conexión a internet');
      return;
    }

    const registros = parsearCSV(csvData);
    if (registros.length === 0) {
      log("CSV vacío o sin registros válidos", "WARN");
      return;
    }

    log(`CSV descargado: ${registros.length} registros encontrados`, "INFO");

    const hoja = getHojaMatriz();
    const datosExistentes = hoja.getDataRange().getValues();
    const dpisExistentes = new Set(datosExistentes.slice(1).map(row => row[CONFIG.COLUMNAS.DPI - 1]));

    const registrosNuevos = registros.filter(reg => {
      const dpi = reg.DPI || "";
      return dpi && !dpisExistentes.has(dpi);
    });

    log(`Registros nuevos detectados: ${registrosNuevos.length}`, "INFO");

    if (registrosNuevos.length === 0) {
      SpreadsheetApp.getUi().alert('✅ No hay registros nuevos para importar');
      return;
    }

    let registrosProcesados = 0;
    for (const registro of registrosNuevos) {
      try {
        agregarNuevoParticipante(registro);
        registrosProcesados++;
      } catch (error) {
        log(`Error procesando participante ${registro.NOMBRE_COMPLETO}: ${error}`, "ERROR");
      }
    }

    log(`=== IMPORTACIÓN COMPLETADA: ${registrosProcesados}/${registrosNuevos.length} ===`, "INFO");
    SpreadsheetApp.getUi().alert(`✅ Importación completada\n${registrosProcesados} participantes agregados`);

  } catch (error) {
    log(`Error en importación general: ${error}`, "ERROR");
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

function descargarCSVKobo() {
  try {
    const url = CONFIG.KOBO_CSV_URL;
    const auth = Utilities.base64Encode(`${CONFIG.KOBO_USERNAME}:${CONFIG.KOBO_PASSWORD}`);

    const options = {
      headers: { "Authorization": `Basic ${auth}` },
      muteHttpExceptions: true,
      timeout: 60
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      log(`Error Kobo HTTP ${response.getResponseCode()}: ${response.getContentText()}`, "ERROR");
      return null;
    }

    return response.getContentText();

  } catch (error) {
    log(`Error descargando CSV Kobo: ${error}`, "ERROR");
    return null;
  }
}

function parsearCSV(csvText) {
  const lineas = csvText.split("\n").filter(l => l.trim());
  if (lineas.length < 2) return [];

  const headers = lineas[0].split(",").map(h => h.trim().toLowerCase());
  const registros = [];

  for (let i = 1; i < lineas.length; i++) {
    const valores = csvText.split("\n")[i].split(",");
    const registro = {};

    headers.forEach((header, index) => {
      registro[header.toUpperCase()] = (valores[index] || "").trim();
    });

    if (registro.DPI) {
      registros.push(registro);
    }
  }

  return registros;
}

function agregarNuevoParticipante(datosKobo) {
  const hoja = getHojaMatriz();
  const idCreamos = generarIdCreamos(datosKobo);

  const nuevaFila = [];
  for (let i = 1; i <= 31; i++) {
    nuevaFila.push("");
  }

  nuevaFila[CONFIG.COLUMNAS.ID_CREAMOS - 1] = idCreamos;
  nuevaFila[CONFIG.COLUMNAS.FECHA_REGISTRO - 1] = new Date();
  nuevaFila[CONFIG.COLUMNAS.NOMBRE_COMPLETO - 1] = datosKobo.NOMBRE_COMPLETO_SEGUN_DPI || datosKobo.NOMBRE_COMPLETO || "";
  nuevaFila[CONFIG.COLUMNAS.DPI - 1] = datosKobo.DPI || "";
  nuevaFila[CONFIG.COLUMNAS.EDAD - 1] = parseInt(datosKobo.EDAD) || "";
  nuevaFila[CONFIG.COLUMNAS.TELEFONO - 1] = datosKobo.TELEFONO || "";
  nuevaFila[CONFIG.COLUMNAS.ZONA - 1] = datosKobo.ZONA || "";
  nuevaFila[CONFIG.COLUMNAS.NIVEL_EDUCATIVO - 1] = datosKobo.NIVEL_EDUCATIVO || "";
  nuevaFila[CONFIG.COLUMNAS.PROBLEMAS_VITALES - 1] = datosKobo.EN_QUE_AREA_ESTA_INTERESADO || "";
  nuevaFila[CONFIG.COLUMNAS.FUENTE - 1] = "KoboToolbox";
  nuevaFila[CONFIG.COLUMNAS.PROGRAMA - 1] = datosKobo.PROGRAMA || "General";
  nuevaFila[CONFIG.COLUMNAS.ESTADO_ACTUAL - 1] = "Orientación";

  hoja.appendRow(nuevaFila);
  log(`Participante agregado: ${idCreamos} - ${datosKobo.NOMBRE_COMPLETO}`, "INFO");
}

function generarIdCreamos(datos) {
  const nombre = datos.NOMBRE_COMPLETO || datos.NOMBRE_COMPLETO_SEGUN_DPI || "X";
  const dpi = datos.DPI || "000000";
  const primerLetra = nombre.charAt(0).toUpperCase();
  const ultimosDpi = dpi.slice(-6).padStart(6, "0");
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "ddMM");

  return `${primerLetra}${ultimosDpi.substring(0, 3)}${fecha}`;
}

// ============================================================================
// ALERTAS DE MENTORÍA (FUNCIONES)
// ============================================================================

function revisarAlertasMentoria() {
  cargarConfiguracion();
  log("=== REVISIÓN DE ALERTAS DE MENTORÍA ===", "INFO");

  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    let alertasEncontradas = [];
    const ahora = new Date();

    for (let i = 1; i < valores.length; i++) {
      const fila = valores[i];
      const estado = fila[CONFIG.COLUMNAS.ESTADO_ACTUAL - 1];

      if (estado !== "Mentoría") continue;

      const idCreamos = fila[CONFIG.COLUMNAS.ID_CREAMOS - 1];
      const nombreCompleto = fila[CONFIG.COLUMNAS.NOMBRE_COMPLETO - 1];
      const inicioMentoria = new Date(fila[CONFIG.COLUMNAS.INICIO_MENTORIA - 1]);
      const sesionesCompletadas = parseInt(fila[CONFIG.COLUMNAS.SESIONES_COMPLETADAS - 1]) || 0;
      const responsable = fila[CONFIG.COLUMNAS.RESPONSABLE_ACTUAL - 1];

      const duracionMeses = calcularMeses(inicioMentoria, ahora);

      if (duracionMeses >= CONFIG.ALERTAS.DURACION_MENTORIA_MESES_MAX ||
          sesionesCompletadas >= CONFIG.ALERTAS.SESIONES_MENTORIA_MAX) {

        const motivo = [];
        if (duracionMeses >= CONFIG.ALERTAS.DURACION_MENTORIA_MESES_MAX) {
          motivo.push(`Duración: ${duracionMeses} meses`);
        }
        if (sesionesCompletadas >= CONFIG.ALERTAS.SESIONES_MENTORIA_MAX) {
          motivo.push(`Sesiones: ${sesionesCompletadas}`);
        }

        alertasEncontradas.push({
          fila: i + 1,
          idCreamos,
          nombreCompleto,
          responsable,
          duracionMeses,
          sesionesCompletadas,
          motivo: motivo.join(" + ")
        });

        hoja.getRange(i + 1, CONFIG.COLUMNAS.ALERTA_MENTORIA)
          .setValue("🔴 ALERTA - Revisar")
          .setBackground("#ff6666")
          .setFontColor("#ffffff");
      }
    }

    log(`Alertas detectadas: ${alertasEncontradas.length}`, "INFO");
    SpreadsheetApp.getUi().alert(`✅ Revisión completada\n${alertasEncontradas.length} alertas encontradas`);

  } catch (error) {
    log(`Error en revisión de alertas: ${error}`, "ERROR");
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

function calcularMeses(fechaInicio, fechaFin) {
  const inicioDate = new Date(fechaInicio);
  const finDate = new Date(fechaFin);

  const meses = (finDate.getFullYear() - inicioDate.getFullYear()) * 12 +
    (finDate.getMonth() - inicioDate.getMonth());

  return Math.floor(meses);
}

// ============================================================================
// FIN DEL SISTEMA PASO A PASO
// ============================================================================
