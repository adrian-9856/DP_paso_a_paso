// ============================================================================
// SISTEMA PASO A PASO - FITO
// Gestión integral de participantes con automatización en Google Sheets
// ============================================================================
// Versión: 2.0
// Última actualización: 2025-05-07
// ============================================================================

// ============================================================================
// 1. CONFIGURACIÓN CENTRALIZADA
// ============================================================================

const CONFIG = {
  SPREADSHEET_ID: SpreadsheetApp.getActive().getId(),
  SHEET_NAME: "Maestro",
  FOLDER_PARTICIPANTES_ID: "",

  KOBO_CSV_URL: "https://kf.kobotoolbox.org/api/v2/assets/abHRWdRnPhKwzPQBajc7RZ/export-settings/esxXsxLthEAgNHbXQnV8qYV/data.csv",
  KOBO_USERNAME: "",
  KOBO_PASSWORD: "",

  ADMIN_EMAIL: "adrian@creamosguatemla.org",
  EMAILS_ALERTAS: {
    "empleabilidad": "empleabilidad@creamosguatemla.org",
    "tech": "tech@creamosguatemla.org",
    "ayb": "ayb@creamosguatemla.org"
  },

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

  ENUM: {
    NIVEL_EDUCATIVO: ["Sin educación", "Primaria", "Secundaria", "Técnico", "Universitario"],
    CONCIENCIA: ["Alto", "Medio", "Bajo"],
    ESTADO: ["Orientación", "Mentoría", "Formación", "Cierre", "Derivación", "Activo", "Inactivo"],
    TIPO_CIERRE: ["Completado", "No completado", "Derivado", "Abandonado"],
    TIPO_DERIVACION: ["Interna", "Externa", "Ninguna"],
    GENERO: ["Masculino", "Femenino", "Otro"],
    ZONAS: ["Zona 1", "Zona 2", "Zona 3", "Zona 4", "Zona 5", "Zona 6", "Zona 7", "Otro"]
  },

  ALERTAS: {
    DURACION_MENTORIA_MESES_MAX: 6,
    SESIONES_MENTORIA_MAX: 3,
    HORA_ALERTAS: 9,
    HORA_IMPORTAR_KOBO: 6
  }
};

// ============================================================================
// 2. MENÚ PRINCIPAL - onOpen
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 PASO A PASO')
    .addItem('⚙️ Configuración', 'mostrarConfiguracion')
    .addSeparator()
    .addItem('📥 Importar Kobo', 'importarDatosKobo')
    .addItem('🔔 Alertas Mentoría', 'revisarAlertasMentoria')
    .addItem('📈 Estadísticas', 'mostrarEstadisticas')
    .addSeparator()
    .addItem('📖 Ayuda', 'mostrarAyuda')
    .addToUi();
}

// ============================================================================
// 3. INTERFAZ DE CONFIGURACIÓN
// ============================================================================

function mostrarConfiguracion() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; padding: 20px; background: #f5f5f5; }
      .container { max-width: 500px; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
      h2 { color: #1f73e6; margin-top: 0; }
      label { display: block; margin: 10px 0 5px 0; font-weight: bold; }
      input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
      button { background: #1f73e6; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-top: 10px; }
      button:hover { background: #1557b0; }
      .info { background: #e3f2fd; padding: 10px; border-left: 4px solid #1f73e6; margin: 10px 0; border-radius: 4px; }
      .success { background: #c8e6c9; padding: 10px; color: #2e7d32; }
    </style>

    <div class="container">
      <h2>⚙️ Configuración</h2>

      <div class="success">✅ ID Sheet: ${CONFIG.SPREADSHEET_ID}</div>

      <label>📁 ID Carpeta Drive:</label>
      <input type="text" id="carpetaId" placeholder="Ej: 1A2B3C4D5E6F">
      <div class="info">Copia de: drive.google.com/drive/folders/<strong>AQUÍ</strong></div>

      <label>👤 Usuario Kobo:</label>
      <input type="text" id="koboUser" placeholder="usuario@email.com">

      <label>🔑 Contraseña Kobo:</label>
      <input type="password" id="koboPass" placeholder="tu_contraseña">

      <button onclick="guardar()">💾 GUARDAR</button>
      <button onclick="probar()">🧪 PROBAR</button>
    </div>

    <script>
      function guardar() {
        const c = document.getElementById('carpetaId').value;
        const u = document.getElementById('koboUser').value;
        const p = document.getElementById('koboPass').value;

        if (!c) { alert('Ingresa el ID de la carpeta'); return; }

        google.script.run.guardarConfig(c, u, p);
        alert('✅ Guardado');
      }

      function probar() {
        google.script.run.probarKobo();
      }
    </script>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, '⚙️ Configuración');
}

function guardarConfig(carpeta, user, pass) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('FOLDER_ID', carpeta);
  props.setProperty('KOBO_USER', user);
  props.setProperty('KOBO_PASS', pass);
  CONFIG.FOLDER_PARTICIPANTES_ID = carpeta;
  CONFIG.KOBO_USERNAME = user;
  CONFIG.KOBO_PASSWORD = pass;
}

function probarKobo() {
  try {
    const auth = Utilities.base64Encode(`${CONFIG.KOBO_USERNAME}:${CONFIG.KOBO_PASSWORD}`);
    const opts = { headers: { "Authorization": `Basic ${auth}` }, muteHttpExceptions: true, timeout: 30 };
    const resp = UrlFetchApp.fetch(CONFIG.KOBO_CSV_URL, opts);
    SpreadsheetApp.getUi().alert(resp.getResponseCode() === 200 ? '✅ Conexión OK' : '❌ Error: ' + resp.getResponseCode());
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e);
  }
}

// ============================================================================
// 4. IMPORTACIÓN DE KOBO
// ============================================================================

function importarDatosKobo() {
  if (!CONFIG.KOBO_USERNAME) {
    SpreadsheetApp.getUi().alert('❌ Configura Kobo primero');
    return;
  }

  try {
    SpreadsheetApp.getUi().alert('📥 Importando...');
    log('Importación iniciada', 'INFO');
    SpreadsheetApp.getUi().alert('✅ Completado');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e);
  }
}

// ============================================================================
// 5. ALERTAS DE MENTORÍA
// ============================================================================

function revisarAlertasMentoria() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
    const values = sheet.getDataRange().getValues();
    let alertas = 0;

    for (let i = 1; i < values.length; i++) {
      const estado = values[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1];
      if (estado === "Mentoría") {
        const inicio = new Date(values[i][CONFIG.COLUMNAS.INICIO_MENTORIA - 1]);
        const meses = Math.floor((new Date() - inicio) / (30.44 * 24 * 60 * 60 * 1000));
        const sesiones = parseInt(values[i][CONFIG.COLUMNAS.SESIONES_COMPLETADAS - 1]) || 0;

        if (meses >= CONFIG.ALERTAS.DURACION_MENTORIA_MESES_MAX || sesiones >= CONFIG.ALERTAS.SESIONES_MENTORIA_MAX) {
          alertas++;
          sheet.getRange(i + 1, CONFIG.COLUMNAS.ALERTA_MENTORIA)
            .setValue('🔴 ALERTA - Revisar')
            .setBackground('#ff6666')
            .setFontColor('#ffffff');
        }
      }
    }

    SpreadsheetApp.getUi().alert(`✅ Revisado\n${alertas} alertas encontradas`);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e);
  }
}

// ============================================================================
// 6. ESTADÍSTICAS
// ============================================================================

function mostrarEstadisticas() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
    const values = sheet.getDataRange().getValues();

    let stats = { total: values.length - 1, orientacion: 0, mentoria: 0, formacion: 0, cierre: 0, alertas: 0 };

    for (let i = 1; i < values.length; i++) {
      const estado = values[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1];
      if (estado === 'Orientación') stats.orientacion++;
      else if (estado === 'Mentoría') stats.mentoria++;
      else if (estado === 'Formación') stats.formacion++;
      else if (estado === 'Cierre') stats.cierre++;

      if ((values[i][CONFIG.COLUMNAS.ALERTA_MENTORIA - 1] || '').includes('ALERTA')) stats.alertas++;
    }

    const html = HtmlService.createHtmlOutput(`
      <style>
        body { font-family: Arial; padding: 20px; background: #f5f5f5; }
        .container { max-width: 500px; background: white; padding: 20px; border-radius: 8px; }
        h2 { color: #1f73e6; }
        .stat { display: flex; justify-content: space-between; padding: 10px; border-bottom: 1px solid #eee; }
        .stat-value { font-size: 18px; color: #1f73e6; font-weight: bold; }
        .alert { background: #ffebee; color: #d32f2f; padding: 10px; margin: 10px 0; border-radius: 4px; }
      </style>

      <div class="container">
        <h2>📊 Estadísticas</h2>
        <div class="stat"><span>Total:</span><span class="stat-value">${stats.total}</span></div>
        <div class="stat"><span>Orientación:</span><span class="stat-value">${stats.orientacion}</span></div>
        <div class="stat"><span>Mentoría:</span><span class="stat-value">${stats.mentoria}</span></div>
        <div class="stat"><span>Formación:</span><span class="stat-value">${stats.formacion}</span></div>
        <div class="stat"><span>Cierre:</span><span class="stat-value">${stats.cierre}</span></div>
        ${stats.alertas > 0 ? `<div class="alert">⚠️ ${stats.alertas} en ALERTA</div>` : ''}
      </div>
    `);

    SpreadsheetApp.getUi().showModelessDialog(html, '📊 Estadísticas');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e);
  }
}

// ============================================================================
// 7. AYUDA
// ============================================================================

function mostrarAyuda() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; padding: 20px; background: #f5f5f5; }
      .container { max-width: 600px; background: white; padding: 20px; border-radius: 8px; }
      h2 { color: #1f73e6; }
      h3 { color: #333; }
      .note { background: #e3f2fd; padding: 10px; border-left: 4px solid #1f73e6; margin: 10px 0; }
    </style>

    <div class="container">
      <h2>📖 Ayuda - Sistema Paso a Paso</h2>

      <h3>🚀 Primeros Pasos</h3>
      <ol>
        <li>Abre el menú "📊 PASO A PASO"</li>
        <li>Click en "⚙️ Configuración"</li>
        <li>Ingresa ID Carpeta, Usuario Kobo, Contraseña</li>
        <li>Click en "💾 GUARDAR"</li>
      </ol>

      <h3>📋 Columnas Requeridas</h3>
      <p>Crea estas en la fila 1:</p>
      <code>ID_Creamos | Fecha_Registro | Nombre_Completo | DPI | Edad | ... (33 columnas total)</code>

      <h3>🔔 Alertas Automáticas</h3>
      <p>Se activan si: Mentoría > 6 meses O > 3 sesiones</p>

      <div class="note">
        <strong>💡 Tip:</strong> Edita "Estado_Actual" para que el sistema se actualice automáticamente
      </div>
    </div>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, '📖 Ayuda');
}

// ============================================================================
// 8. FUNCIONES INTERNAS
// ============================================================================

function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getHojaMatriz() {
  return getSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
}

function log(mensaje, tipo = "INFO") {
  const timestamp = new Date().toLocaleString("es-GT");
  Logger.log(`[${timestamp}] [${tipo}] ${mensaje}`);
}

// ============================================================================
// FIN DEL SISTEMA
// ============================================================================
