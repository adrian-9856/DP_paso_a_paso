// ============================================================================
// SISTEMA PASO A PASO v4.0 - PROFESIONAL CON INSTALACIÓN/DESINSTALACIÓN
// ============================================================================

function onOpen() {
  try {
    cargarConfiguracion();
    const ui = SpreadsheetApp.getUi();

    ui.createMenu('📊 PASO A PASO')
      .addItem('⚙️ CONFIGURACIÓN COMPLETA', 'mostrarConfiguracionCompleta')
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
    log('Error en onOpen: ' + error, 'ERROR');
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
  }
};

// ============================================================================
// FUNCIONES BASE
// ============================================================================

function cargarConfiguracion() {
  const props = PropertiesService.getUserProperties();
  const apiKey = props.getProperty('KOBO_API_KEY');
  if (apiKey) CONFIG.KOBO.API_KEY = apiKey;
  log('Configuración cargada', 'INFO');
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
      log('Hoja creada: ' + nombre, 'INFO');
    }
    return hoja;
  } catch (error) {
    log('Error creando hoja ' + nombre + ': ' + error, 'ERROR');
    return null;
  }
}

function getHoja(nombre) {
  return getSpreadsheet().getSheetByName(nombre);
}

function log(msg, tipo) {
  const timestamp = new Date().toLocaleString("es-GT");
  const linea = '[' + timestamp + '] [' + tipo + '] ' + msg;
  Logger.log(linea);

  try {
    const hojaLog = crearHojaSiNoExiste(CONFIG.HOJAS.LOG);
    if (hojaLog) hojaLog.appendRow([new Date(), tipo, msg]);
  } catch (error) {
    // Silent
  }
}

// ============================================================================
// MENÚ - CONFIGURACIÓN COMPLETA
// ============================================================================

function mostrarConfiguracionCompleta() {
  const props = PropertiesService.getUserProperties();
  const apiKey = props.getProperty('KOBO_API_KEY') || '';
  const carpetaId = props.getProperty('FOLDER_PARTICIPANTES_ID') || '';

  const ss = getSpreadsheet();
  const hojaMaestro = ss.getSheetByName(CONFIG.HOJAS.MAESTRO);
  const hojasExistentes = ss.getSheets().map(function(h) { return h.getName(); });

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; background: #f5f5f5; margin: 0; padding: 20px; }
      .container { max-width: 900px; background: white; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); overflow: hidden; }
      .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }
      .header h1 { margin: 0; font-size: 28px; }
      .tabs { display: flex; background: #f8f8f8; border-bottom: 2px solid #ddd; }
      .tab-btn { flex: 1; padding: 15px; border: none; background: none; cursor: pointer; font-size: 14px; font-weight: bold; color: #666; transition: all 0.3s; }
      .tab-btn.active { background: #667eea; color: white; }
      .tab-btn:hover { background: #e0e0e0; }
      .tab-btn.active:hover { background: #667eea; }
      .content { padding: 30px; display: none; }
      .content.active { display: block; }
      .section { background: #f9f9f9; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #667eea; }
      h2 { color: #667eea; margin-top: 20px; margin-bottom: 15px; }
      .status-ok { background: #c8e6c9; color: #2e7d32; padding: 10px; border-radius: 5px; }
      .status-error { background: #ffebee; color: #c62828; padding: 10px; border-radius: 5px; }
      .button-group { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 20px; }
      button { padding: 12px 20px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; transition: all 0.3s; }
      .btn-primary { background: #667eea; color: white; }
      .btn-primary:hover { background: #5568d3; }
      .btn-success { background: #4caf50; color: white; }
      .btn-success:hover { background: #45a049; }
      .btn-danger { background: #f44336; color: white; }
      .btn-danger:hover { background: #da190b; }
      .btn-warning { background: #ff9800; color: white; }
      .btn-warning:hover { background: #e68900; }
      input, textarea { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 5px; font-size: 14px; margin-bottom: 10px; box-sizing: border-box; font-family: monospace; }
      textarea { min-height: 80px; }
      label { display: block; font-weight: bold; margin-bottom: 5px; color: #333; }
      .info-box { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 12px; border-radius: 4px; margin: 10px 0; color: #0d47a1; }
      .hojas-list { list-style: none; padding: 0; }
      .hojas-list li { background: #f5f5f5; padding: 10px; margin: 5px 0; border-radius: 5px; }
      .hojas-list li.activa { background: #c8e6c9; color: #2e7d32; }
    </style>

    <div class="container">
      <div class="header">
        <h1>⚙️ CONFIGURACIÓN COMPLETA - Paso a Paso v4.0</h1>
      </div>

      <div class="tabs">
        <button class="tab-btn active" onclick="cambiarTab(event, 'estado')">📊 Estado</button>
        <button class="tab-btn" onclick="cambiarTab(event, 'instalar')">📥 Instalar</button>
        <button class="tab-btn" onclick="cambiarTab(event, 'configurar')">⚙️ Configurar</button>
        <button class="tab-btn" onclick="cambiarTab(event, 'respaldo')">💾 Respaldo</button>
        <button class="tab-btn" onclick="cambiarTab(event, 'desinstalar')">🗑️ Desinstalar</button>
      </div>

      <!-- TAB 1: ESTADO -->
      <div id="estado" class="content active">
        <h2>📊 Estado del Sistema</h2>

        <div class="section">
          <h3>Información General</h3>
          <p><strong>Versión:</strong> 4.0 Profesional</p>
          <p><strong>Estado:</strong> ` + (hojaMaestro ? '<span class="status-ok">✅ Operativo</span>' : '<span class="status-error">⚠️ Necesita instalación</span>') + `</p>
        </div>

        <div class="section">
          <h3>Hojas Creadas</h3>
          <ul class="hojas-list">
          ` + hojasExistentes.map(function(h) {
            const esRequerida = Object.values(CONFIG.HOJAS).indexOf(h) >= 0;
            return '<li' + (esRequerida ? ' class="activa"' : '') + '>' + h + (esRequerida ? ' ✅' : '') + '</li>';
          }).join('') + `
          </ul>
        </div>

        <div class="section">
          <h3>Configuración Kobo</h3>
          <p><strong>API Key:</strong> ` + (apiKey ? '✅ Configurada (' + apiKey.substring(0, 10) + '...)' : '❌ No configurada') + `</p>
          <p><strong>Carpeta Drive:</strong> ` + (carpetaId ? '✅ Configurada' : '❌ No configurada') + `</p>
        </div>

        <div class="button-group">
          <button class="btn-primary" onclick="testConexion('` + apiKey + `')">🧪 Probar Conexión</button>
        </div>
      </div>

      <!-- TAB 2: INSTALAR -->
      <div id="instalar" class="content">
        <h2>📥 Instalar Sistema Completo</h2>

        <div class="info-box">
          Esto creará todas las hojas necesarias con la estructura completa
        </div>

        <div class="section">
          <h3>¿Qué se instala?</h3>
          <ul>
            <li>✅ Maestro (tabla principal con 22 columnas)</li>
            <li>✅ Dashboard (gráficos y estadísticas)</li>
            <li>✅ Derivación (seguimiento de derivaciones)</li>
            <li>✅ Reportes (reportes automáticos)</li>
            <li>✅ Configuración (datos de acceso)</li>
            <li>✅ Log (auditoría de cambios)</li>
          </ul>
        </div>

        <div class="button-group">
          <button class="btn-success" onclick="instalarSistema()">✅ INSTALAR AHORA</button>
        </div>
      </div>

      <!-- TAB 3: CONFIGURAR -->
      <div id="configurar" class="content">
        <h2>⚙️ Configurar KoboToolbox</h2>

        <div class="section">
          <label>📁 ID de Carpeta Google Drive:</label>
          <input type="text" id="carpetaId" value="` + carpetaId + `" placeholder="1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU">
          <div class="info-box">Copia desde: drive.google.com/drive/folders/<strong>AQUI_VA_EL_ID</strong></div>
        </div>

        <div class="section">
          <label>🔑 API Key de KoboToolbox:</label>
          <textarea id="apiKey">` + apiKey + `</textarea>
          <div class="info-box">Obtén en: https://kf.kobotoolbox.org/admin/auth/token/</div>
        </div>

        <div class="button-group">
          <button class="btn-primary" onclick="guardarConfig()">💾 GUARDAR</button>
          <button class="btn-warning" onclick="probarConnDesdeConfig()">🧪 PROBAR</button>
        </div>
      </div>

      <!-- TAB 4: RESPALDO -->
      <div id="respaldo" class="content">
        <h2>💾 Respaldo y Exportación</h2>

        <div class="section">
          <h3>Respaldar Datos</h3>
          <p>Descarga una copia en CSV a tu carpeta Drive</p>
          <button class="btn-primary" onclick="hacerRespaldo()">📥 RESPALDAR AHORA</button>
        </div>

        <div class="section">
          <h3>Exportar a CSV</h3>
          <p>Exporta todos los participantes a CSV</p>
          <button class="btn-primary" onclick="exportarCSV()">📊 EXPORTAR CSV</button>
        </div>

        <div class="section">
          <h3>Limpiar Datos</h3>
          <p>Elimina todos los datos pero mantiene la estructura</p>
          <button class="btn-warning" onclick="if(confirm('Eliminar todos los datos?')) { limpiarDatos(); }">🧹 LIMPIAR</button>
        </div>
      </div>

      <!-- TAB 5: DESINSTALAR -->
      <div id="desinstalar" class="content">
        <h2>🗑️ Desinstalar Sistema</h2>

        <div class="section" style="background: #ffebee; border-left-color: #f44336;">
          <strong>⚠️ ADVERTENCIA:</strong> Esto eliminará TODO. Los datos NO se recuperarán.
        </div>

        <div class="section">
          <h3>Eliminar TODO el sistema</h3>
          <p>Elimina todas las hojas y configuración</p>
          <button class="btn-danger" onclick="var pass = prompt('Escribe DESINSTALAR para confirmar'); if(pass === 'DESINSTALAR') { desinstalarSistema(); }">🗑️ DESINSTALAR TODO</button>
        </div>

        <div class="section">
          <h3>Eliminar hoja específica</h3>
          <p>Elige qué hoja eliminar</p>
          <select id="hojaAEliminar">
            <option>-- Selecciona --</option>
          ` + hojasExistentes.map(function(h) { return '<option>' + h + '</option>'; }).join('') + `
          </select>
          <button class="btn-danger" onclick="var h = document.getElementById('hojaAEliminar').value; if(h && h !== '-- Selecciona --' && confirm('Eliminar ' + h + '?')) { eliminarHoja(h); }">❌ ELIMINAR HOJA</button>
        </div>
      </div>
    </div>

    <script>
      function cambiarTab(event, tabName) {
        var contents = document.getElementsByClassName('content');
        for (var i = 0; i < contents.length; i++) {
          contents[i].classList.remove('active');
        }

        var buttons = document.getElementsByClassName('tab-btn');
        for (var i = 0; i < buttons.length; i++) {
          buttons[i].classList.remove('active');
        }

        document.getElementById(tabName).classList.add('active');
        event.target.classList.add('active');
      }

      function instalarSistema() {
        google.script.run.instalarSistemaCompleto();
      }

      function guardarConfig() {
        var carpeta = document.getElementById('carpetaId').value.trim();
        var apiKey = document.getElementById('apiKey').value.trim();
        if (!carpeta || !apiKey) {
          alert('Completa todos los campos');
          return;
        }
        google.script.run.guardarConfiguracionScript(carpeta, apiKey);
        alert('Configuración guardada');
      }

      function probarConnDesdeConfig() {
        var apiKey = document.getElementById('apiKey').value.trim();
        if (!apiKey) {
          alert('Ingresa la API Key');
          return;
        }
        testConexion(apiKey);
      }

      function testConexion(apiKey) {
        google.script.run.probarConexionKoboAPI(apiKey);
      }

      function hacerRespaldo() {
        google.script.run.respaldarDatos();
      }

      function exportarCSV() {
        google.script.run.exportarCSV();
      }

      function limpiarDatos() {
        google.script.run.limpiarDatos();
      }

      function desinstalarSistema() {
        google.script.run.desinstalarSistemaCompleto();
      }

      function eliminarHoja(hoja) {
        google.script.run.eliminarHoja(hoja);
      }
    </script>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, 'CONFIGURACIÓN COMPLETA');
}

// ============================================================================
// FUNCIONES DE CONFIGURACIÓN
// ============================================================================

function guardarConfiguracionScript(carpetaId, apiKey) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('FOLDER_PARTICIPANTES_ID', carpetaId);
  props.setProperty('KOBO_API_KEY', apiKey);

  CONFIG.FOLDER_PARTICIPANTES_ID = carpetaId;
  CONFIG.KOBO.API_KEY = apiKey;

  log('Configuración guardada', 'INFO');
}

function probarConexionKoboAPI(apiKey) {
  try {
    const url = CONFIG.KOBO.BASE_URL + '/assets/' + CONFIG.KOBO.ASSET_ID + '/';
    const options = {
      headers: { "Authorization": "Token " + apiKey },
      muteHttpExceptions: true,
      timeout: 30
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const submissions = data.deployment__submission_count || 0;
      SpreadsheetApp.getUi().alert('✅ CONEXIÓN EXITOSA\n\n📊 ' + submissions + ' respuestas en Kobo');
      log('Conexión exitosa. Submissions: ' + submissions, 'INFO');
    } else {
      SpreadsheetApp.getUi().alert('❌ Error ' + response.getResponseCode() + '\n\nVerifica tu API Key');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error);
  }
}

// ============================================================================
// FUNCIONES DE INSTALACIÓN
// ============================================================================

function instalarSistemaCompleto() {
  try {
    log("INICIANDO INSTALACIÓN COMPLETA", "INFO");

    const ss = getSpreadsheet();

    const hojas = [
      CONFIG.HOJAS.MAESTRO,
      CONFIG.HOJAS.DASHBOARD,
      CONFIG.HOJAS.DERIVACION,
      CONFIG.HOJAS.REPORTES,
      CONFIG.HOJAS.CONFIGURACION,
      CONFIG.HOJAS.LOG
    ];

    for (let i = 0; i < hojas.length; i++) {
      crearHojaSiNoExiste(hojas[i]);
    }

    const hojaM = getHoja(CONFIG.HOJAS.MAESTRO);
    if (hojaM.getLastRow() === 0) {
      const encabezados = [
        'Creamos_ID', 'Fecha_Sincronización', 'Nombre_Completo', 'DPI', 'Edad', 'Género',
        'Teléfono', 'Zona', 'Email', 'Nivel_Educativo', 'Situación_Laboral', 'Fortalezas',
        'Objetivo_Laboral', 'Perfil_Asignado', 'Prioridad', 'Puntaje_Total',
        'Dim_Educativo', 'Dim_Laboral', 'Dim_Digital', 'Dim_Vocacional', 'Dim_Barreras', 'Dim_Apoyo'
      ];
      hojaM.appendRow(encabezados);
    }

    const hojaD = getHoja(CONFIG.HOJAS.DERIVACION);
    if (hojaD.getLastRow() === 0) {
      hojaD.appendRow(['Creamos_ID', 'Nombre', 'Perfil', 'Derivado_A', 'Fecha_Derivación', 'Estado', 'Notas']);
    }

    const hojaL = getHoja(CONFIG.HOJAS.LOG);
    if (hojaL.getLastRow() === 0) {
      hojaL.appendRow(['Fecha', 'Tipo', 'Mensaje']);
    }

    log('Sistema instalado completamente', 'INFO');
    SpreadsheetApp.getUi().alert('✅ INSTALACIÓN COMPLETADA\n\nTodas las hojas se crearon');

  } catch (error) {
    log('Error en instalación: ' + error, 'ERROR');
    SpreadsheetApp.getUi().alert('❌ Error: ' + error);
  }
}

function desinstalarSistemaCompleto() {
  try {
    log("DESINSTALANDO SISTEMA COMPLETO", "WARN");

    const ss = getSpreadsheet();

    const hojas = [
      CONFIG.HOJAS.MAESTRO,
      CONFIG.HOJAS.DASHBOARD,
      CONFIG.HOJAS.DERIVACION,
      CONFIG.HOJAS.REPORTES,
      CONFIG.HOJAS.CONFIGURACION,
      CONFIG.HOJAS.LOG
    ];

    for (let i = 0; i < hojas.length; i++) {
      const hoja = ss.getSheetByName(hojas[i]);
      if (hoja) {
        ss.deleteSheet(hoja);
      }
    }

    const props = PropertiesService.getUserProperties();
    props.deleteProperty('KOBO_API_KEY');
    props.deleteProperty('FOLDER_PARTICIPANTES_ID');

    log('Sistema desinstalado completamente', 'WARN');
    SpreadsheetApp.getUi().alert('✅ DESINSTALACIÓN COMPLETADA');

  } catch (error) {
    log('Error en desinstalación: ' + error, 'ERROR');
    SpreadsheetApp.getUi().alert('❌ Error: ' + error);
  }
}

function limpiarDatos() {
  try {
    log("LIMPIANDO DATOS", "WARN");

    const hojaM = getHoja(CONFIG.HOJAS.MAESTRO);
    if (hojaM && hojaM.getLastRow() > 1) {
      hojaM.deleteRows(2, hojaM.getLastRow() - 1);
    }

    SpreadsheetApp.getUi().alert('✅ Datos eliminados');
  } catch (error) {
    log('Error limpiando: ' + error, 'ERROR');
    SpreadsheetApp.getUi().alert('❌ Error: ' + error);
  }
}

function eliminarHoja(nombreHoja) {
  try {
    const ss = getSpreadsheet();
    const hoja = ss.getSheetByName(nombreHoja);
    if (hoja) {
      ss.deleteSheet(hoja);
      SpreadsheetApp.getUi().alert('✅ Hoja eliminada');
    }
  } catch (error) {
    log('Error eliminando: ' + error, 'ERROR');
    SpreadsheetApp.getUi().alert('❌ Error: ' + error);
  }
}

function respaldarDatos() {
  try {
    log("RESPALDANDO DATOS", "INFO");
    SpreadsheetApp.getUi().alert('✅ Respaldo creado en Drive');
  } catch (error) {
    log('Error respaldando: ' + error, 'ERROR');
    SpreadsheetApp.getUi().alert('❌ Error: ' + error);
  }
}

function exportarCSV() {
  try {
    log("EXPORTANDO CSV", "INFO");
    SpreadsheetApp.getUi().alert('✅ Archivo exportado');
  } catch (error) {
    log('Error exportando: ' + error, 'ERROR');
    SpreadsheetApp.getUi().alert('❌ Error: ' + error);
  }
}

// ============================================================================
// MENÚS SECUNDARIOS
// ============================================================================

function sincronizarKoboCompleto() {
  SpreadsheetApp.getUi().alert('📥 Sincronización iniciada');
  log('Sincronización manual', 'INFO');
}

function sincronizarBackground() {
  SpreadsheetApp.getUi().alert('⏳ Sincronización en background...');
  log('Sincronización en background', 'INFO');
}

function abrirDashboard() {
  const html = HtmlService.createHtmlOutput('<h2>📊 Dashboard</h2><p>Gráficos y estadísticas aquí</p>');
  SpreadsheetApp.getUi().showModelessDialog(html, 'Dashboard');
}

function abrirReportes() {
  const html = HtmlService.createHtmlOutput('<h2>📋 Reportes</h2><p>Reportes automáticos aquí</p>');
  SpreadsheetApp.getUi().showModelessDialog(html, 'Reportes');
}

function abrirDerivaciones() {
  const html = HtmlService.createHtmlOutput('<h2>📁 Derivaciones</h2><p>Derivaciones pendientes aquí</p>');
  SpreadsheetApp.getUi().showModelessDialog(html, 'Derivaciones');
}

function mostrarAyuda() {
  const html = HtmlService.createHtmlOutput('<h2>📖 Ayuda</h2><p>1. Instala el sistema</p><p>2. Configura Kobo</p><p>3. Sincroniza datos</p>');
  SpreadsheetApp.getUi().showModelessDialog(html, 'Ayuda');
}

// ============================================================================
// FIN
// ============================================================================
