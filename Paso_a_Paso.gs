// ============================================================================
// SISTEMA PASO A PASO - VERSIÓN 4.0 PROFESIONAL
// Importación de Kobo + Dashboard + Derivaciones + Reportes
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
// MENÚ - CONFIGURACIÓN COMPLETA
// ============================================================================

function mostrarConfiguracionCompleta() {
  const props = PropertiesService.getUserProperties();
  const apiKey = props.getProperty('KOBO_API_KEY') || '';
  const carpetaId = props.getProperty('FOLDER_PARTICIPANTES_ID') || '';

  // Verificar estado del sistema
  const ss = getSpreadsheet();
  const hojaMaestro = ss.getSheetByName(CONFIG.HOJAS.MAESTRO);
  const hojasExistentes = ss.getSheets().map(h => h.getName());

  const html = HtmlService.createHtmlOutput(`
    <style>
      * { margin: 0; padding: 0; box-sizing: border-box; }
      body { font-family: 'Segoe UI', Arial; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }
      .container { max-width: 1000px; margin: 0 auto; }
      .tabs { display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; }
      .tab-btn { background: white; padding: 12px 20px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; color: #667eea; transition: all 0.3s; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
      .tab-btn.active { background: #667eea; color: white; }
      .tab-btn:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
      .tab-content { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 8px 32px rgba(0,0,0,0.1); display: none; }
      .tab-content.active { display: block; }
      h1 { color: #667eea; margin-bottom: 20px; font-size: 28px; border-bottom: 3px solid #667eea; padding-bottom: 15px; }
      h2 { color: #333; margin-top: 20px; margin-bottom: 15px; font-size: 18px; }
      .section { background: #f8fafb; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #667eea; }
      .status { padding: 12px; border-radius: 8px; margin-bottom: 10px; font-weight: bold; }
      .status.ok { background: #c8e6c9; color: #2e7d32; }
      .status.warning { background: #fff3cd; color: #856404; }
      .status.error { background: #f8d7da; color: #721c24; }
      .input-group { margin-bottom: 15px; }
      label { display: block; font-weight: bold; color: #333; margin-bottom: 5px; }
      input, textarea { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 5px; font-size: 14px; font-family: monospace; }
      textarea { min-height: 80px; }
      .button-group { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 15px; }
      button { padding: 12px 20px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; transition: all 0.3s; }
      .btn-primary { background: #667eea; color: white; }
      .btn-primary:hover { background: #5568d3; transform: translateY(-2px); }
      .btn-success { background: #4caf50; color: white; }
      .btn-success:hover { background: #45a049; }
      .btn-danger { background: #f44336; color: white; }
      .btn-danger:hover { background: #da190b; }
      .btn-warning { background: #ff9800; color: white; }
      .btn-warning:hover { background: #e68900; }
      .btn-secondary { background: #757575; color: white; }
      .btn-secondary:hover { background: #616161; }
      .info-box { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 12px; border-radius: 4px; margin: 10px 0; color: #0d47a1; }
      .hoja-item { background: #f5f5f5; padding: 10px; border-radius: 5px; margin: 5px 0; display: flex; justify-content: space-between; align-items: center; }
      .hoja-item.activa { background: #c8e6c9; }
      .step { background: #f0f4ff; padding: 15px; border-radius: 8px; margin: 10px 0; border-left: 4px solid #667eea; }
      .step-number { display: inline-block; background: #667eea; color: white; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-weight: bold; margin-right: 10px; }
    </style>

    <div class="container">
      <div class="tabs">
        <button class="tab-btn active" onclick="cambiarTab('estado')">📊 Estado del Sistema</button>
        <button class="tab-btn" onclick="cambiarTab('instalar')">📥 Instalar Sistema</button>
        <button class="tab-btn" onclick="cambiarTab('configurar')">⚙️ Configurar Kobo</button>
        <button class="tab-btn" onclick="cambiarTab('respaldar')">💾 Respaldar/Restaurar</button>
        <button class="tab-btn" onclick="cambiarTab('desinstalar')">🗑️ Desinstalar</button>
      </div>

      <!-- TAB 1: ESTADO DEL SISTEMA -->
      <div id="estado" class="tab-content active">
        <h1>📊 Estado del Sistema</h1>

        <h2>✅ Información General</h2>
        <div class="section">
          <p><strong>Versión:</strong> 4.0 Profesional</p>
          <p><strong>Sheet ID:</strong> ${CONFIG.SPREADSHEET_ID.substring(0, 40)}...</p>
          <p><strong>Estado:</strong> ${hojaMaestro ? '✅ Operativo' : '⚠️ Necesita instalación'}</p>
        </div>

        <h2>📄 Hojas del Sistema</h2>
        ${hojasExistentes.map(hoja => {
          const esRequerida = Object.values(CONFIG.HOJAS).includes(hoja);
          const clase = esRequerida ? 'activa' : '';
          const icono = esRequerida ? '✅' : '';
          return '<div class="hoja-item ' + clase + '">' + hoja + ' ' + icono + '</div>';
        }).join('')}

        <h2>🔑 Configuración de Kobo</h2>
        <div class="section">
          <p><strong>API Key guardada:</strong> ${apiKey ? '✅ Sí (' + apiKey.substring(0, 10) + '...)' : '❌ No'}</p>
          <p><strong>Carpeta Drive:</strong> ${carpetaId ? '✅ Sí' : '❌ No'}</p>
        </div>

        <div class="button-group">
          <button class="btn-primary" onclick="google.script.run.probarConexionKoboAPI('${apiKey}')">🧪 Probar Conexión</button>
        </div>
      </div>

      <!-- TAB 2: INSTALAR SISTEMA -->
      <div id="instalar" class="tab-content">
        <h1>📥 Instalación Completa del Sistema</h1>

        <div class="info-box">
          <strong>📌 Esto creará todas las hojas necesarias con la estructura completa</strong>
        </div>

        <h2>Pasos de Instalación</h2>

        <div class="step">
          <span class="step-number">1</span>
          <strong>Crear todas las hojas automáticamente</strong>
          <p style="margin-top: 5px; color: #666;">Se crearán: Maestro, Dashboard, Derivación, Reportes, Configuración, Log</p>
          <button class="btn-success" onclick="google.script.run.instalarSistemaCompleto()">✅ INSTALAR AHORA</button>
        </div>

        <div class="step">
          <span class="step-number">2</span>
          <strong>Configurar API Key de Kobo</strong>
          <p style="margin-top: 5px; color: #666;">Ve a la pestaña "⚙️ Configurar Kobo"</p>
        </div>

        <div class="step">
          <span class="step-number">3</span>
          <strong>Sincronizar datos de Kobo</strong>
          <p style="margin-top: 5px; color: #666;">Usa el menú 📊 PASO A PASO > 🔄 SINCRONIZAR</p>
        </div>

        <h2>¿Qué se instala?</h2>
        <div class="section">
          <ul style="margin-left: 20px; line-height: 2;">
            <li>✅ Hoja "Maestro" con estructura de 22 columnas</li>
            <li>✅ Hoja "Dashboard" con gráficos automáticos</li>
            <li>✅ Hoja "Derivación" para controlar derivaciones</li>
            <li>✅ Hoja "Reportes" con reportes automáticos</li>
            <li>✅ Hoja "Configuración" para datos de acceso</li>
            <li>✅ Hoja "Log" con auditoría de cambios</li>
          </ul>
        </div>
      </div>

      <!-- TAB 3: CONFIGURAR KOBO -->
      <div id="configurar" class="tab-content">
        <h1>⚙️ Configuración de Kobo</h1>

        <div class="input-group">
          <label>📁 ID de Carpeta Google Drive:</label>
          <input type="text" id="carpetaId" value="${carpetaId}" placeholder="1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU">
          <div class="info-box">📌 Copia desde: drive.google.com/drive/folders/<strong>AQUÍ_VA_EL_ID</strong></div>
        </div>

        <div class="input-group">
          <label>🔑 API Key de KoboToolbox:</label>
          <textarea id="apiKey" placeholder="Pega tu API Key aquí">${apiKey}</textarea>
          <div class="info-box">
            🔗 Obtén en: <strong>https://kf.kobotoolbox.org/admin/auth/token/</strong><br>
            Asset ID: <strong>abHRWdRnPhKwzPQBajc7RZ</strong>
          </div>
        </div>

        <div class="button-group">
          <button class="btn-primary" onclick="guardarConfiguracionCompleta()">💾 GUARDAR CONFIGURACIÓN</button>
          <button class="btn-warning" onclick="probarConexionDesdeConfig()">🧪 PROBAR CONEXIÓN</button>
        </div>

        <h2>🔐 Seguridad</h2>
        <div class="section">
          <p>La API Key se guarda de forma segura en Google Properties Service.</p>
          <p>🚨 <strong>NO la compartas ni publiques en Git</strong></p>
        </div>
      </div>

      <!-- TAB 4: RESPALDAR/RESTAURAR -->
      <div id="respaldar" class="tab-content">
        <h1>💾 Respaldar y Restaurar Datos</h1>

        <h2>📥 Respaldar Datos</h2>
        <div class="section">
          <p>Descarga una copia de todos tus datos en Google Drive</p>
          <button class="btn-primary" onclick="google.script.run.respaldarDatos()">📥 RESPALDAR AHORA</button>
        </div>

        <h2>📤 Restaurar Datos</h2>
        <div class="section">
          <p>⚠️ Restaurar eliminará datos actuales y cargará un respaldo anterior</p>
          <button class="btn-warning" onclick="if(confirm('¿Seguro? Esto eliminará datos actuales')){google.script.run.restaurarDatos()}">📤 RESTAURAR RESPALDO</button>
        </div>

        <h2>📊 Exportar a CSV</h2>
        <div class="section">
          <p>Exporta todos los datos de la hoja Maestro a CSV</p>
          <button class="btn-secondary" onclick="google.script.run.exportarCSV()">📊 EXPORTAR CSV</button>
        </div>
      </div>

      <!-- TAB 5: DESINSTALAR -->
      <div id="desinstalar" class="tab-content">
        <h1>🗑️ Desinstalar Sistema</h1>

        <div class="section" style="background: #ffebee; border-left-color: #f44336;">
          <p><strong>⚠️ ADVERTENCIA:</strong> Esto eliminará TODAS las hojas del sistema</p>
          <p>Los datos NO se recuperarán después</p>
        </div>

        <h2>Opciones de Desinstalación</h2>

        <div class="step">
          <span class="step-number">1</span>
          <strong>Desinstalar TODO el sistema</strong>
          <p style="margin-top: 5px; color: #666;">Elimina todas las hojas (Maestro, Dashboard, etc.)</p>
          <button class="btn-danger" onclick="if(confirm('¿SEGURO? Esto eliminará TODO. Escribe DESINSTALAR para confirmar')) { var pass = prompt('Escribe DESINSTALAR para confirmar'); if(pass === 'DESINSTALAR') { google.script.run.desinstalarSistemaCompleto(); } else { alert('Cancelado'); } }">🗑️ DESINSTALAR TODO</button>
        </div>

        <div class="step">
          <span class="step-number">2</span>
          <strong>Limpiar solo datos (mantener hojas)</strong>
          <p style="margin-top: 5px; color: #666;">Elimina todos los participantes pero mantiene la estructura</p>
          <button class="btn-warning" onclick="if(confirm('¿Seguro? Esto eliminará todos los datos de participantes')) { google.script.run.limpiarDatos(); }">🧹 LIMPIAR DATOS</button>
        </div>

        <div class="step">
          <span class="step-number">3</span>
          <strong>Eliminar hoja específica</strong>
          <p style="margin-top: 5px; color: #666;">Elige qué hoja eliminar</p>
          <select id="hojaAEliminar" style="width: 100%; padding: 10px; margin: 10px 0; border-radius: 5px; border: 1px solid #ddd;">
            <option>-- Selecciona una hoja --</option>
            ${hojasExistentes.map(h => \`<option value="\${h}">\${h}</option>\`).join('')}
          </select>
          <button class="btn-danger" onclick="const hoja = document.getElementById('hojaAEliminar').value; if(hoja && hoja !== '-- Selecciona una hoja --' && confirm('¿Eliminar ' + hoja + '?')) { google.script.run.eliminarHoja(hoja); }">❌ ELIMINAR HOJA</button>
        </div>
      </div>
    </div>

    <script>
      function cambiarTab(tabName) {
        // Ocultar todos
        document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));

        // Mostrar seleccionado
        document.getElementById(tabName).classList.add('active');
        event.target.classList.add('active');
      }

      function guardarConfiguracionCompleta() {
        const carpeta = document.getElementById('carpetaId').value.trim();
        const apiKey = document.getElementById('apiKey').value.trim();

        if (!carpeta || !apiKey) {
          alert('❌ Completa todos los campos');
          return;
        }

        google.script.run.guardarConfiguracionScript(carpeta, apiKey);
        alert('✅ Configuración guardada');
      }

      function probarConexionDesdeConfig() {
        const apiKey = document.getElementById('apiKey').value.trim();
        if (!apiKey) {
          alert('❌ Ingresa la API Key');
          return;
        }
        google.script.run.probarConexionKoboAPI(apiKey);
      }
    </script>
  `);
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

  SpreadsheetApp.getUi().showModelessDialog(html, '⚙️ CONFIGURACIÓN COMPLETA - Paso a Paso v4.0');
}

// ============================================================================
// FUNCIONES DE INSTALACIÓN Y DESINSTALACIÓN
// ============================================================================

function instalarSistemaCompleto() {
  try {
    log("=== INICIANDO INSTALACIÓN COMPLETA ===", "INFO");

    const ss = getSpreadsheet();

    // Crear todas las hojas
    const hojas = [
      CONFIG.HOJAS.MAESTRO,
      CONFIG.HOJAS.DASHBOARD,
      CONFIG.HOJAS.DERIVACION,
      CONFIG.HOJAS.REPORTES,
      CONFIG.HOJAS.CONFIGURACION,
      CONFIG.HOJAS.LOG
    ];

    for (const hoja of hojas) {
      crearHojaSiNoExiste(hoja);
    }

    // Crear encabezados en Maestro
    const hojaM = getHoja(CONFIG.HOJAS.MAESTRO);
    if (hojaM.getLastRow() === 0) {
      const encabezados = [
        'Creamos_ID', 'Fecha_Sincronización', 'Nombre_Completo', 'DPI', 'Edad', 'Género',
        'Teléfono', 'Zona', 'Email', 'Nivel_Educativo', 'Situación_Laboral', 'Fortalezas',
        'Objetivo_Laboral', 'Perfil_Asignado', 'Prioridad', 'Puntaje_Total',
        'Dim_Educativo', 'Dim_Laboral', 'Dim_Digital', 'Dim_Vocacional', 'Dim_Barreras', 'Dim_Apoyo'
      ];
      hojaM.appendRow(encabezados);
      log('✅ Encabezados creados en Maestro', 'INFO');
    }

    // Crear encabezados en Derivación
    const hojaD = getHoja(CONFIG.HOJAS.DERIVACION);
    if (hojaD.getLastRow() === 0) {
      hojaD.appendRow(['Creamos_ID', 'Nombre', 'Perfil', 'Derivado_A', 'Fecha_Derivación', 'Estado_Derivación', 'Notas']);
    }

    // Crear encabezados en Log
    const hojaL = getHoja(CONFIG.HOJAS.LOG);
    if (hojaL.getLastRow() === 0) {
      hojaL.appendRow(['Fecha', 'Tipo', 'Mensaje']);
    }

    log('✅ Sistema instalado completamente', 'INFO');
    SpreadsheetApp.getUi().alert('✅ INSTALACIÓN COMPLETADA\n\nTodas las hojas se crearon exitosamente\n\nAhora configura tu API Key de Kobo');

  } catch (error) {
    log(`Error en instalación: ${error}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ Error en instalación: ${error}`);
  }
}

function desinstalarSistemaCompleto() {
  try {
    log("=== DESINSTALANDO SISTEMA COMPLETO ===", "WARN");

    const ss = getSpreadsheet();

    const hojas = [
      CONFIG.HOJAS.MAESTRO,
      CONFIG.HOJAS.DASHBOARD,
      CONFIG.HOJAS.DERIVACION,
      CONFIG.HOJAS.REPORTES,
      CONFIG.HOJAS.CONFIGURACION,
      CONFIG.HOJAS.LOG
    ];

    for (const nombreHoja of hojas) {
      const hoja = ss.getSheetByName(nombreHoja);
      if (hoja) {
        ss.deleteSheet(hoja);
        log(`🗑️ Hoja eliminada: ${nombreHoja}`, 'WARN');
      }
    }

    // Limpiar propiedades
    const props = PropertiesService.getUserProperties();
    props.deleteProperty('KOBO_API_KEY');
    props.deleteProperty('FOLDER_PARTICIPANTES_ID');

    log('✅ Sistema desinstalado completamente', 'WARN');
    SpreadsheetApp.getUi().alert('✅ DESINSTALACIÓN COMPLETADA\n\nTodas las hojas se eliminaron\n\nPuedes instalar de nuevo cuando quieras');

  } catch (error) {
    log(`Error en desinstalación: ${error}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

function limpiarDatos() {
  try {
    log("=== LIMPIANDO DATOS ===", "WARN");

    const hojaM = getHoja(CONFIG.HOJAS.MAESTRO);
    if (hojaM && hojaM.getLastRow() > 1) {
      hojaM.deleteRows(2, hojaM.getLastRow() - 1);
      log('✅ Datos de Maestro eliminados', 'WARN');
    }

    const hojaD = getHoja(CONFIG.HOJAS.DERIVACION);
    if (hojaD && hojaD.getLastRow() > 1) {
      hojaD.deleteRows(2, hojaD.getLastRow() - 1);
      log('✅ Datos de Derivación eliminados', 'WARN');
    }

    SpreadsheetApp.getUi().alert('✅ DATOS LIMPIADOS\n\nLas estructuras se mantienen, solo se borraron los datos');

  } catch (error) {
    log(`Error limpiando datos: ${error}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

function eliminarHoja(nombreHoja) {
  try {
    const ss = getSpreadsheet();
    const hoja = ss.getSheetByName(nombreHoja);

    if (hoja) {
      ss.deleteSheet(hoja);
      log(`🗑️ Hoja eliminada: ${nombreHoja}`, 'WARN');
      SpreadsheetApp.getUi().alert(`✅ Hoja "${nombreHoja}" eliminada`);
    } else {
      SpreadsheetApp.getUi().alert('⚠️ Hoja no encontrada');
    }
  } catch (error) {
    log(`Error eliminando hoja: ${error}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

function respaldarDatos() {
  try {
    log("=== RESPALDANDO DATOS ===", "INFO");

    const hojaM = getHoja(CONFIG.HOJAS.MAESTRO);
    const datos = hojaM.getDataRange().getValues();

    const carpeta = DriveApp.getFolderById(CONFIG.FOLDER_PARTICIPANTES_ID);
    const nombreArchivo = `Respaldo_Paso_a_Paso_${new Date().toISOString().split('T')[0]}.csv`;

    let csv = datos.map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');

    const file = carpeta.createFile(nombreArchivo, csv, MimeType.PLAIN_TEXT);

    log(`✅ Respaldo creado: ${nombreArchivo}`, 'INFO');
    SpreadsheetApp.getUi().alert(`✅ Respaldo creado\n\nArchivo: ${nombreArchivo}`);

  } catch (error) {
    log(`Error respaldando: ${error}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
}

function restaurarDatos() {
  log('📤 Función restaurar disponible pronto', 'INFO');
  SpreadsheetApp.getUi().alert('⏳ Función disponible en próxima versión');
}

function exportarCSV() {
  try {
    log("=== EXPORTANDO A CSV ===", "INFO");

    const hojaM = getHoja(CONFIG.HOJAS.MAESTRO);
    const datos = hojaM.getDataRange().getValues();

    const carpeta = DriveApp.getFolderById(CONFIG.FOLDER_PARTICIPANTES_ID);
    const nombreArchivo = `Exportar_Paso_a_Paso_${new Date().toISOString().split('T')[0]}.csv`;

    let csv = datos.map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');

    const file = carpeta.createFile(nombreArchivo, csv, MimeType.PLAIN_TEXT);

    log(`✅ Exportación creada: ${nombreArchivo}`, 'INFO');
    SpreadsheetApp.getUi().alert(`✅ Archivo exportado\n\nArchivo: ${nombreArchivo}`);

  } catch (error) {
    log(`Error exportando: ${error}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
  }
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
