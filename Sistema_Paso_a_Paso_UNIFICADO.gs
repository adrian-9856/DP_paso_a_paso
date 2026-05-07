// CONFIGURACIÓN CENTRALIZADA DEL SISTEMA PASO A PASO

const CONFIG = {
  // IDs de Google Sheets y Drives
  SPREADSHEET_ID: "", // EDITAR: Reemplazar con ID del Sheet maestro
  SHEET_NAME: "Maestro", // Nombre de la hoja principal
  FOLDER_PARTICIPANTES_ID: "", // EDITAR: ID de carpeta raíz para participantes

  // URLs y Credenciales Kobo
  KOBO_CSV_URL: "https://kf.kobotoolbox.org/api/v2/assets/abHRWdRnPhKwzPQBajc7RZ/export-settings/esxXsxLthEAgNHbXQnV8qYV/data.csv",
  KOBO_USERNAME: "", // EDITAR: Usuario Kobo
  KOBO_PASSWORD: "", // EDITAR: Contraseña Kobo

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
    HORA_ALERTAS: 9, // 9 AM
    HORA_IMPORTAR_KOBO: 6 // 6 AM
  }
};

/**
 * Obtiene el spreadsheet maestro
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/**
 * Obtiene la hoja maestra
 */
function getHojaMatriz() {
  return getSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
}

/**
 * Obtiene la carpeta raíz de participantes
 */
function getCarpetaRaiz() {
  return DriveApp.getFolderById(CONFIG.FOLDER_PARTICIPANTES_ID);
}

/**
 * Log centralizado (para debugging)
 */
function log(mensaje, tipo = "INFO") {
  const timestamp = new Date().toLocaleString("es-GT");
  const linea = `[${timestamp}] [${tipo}] ${mensaje}`;
  Logger.log(linea);
  console.log(linea);
}
/**
 * SCRIPT: Importar datos de KoboToolbox
 * Función: Descarga CSV de Kobo, detecta nuevos registros y los agrega a Sheets
 * Trigger: Diario 6 AM
 */

/**
 * Función principal para importar datos de Kobo
 */
function importarDatosKobo() {
  log("=== INICIANDO IMPORTACIÓN KOBO ===", "INFO");

  try {
    // 1. Descargar CSV de Kobo
    const csvData = descargarCSVKobo();
    if (!csvData) {
      log("No se pudo descargar CSV de Kobo", "ERROR");
      return;
    }

    // 2. Parsear CSV
    const registros = parsearCSV(csvData);
    if (registros.length === 0) {
      log("CSV vacío o sin registros válidos", "WARN");
      return;
    }

    log(`CSV descargado: ${registros.length} registros encontrados`, "INFO");

    // 3. Obtener últimos registros en Sheets para comparar
    const hoja = getHojaMatriz();
    const datosExistentes = hoja.getDataRange().getValues();
    const dpisExistentes = new Set(datosExistentes.slice(1).map(row => row[CONFIG.COLUMNAS.DPI - 1]));

    // 4. Filtrar registros nuevos
    const registrosNuevos = registros.filter(reg => {
      const dpi = reg.DPI || "";
      return dpi && !dpisExistentes.has(dpi);
    });

    log(`Registros nuevos detectados: ${registrosNuevos.length}`, "INFO");

    if (registrosNuevos.length === 0) {
      log("No hay registros nuevos para importar", "INFO");
      return;
    }

    // 5. Procesar cada registro nuevo
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

    // 6. Enviar resumen por email
    if (registrosProcesados > 0) {
      enviarResumenImportacion(registrosProcesados);
    }

  } catch (error) {
    log(`Error en importación general: ${error}`, "ERROR");
    enviarAlertaError("Error en importación Kobo", error);
  }
}

/**
 * Descarga CSV de Kobo con autenticación básica
 */
function descargarCSVKobo() {
  try {
    const url = CONFIG.KOBO_CSV_URL;
    const auth = Utilities.base64Encode(`${CONFIG.KOBO_USERNAME}:${CONFIG.KOBO_PASSWORD}`);

    const options = {
      headers: {
        "Authorization": `Basic ${auth}`
      },
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

/**
 * Parsea CSV a array de objetos
 */
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

/**
 * Agrega un nuevo participante a Sheets
 */
function agregarNuevoParticipante(datosKobo) {
  const hoja = getHojaMatriz();

  // Generar ID Creamos
  const idCreamos = generarIdCreamos(datosKobo);

  // Preparar fila nueva
  const nuevaFila = [];
  const numColumnas = 31; // Ajustar según total de columnas

  for (let i = 1; i <= numColumnas; i++) {
    nuevaFila.push("");
  }

  // Llenar datos disponibles
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

  // Agregar fila
  hoja.appendRow(nuevaFila);

  log(`Participante agregado: ${idCreamos} - ${datosKobo.NOMBRE_COMPLETO}`, "INFO");

  // Crear carpeta y expediente
  crearCarpetaYExpediente(idCreamos, datosKobo.NOMBRE_COMPLETO);
}

/**
 * Genera ID Creamos basado en nombre y DPI
 * Formato: PRIMERA_LETRA_NOMBRE + ÚLTIMOS_6_DPI + FECHA (DDMM)
 */
function generarIdCreamos(datos) {
  const nombre = datos.NOMBRE_COMPLETO || datos.NOMBRE_COMPLETO_SEGUN_DPI || "X";
  const dpi = datos.DPI || "000000";
  const primerLetra = nombre.charAt(0).toUpperCase();
  const ultimosDpi = dpi.slice(-6).padStart(6, "0");
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "ddMM");

  return `${primerLetra}${ultimosDpi.substring(0, 3)}${fecha}`;
}

/**
 * Envía resumen de importación por email
 */
function enviarResumenImportacion(cantidad) {
  const asunto = `[Paso a Paso] Importación Kobo - ${cantidad} nuevos participantes`;
  const cuerpo = `
    Se completó la importación de datos de KoboToolbox.

    Participantes agregados: ${cantidad}
    Fecha/Hora: ${new Date().toLocaleString("es-GT")}

    Accede al sistema para revisar los nuevos registros.

    Sistema Paso a Paso
  `;

  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, asunto, cuerpo);
}

/**
 * Envía alerta de error
 */
function enviarAlertaError(asunto, error) {
  const cuerpo = `
    Error detectado en el sistema Paso a Paso:

    ${asunto}
    ${error}

    Revisa los logs para más información.
  `;

  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, `[ALERTA] ${asunto}`, cuerpo);
}

/**
 * Crea carpeta y expediente digital para el participante
 * (Implementado en otro script: crear_documentos.gs)
 */
function crearCarpetaYExpediente(idCreamos, nombreCompleto) {
  // Ver script crear_documentos.gs para implementación
  log(`Pendiente crear carpeta para: ${idCreamos}`, "INFO");
}
/**
 * SCRIPT: Crear Carpeta y Expediente Digital
 * Función: Automatiza creación de carpeta en Drive y documento expediente
 * Triggers: Nuevo participante O cambio de estado a "Mentoría"
 */

/**
 * Crea carpeta para participante en Drive
 * @param {string} idCreamos - ID único del participante
 * @param {string} nombreCompleto - Nombre completo del participante
 * @returns {string} ID de la carpeta creada
 */
function crearCarpetaParticipante(idCreamos, nombreCompleto) {
  try {
    const raiz = getCarpetaRaiz();
    const nombreCarpeta = `[${idCreamos}] ${nombreCompleto}`;

    // Verificar si ya existe
    const carpetasExistentes = raiz.getFoldersByName(nombreCarpeta);
    if (carpetasExistentes.hasNext()) {
      const carpeta = carpetasExistentes.next();
      log(`Carpeta ya existe: ${nombreCarpeta}`, "WARN");
      return carpeta.getId();
    }

    // Crear carpeta
    const carpeta = raiz.createFolder(nombreCarpeta);
    log(`Carpeta creada: ${nombreCarpeta} (${carpeta.getId()})`, "INFO");

    return carpeta.getId();

  } catch (error) {
    log(`Error creando carpeta para ${idCreamos}: ${error}`, "ERROR");
    return null;
  }
}

/**
 * Crea Expediente Digital (Google Doc)
 * @param {string} idCreamos - ID único del participante
 * @param {string} nombreCompleto - Nombre completo
 * @param {string} carpetaId - ID de la carpeta padre
 * @param {object} datosParticipante - Datos del participante desde Sheets
 * @returns {string} ID del documento creado
 */
function crearExpedienteDigital(idCreamos, nombreCompleto, carpetaId, datosParticipante = {}) {
  try {
    const carpeta = DriveApp.getFolderById(carpetaId);
    const nombreDoc = `EXPEDIENTE - ${idCreamos} - ${nombreCompleto}`;

    // Crear documento
    const documento = carpeta.createDocument(nombreDoc);
    const docId = documento.getId();

    // Obtener contenido
    const contenido = documento.getBody();

    // Encabezado
    contenido.appendParagraph(`EXPEDIENTE DIGITAL`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    contenido.appendParagraph(`ID Creamos: ${idCreamos}`)
      .setBold(true);

    // Datos básicos
    contenido.appendParagraph(`DATOS BÁSICOS`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const tabla1 = contenido.appendTable([
      ["Campo", "Valor"],
      ["Nombre Completo", nombreCompleto],
      ["DPI", datosParticipante.dpi || ""],
      ["Edad", datosParticipante.edad || ""],
      ["Género", datosParticipante.genero || ""],
      ["Teléfono", datosParticipante.telefono || ""],
      ["Zona", datosParticipante.zona || ""],
      ["Nivel Educativo", datosParticipante.nivel_educativo || ""],
      ["Programa", datosParticipante.programa || ""],
      ["Fecha Registro", datosParticipante.fecha_registro || ""]
    ]);

    // Estilo tabla
    tabla1.getRow(0).getElement().getChild(0).asBold().setBackground("#4285F4");
    tabla1.getRow(0).getElement().getChild(1).asBold().setBackground("#4285F4");
    tabla1.getRow(0).getElement().getChild(0).asText().setForegroundColor("#ffffff");
    tabla1.getRow(0).getElement().getChild(1).asText().setForegroundColor("#ffffff");

    // Perfilado
    contenido.appendParagraph(`PERFILADO INICIAL`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const tabla2 = contenido.appendTable([
      ["Dimensión", "Valor"],
      ["Nivel Educativo", datosParticipante.nivel_educativo || "Pendiente"],
      ["Problemas Vitales", datosParticipante.problemas_vitales || "Pendiente"],
      ["Conciencia del Participante", datosParticipante.conciencia || "Por evaluar"],
      ["Perfil Resultante", datosParticipante.perfil || "Por determinar"]
    ]);

    // Estado
    contenido.appendParagraph(`SEGUIMIENTO DE ESTADO`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const tabla3 = contenido.appendTable([
      ["Estado Actual", datosParticipante.estado || "Orientación"],
      ["Fecha Cambio", datosParticipante.fecha_cambio || new Date().toLocaleDateString("es-GT")],
      ["Responsable", datosParticipante.responsable || "Por asignar"],
      ["Notas", ""]
    ]);

    // Mentoría (si aplica)
    if (datosParticipante.estado === "Mentoría") {
      contenido.appendParagraph(`SEGUIMIENTO DE MENTORÍA`)
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);

      const tabla4 = contenido.appendTable([
        ["Métrica", "Valor"],
        ["Fecha Inicio Mentoría", datosParticipante.fecha_inicio || ""],
        ["Sesiones Completadas", "0"],
        ["Última Sesión", "Pendiente"],
        ["Próxima Sesión", "Pendiente"],
        ["Duración (meses)", "0"],
        ["Alerta", "🟢 En progreso"]
      ]);
    }

    // Sección de Notas
    contenido.appendParagraph(`NOTAS Y OBSERVACIONES`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    contenido.appendParagraph(`Fecha | Responsable | Nota`)
      .setBold(true);

    for (let i = 0; i < 10; i++) {
      contenido.appendParagraph("___ | ___ | ___");
    }

    // Pie
    contenido.appendParagraph(`\nDocument creado: ${new Date().toLocaleString("es-GT")}`)
      .setItalic(true)
      .setFontSize(10);

    log(`Expediente creado: ${nombreDoc} (${docId})`, "INFO");

    return docId;

  } catch (error) {
    log(`Error creando expediente para ${idCreamos}: ${error}`, "ERROR");
    return null;
  }
}

/**
 * Actualiza expediente cuando cambia el estado
 * @param {string} docId - ID del documento expediente
 * @param {string} nuevoEstado - Nuevo estado
 * @param {string} responsable - Persona que hace el cambio
 * @param {string} notas - Notas del cambio
 */
function actualizarExpedienteEstado(docId, nuevoEstado, responsable, notas = "") {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Buscar sección "SEGUIMIENTO DE ESTADO"
    const texto = body.getText();
    if (texto.includes("SEGUIMIENTO DE ESTADO")) {
      // Insertar nueva línea en historial
      const timestamp = new Date().toLocaleString("es-GT");
      const entrada = `${timestamp} | ${responsable} | Estado: ${nuevoEstado} | ${notas}`;

      body.appendParagraph(entrada);
      doc.saveAndClose();

      log(`Expediente actualizado: ${docId}`, "INFO");
    }

  } catch (error) {
    log(`Error actualizando expediente ${docId}: ${error}`, "ERROR");
  }
}

/**
 * Agrega nota a expediente
 * @param {string} docId - ID del documento
 * @param {string} nota - Contenido de la nota
 * @param {string} responsable - Quién agrega la nota
 */
function agregarNotaExpediente(docId, nota, responsable) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    const timestamp = new Date().toLocaleString("es-GT");
    const entrada = `${timestamp} | ${responsable} | ${nota}`;

    // Buscar sección de notas
    const parrafos = body.getParagraphs();
    let encontrado = false;

    for (let i = 0; i < parrafos.length; i++) {
      if (parrafos[i].getText().includes("NOTAS Y OBSERVACIONES")) {
        body.insertParagraph(i + 2, entrada);
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      body.appendParagraph(entrada);
    }

    doc.saveAndClose();
    log(`Nota agregada a expediente: ${docId}`, "INFO");

  } catch (error) {
    log(`Error agregando nota a expediente ${docId}: ${error}`, "ERROR");
  }
}

/**
 * Función principal que se ejecuta al crear nuevo participante
 * (Llamada desde importar_kobo.gs)
 */
function crearCarpetaYExpediente(idCreamos, nombreCompleto, datosParticipante = {}) {
  try {
    // 1. Crear carpeta
    const carpetaId = crearCarpetaParticipante(idCreamos, nombreCompleto);
    if (!carpetaId) return;

    // 2. Crear expediente
    const expedienteId = crearExpedienteDigital(idCreamos, nombreCompleto, carpetaId, datosParticipante);
    if (!expedienteId) return;

    // 3. Actualizar Sheets con IDs
    actualizarIdsEnSheets(idCreamos, carpetaId, expedienteId);

    log(`Documentos listos para ${idCreamos}`, "INFO");

  } catch (error) {
    log(`Error en flujo crear carpeta/expediente: ${error}`, "ERROR");
  }
}

/**
 * Actualiza las columnas de ID en Sheets
 */
function actualizarIdsEnSheets(idCreamos, carpetaId, expedienteId) {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    for (let i = 1; i < valores.length; i++) {
      if (valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1] === idCreamos) {
        hoja.getRange(i + 1, CONFIG.COLUMNAS.CARPETA_DRIVE_ID).setValue(carpetaId);
        hoja.getRange(i + 1, CONFIG.COLUMNAS.EXPEDIENTE_ID).setValue(expedienteId);
        log(`IDs actualizados en Sheets para ${idCreamos}`, "INFO");
        return;
      }
    }

  } catch (error) {
    log(`Error actualizando IDs en Sheets: ${error}`, "ERROR");
  }
}
/**
 * SCRIPT: Alertas de Mentoría
 * Función: Monitorea participantes en mentoría y genera alertas si exceden duración/sesiones
 * Trigger: Diario 9 AM
 */

/**
 * Función principal - Revisa alertas de mentoría
 */
function revisarAlertasMentoria() {
  log("=== REVISIÓN DE ALERTAS DE MENTORÍA ===", "INFO");

  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    let alertasEncontradas = [];
    const ahora = new Date();

    for (let i = 1; i < valores.length; i++) {
      const fila = valores[i];
      const estado = fila[CONFIG.COLUMNAS.ESTADO_ACTUAL - 1];

      // Solo revisar participantes en Mentoría
      if (estado !== "Mentoría") continue;

      const idCreamos = fila[CONFIG.COLUMNAS.ID_CREAMOS - 1];
      const nombreCompleto = fila[CONFIG.COLUMNAS.NOMBRE_COMPLETO - 1];
      const inicioMentoria = new Date(fila[CONFIG.COLUMNAS.INICIO_MENTORIA - 1]);
      const sesionesCompletadas = parseInt(fila[CONFIG.COLUMNAS.SESIONES_COMPLETADAS - 1]) || 0;
      const responsable = fila[CONFIG.COLUMNAS.RESPONSABLE_ACTUAL - 1];

      // Calcular duración en meses
      const duracionMeses = calcularMeses(inicioMentoria, ahora);

      // Verificar alertas
      if (duracionMeses >= CONFIG.ALERTAS.DURACION_MENTORIA_MESES_MAX ||
          sesionesCompletadas >= CONFIG.ALERTAS.SESIONES_MENTORIA_MAX) {

        const motivo = [];
        if (duracionMeses >= CONFIG.ALERTAS.DURACION_MENTORIA_MESES_MAX) {
          motivo.push(`Duración: ${duracionMeses} meses (máx: ${CONFIG.ALERTAS.DURACION_MENTORIA_MESES_MAX})`);
        }
        if (sesionesCompletadas >= CONFIG.ALERTAS.SESIONES_MENTORIA_MAX) {
          motivo.push(`Sesiones: ${sesionesCompletadas} (máx: ${CONFIG.ALERTAS.SESIONES_MENTORIA_MAX})`);
        }

        alertasEncontradas.push({
          fila: i + 1,
          idCreamos,
          nombreCompleto,
          responsable,
          duracionMeses,
          sesionesCompletadas,
          motivo: motivo.join(" + "),
          inicioMentoria: inicioMentoria.toLocaleDateString("es-GT"),
          expedienteId: fila[CONFIG.COLUMNAS.EXPEDIENTE_ID - 1]
        });
      }
    }

    log(`Alertas detectadas: ${alertasEncontradas.length}`, "INFO");

    // Procesar cada alerta
    for (const alerta of alertasEncontradas) {
      procesarAlertaMentoria(alerta, hoja);
    }

    // Enviar resumen
    if (alertasEncontradas.length > 0) {
      enviarResumenAlertas(alertasEncontradas);
    }

  } catch (error) {
    log(`Error en revisión de alertas: ${error}`, "ERROR");
    enviarAlertaError("Error en revisión de alertas de mentoría", error);
  }
}

/**
 * Procesa una alerta individual
 */
function procesarAlertaMentoria(alerta, hoja) {
  try {
    // 1. Cambiar color de celda a rojo
    const fila = alerta.fila;
    const columnaAlerta = CONFIG.COLUMNAS.ALERTA_MENTORIA;

    hoja.getRange(fila, columnaAlerta)
      .setValue("🔴 ALERTA - Revisar")
      .setBackground("#ff6666")
      .setFontColor("#ffffff");

    // 2. Actualizar expediente
    if (alerta.expedienteId) {
      agregarNotaExpediente(
        alerta.expedienteId,
        `ALERTA AUTOMÁTICA: ${alerta.motivo}. Requiere revisión del responsable.`,
        "Sistema"
      );
    }

    // 3. Enviar email al responsable
    enviarEmailAlerta(alerta);

    log(`Alerta procesada: ${alerta.idCreamos}`, "INFO");

  } catch (error) {
    log(`Error procesando alerta ${alerta.idCreamos}: ${error}`, "ERROR");
  }
}

/**
 * Calcula meses entre dos fechas
 */
function calcularMeses(fechaInicio, fechaFin) {
  const inicioDate = new Date(fechaInicio);
  const finDate = new Date(fechaFin);

  const meses = (finDate.getFullYear() - inicioDate.getFullYear()) * 12 +
    (finDate.getMonth() - inicioDate.getMonth());

  return Math.floor(meses);
}

/**
 * Envía email de alerta al responsable
 */
function enviarEmailAlerta(alerta) {
  try {
    const asunto = `🔴 ALERTA: ${alerta.idCreamos} - ${alerta.nombreCompleto}`;

    const cuerpo = `
      ALERTA DE MENTORÍA

      Participante: ${alerta.nombreCompleto}
      ID: ${alerta.idCreamos}
      Responsable: ${alerta.responsable}

      CAUSA DE ALERTA:
      ${alerta.motivo}

      Fecha Inicio Mentoría: ${alerta.inicioMentoria}
      Duración Actual: ${alerta.duracionMeses} meses
      Sesiones Completadas: ${alerta.sesionesCompletadas}

      ACCIÓN REQUERIDA:
      - Revisar el progreso del participante
      - Evaluar si continúa la mentoría o cambiar de estado
      - Si continúa: resetear contadores o documentar continuidad
      - Si finaliza: marcar estado como "Cierre" o "Derivación"

      Sistema Paso a Paso - ${new Date().toLocaleString("es-GT")}
    `;

    // Enviar al responsable
    if (alerta.responsable && alerta.responsable.includes("@")) {
      GmailApp.sendEmail(alerta.responsable, asunto, cuerpo);
    }

    // Copiar a admin
    GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, asunto, cuerpo);

    log(`Email de alerta enviado a ${alerta.responsable}`, "INFO");

  } catch (error) {
    log(`Error enviando email: ${error}`, "ERROR");
  }
}

/**
 * Envía resumen de alertas
 */
function enviarResumenAlertas(alertas) {
  try {
    let contenidoTabla = "PARTICIPANTE | DURACIÓN | SESIONES | MOTIVO\n";
    contenidoTabla += "---|---|---|---\n";

    for (const alerta of alertas) {
      contenidoTabla += `${alerta.idCreamos} ${alerta.nombreCompleto} | ${alerta.duracionMeses}m | ${alerta.sesionesCompletadas} | ${alerta.motivo}\n`;
    }

    const asunto = `[Paso a Paso] Resumen de Alertas - ${alertas.length} participantes requieren revisión`;

    const cuerpo = `
      RESUMEN DE ALERTAS DE MENTORÍA
      Fecha: ${new Date().toLocaleString("es-GT")}
      Total: ${alertas.length} alertas

      ${contenidoTabla}

      Accede al sistema para revisar y tomar acciones.

      Sistema Paso a Paso
    `;

    GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, asunto, cuerpo);

    log(`Resumen de alertas enviado: ${alertas.length} participantes`, "INFO");

  } catch (error) {
    log(`Error enviando resumen: ${error}`, "ERROR");
  }
}

/**
 * Función para resetear alerta manualmente (cuando responsable toma acción)
 */
function resetearAlertaMentoria(idCreamos) {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    for (let i = 1; i < valores.length; i++) {
      if (valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1] === idCreamos) {
        const fila = i + 1;

        // Si continúa mentoría: resetear sesiones y fechas
        hoja.getRange(fila, CONFIG.COLUMNAS.SESIONES_COMPLETADAS).setValue(0);
        hoja.getRange(fila, CONFIG.COLUMNAS.INICIO_MENTORIA).setValue(new Date());
        hoja.getRange(fila, CONFIG.COLUMNAS.ALERTA_MENTORIA)
          .setValue("🟢 En progreso")
          .setBackground("#ccffcc")
          .setFontColor("#000000");

        log(`Alerta reseteada para ${idCreamos}`, "INFO");
        return;
      }
    }

  } catch (error) {
    log(`Error reseteando alerta: ${error}`, "ERROR");
  }
}

/**
 * Obtiene listado de participantes en alerta
 */
function obtenerParticipantesEnAlerta() {
  const hoja = getHojaMatriz();
  const valores = hoja.getDataRange().getValues();

  const enAlerta = [];

  for (let i = 1; i < valores.length; i++) {
    const fila = valores[i];
    const alerta = fila[CONFIG.COLUMNAS.ALERTA_MENTORIA - 1];

    if (alerta && alerta.includes("ALERTA")) {
      enAlerta.push({
        id: fila[CONFIG.COLUMNAS.ID_CREAMOS - 1],
        nombre: fila[CONFIG.COLUMNAS.NOMBRE_COMPLETO - 1],
        responsable: fila[CONFIG.COLUMNAS.RESPONSABLE_ACTUAL - 1]
      });
    }
  }

  return enAlerta;
}
/**
 * SCRIPT: Gestión de Estados
 * Función: Maneja transiciones de estado de participantes
 * Triggers: onEdit (cambio en columna Estado)
 */

/**
 * Evento onEdit - Se ejecuta cuando hay cambios en la hoja
 * Detecta cambios en la columna Estado y procesa transiciones
 */
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;

    // Verificar si es la hoja maestra y columna correcta
    if (sheet.getName() !== CONFIG.SHEET_NAME) return;
    if (range.getColumn() !== CONFIG.COLUMNAS.ESTADO_ACTUAL) return;

    const fila = range.getRow();
    const nuevoEstado = range.getValue();
    const valores = sheet.getDataRange().getValues();
    const datosParticipante = valores[fila - 1];

    // Obtener datos relevantes
    const idCreamos = datosParticipante[CONFIG.COLUMNAS.ID_CREAMOS - 1];
    const nombreCompleto = datosParticipante[CONFIG.COLUMNAS.NOMBRE_COMPLETO - 1];
    const responsable = Session.getActiveUser().getEmail();

    log(`Cambio de estado detectado: ${idCreamos} → ${nuevoEstado}`, "INFO");

    // Procesar transición
    procesarTransicionEstado(
      sheet,
      fila,
      nuevoEstado,
      datosParticipante,
      responsable
    );

  } catch (error) {
    log(`Error en onEdit: ${error}`, "ERROR");
  }
}

/**
 * Procesa la transición de estado
 */
function procesarTransicionEstado(sheet, fila, nuevoEstado, datosParticipante, responsable) {
  try {
    const idCreamos = datosParticipante[CONFIG.COLUMNAS.ID_CREAMOS - 1];
    const nombreCompleto = datosParticipante[CONFIG.COLUMNAS.NOMBRE_COMPLETO - 1];
    const estadoAnterior = datosParticipante[CONFIG.COLUMNAS.ESTADO_ACTUAL - 1];
    const expedienteId = datosParticipante[CONFIG.COLUMNAS.EXPEDIENTE_ID - 1];

    // 1. Registrar fecha del cambio
    sheet.getRange(fila, CONFIG.COLUMNAS.FECHA_ULTIMO_CAMBIO)
      .setValue(new Date());

    // 2. Registrar responsable si no está
    if (!datosParticipante[CONFIG.COLUMNAS.RESPONSABLE_ACTUAL - 1]) {
      sheet.getRange(fila, CONFIG.COLUMNAS.RESPONSABLE_ACTUAL)
        .setValue(responsable);
    }

    // 3. Lógica específica por estado
    if (nuevoEstado === "Mentoría") {
      iniciarMentoria(sheet, fila, datosParticipante);
    } else if (nuevoEstado === "Formación") {
      iniciarFormacion(sheet, fila, datosParticipante);
    } else if (nuevoEstado === "Cierre") {
      procesarCierre(sheet, fila, datosParticipante);
    } else if (nuevoEstado === "Derivación") {
      marcarComoDerivacion(sheet, fila, datosParticipante);
    }

    // 4. Actualizar expediente
    if (expedienteId) {
      const notas = `Cambio de estado: ${estadoAnterior} → ${nuevoEstado}`;
      actualizarExpedienteEstado(expedienteId, nuevoEstado, responsable, notas);
    }

    // 5. Enviar notificación
    notificarCambioEstado(idCreamos, nombreCompleto, estadoAnterior, nuevoEstado, responsable);

    log(`Transición procesada: ${idCreamos} (${estadoAnterior} → ${nuevoEstado})`, "INFO");

  } catch (error) {
    log(`Error procesando transición: ${error}`, "ERROR");
  }
}

/**
 * Inicia seguimiento de mentoría
 */
function iniciarMentoria(sheet, fila, datosParticipante) {
  try {
    const ahora = new Date();

    // Registrar fecha de inicio
    sheet.getRange(fila, CONFIG.COLUMNAS.INICIO_MENTORIA)
      .setValue(ahora);

    // Resetear sesiones (en caso de cambio de estado)
    sheet.getRange(fila, CONFIG.COLUMNAS.SESIONES_COMPLETADAS)
      .setValue(0);

    // Cambiar indicador a verde
    sheet.getRange(fila, CONFIG.COLUMNAS.ALERTA_MENTORIA)
      .setValue("🟢 En progreso")
      .setBackground("#ccffcc");

    log(`Mentoría iniciada en fila ${fila}`, "INFO");

  } catch (error) {
    log(`Error iniciando mentoría: ${error}`, "ERROR");
  }
}

/**
 * Inicia seguimiento de formación
 */
function iniciarFormacion(sheet, fila, datosParticipante) {
  try {
    const ahora = new Date();

    // Registrar que entra a formación
    sheet.getRange(fila, CONFIG.COLUMNAS.FECHA_ULTIMO_CAMBIO)
      .setValue(ahora);

    log(`Formación iniciada para ${datosParticipante[CONFIG.COLUMNAS.ID_CREAMOS - 1]}`, "INFO");

  } catch (error) {
    log(`Error iniciando formación: ${error}`, "ERROR");
  }
}

/**
 * Procesa cierre de participante
 */
function procesarCierre(sheet, fila, datosParticipante) {
  try {
    const ahora = new Date();
    const idCreamos = datosParticipante[CONFIG.COLUMNAS.ID_CREAMOS - 1];

    // Registrar fecha de cierre
    sheet.getRange(fila, CONFIG.COLUMNAS.FECHA_CIERRE)
      .setValue(ahora);

    // Solicitar motivo (si no está lleno)
    if (!datosParticipante[CONFIG.COLUMNAS.TIPO_CIERRE - 1]) {
      sheet.getRange(fila, CONFIG.COLUMNAS.TIPO_CIERRE)
        .setValue("Pendiente clasificación");
    }

    log(`Cierre registrado para ${idCreamos}`, "INFO");

  } catch (error) {
    log(`Error procesando cierre: ${error}`, "ERROR");
  }
}

/**
 * Marca como derivación
 */
function marcarComoDerivacion(sheet, fila, datosParticipante) {
  try {
    const ahora = new Date();

    sheet.getRange(fila, CONFIG.COLUMNAS.FECHA_DERIVACION || (CONFIG.COLUMNAS.FECHA_CIERRE + 1))
      .setValue(ahora);

    // Marcar estado como "Pendiente" hasta que sea aceptada
    sheet.getRange(fila, (CONFIG.COLUMNAS.ESTADO_DERIVACION || (CONFIG.COLUMNAS.FECHA_CIERRE + 2)))
      .setValue("Pendiente");

    log(`Derivación marcada para fila ${fila}`, "INFO");

  } catch (error) {
    log(`Error marcando derivación: ${error}`, "ERROR");
  }
}

/**
 * Envía notificación de cambio de estado
 */
function notificarCambioEstado(idCreamos, nombreCompleto, estadoAnterior, nuevoEstado, responsable) {
  try {
    const asunto = `[Paso a Paso] Cambio de Estado: ${idCreamos}`;

    const cuerpo = `
      CAMBIO DE ESTADO DE PARTICIPANTE

      Participante: ${nombreCompleto}
      ID: ${idCreamos}

      Estado Anterior: ${estadoAnterior}
      Estado Nuevo: ${nuevoEstado}
      Responsable: ${responsable}
      Fecha: ${new Date().toLocaleString("es-GT")}

      El expediente digital ha sido actualizado automáticamente.

      Sistema Paso a Paso
    `;

    // Enviar a admin y responsable
    GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, asunto, cuerpo);
    if (responsable && responsable.includes("@")) {
      GmailApp.sendEmail(responsable, asunto, cuerpo);
    }

  } catch (error) {
    log(`Error enviando notificación: ${error}`, "ERROR");
  }
}

/**
 * Permite cambiar estado manualmente desde script
 * Útil para automatizaciones
 */
function cambiarEstado(idCreamos, nuevoEstado, responsable = "") {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    for (let i = 1; i < valores.length; i++) {
      if (valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1] === idCreamos) {
        const fila = i + 1;

        hoja.getRange(fila, CONFIG.COLUMNAS.ESTADO_ACTUAL)
          .setValue(nuevoEstado);

        if (responsable) {
          hoja.getRange(fila, CONFIG.COLUMNAS.RESPONSABLE_ACTUAL)
            .setValue(responsable);
        }

        log(`Estado cambiado: ${idCreamos} → ${nuevoEstado}`, "INFO");
        return true;
      }
    }

    log(`Participante no encontrado: ${idCreamos}`, "WARN");
    return false;

  } catch (error) {
    log(`Error cambiando estado: ${error}`, "ERROR");
    return false;
  }
}

/**
 * Obtiene historial de transiciones de un participante
 */
function obtenerHistorialEstados(idCreamos) {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    for (let i = 1; i < valores.length; i++) {
      if (valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1] === idCreamos) {
        // Retornar estado actual y fecha
        return {
          estadoActual: valores[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1],
          fechaUltimoCambio: valores[i][CONFIG.COLUMNAS.FECHA_ULTIMO_CAMBIO - 1],
          responsable: valores[i][CONFIG.COLUMNAS.RESPONSABLE_ACTUAL - 1]
        };
      }
    }

    return null;

  } catch (error) {
    log(`Error obteniendo historial: ${error}`, "ERROR");
    return null;
  }
}
/**
 * SCRIPT: Utilidades y Funciones Helper
 * Función: Funciones reutilizables para el sistema
 */

/**
 * Obtiene datos de un participante por ID
 */
function obtenerParticipante(idCreamos) {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    for (let i = 1; i < valores.length; i++) {
      if (valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1] === idCreamos) {
        return {
          fila: i + 1,
          id: valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1],
          nombre: valores[i][CONFIG.COLUMNAS.NOMBRE_COMPLETO - 1],
          dpi: valores[i][CONFIG.COLUMNAS.DPI - 1],
          edad: valores[i][CONFIG.COLUMNAS.EDAD - 1],
          genero: valores[i][CONFIG.COLUMNAS.GENERO - 1],
          telefono: valores[i][CONFIG.COLUMNAS.TELEFONO - 1],
          zona: valores[i][CONFIG.COLUMNAS.ZONA - 1],
          nivel_educativo: valores[i][CONFIG.COLUMNAS.NIVEL_EDUCATIVO - 1],
          problemas_vitales: valores[i][CONFIG.COLUMNAS.PROBLEMAS_VITALES - 1],
          conciencia: valores[i][CONFIG.COLUMNAS.CONCIENCIA_PARTICIPANTE - 1],
          perfil: valores[i][CONFIG.COLUMNAS.PERFIL_RESULTANTE - 1],
          estado: valores[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1],
          responsable: valores[i][CONFIG.COLUMNAS.RESPONSABLE_ACTUAL - 1],
          expediente_id: valores[i][CONFIG.COLUMNAS.EXPEDIENTE_ID - 1],
          carpeta_id: valores[i][CONFIG.COLUMNAS.CARPETA_DRIVE_ID - 1]
        };
      }
    }

    return null;

  } catch (error) {
    log(`Error obteniendo participante: ${error}`, "ERROR");
    return null;
  }
}

/**
 * Obtiene lista de participantes filtrando por estado
 */
function obtenerParticipantesPorEstado(estado) {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    const resultado = [];

    for (let i = 1; i < valores.length; i++) {
      if (valores[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1] === estado) {
        resultado.push({
          id: valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1],
          nombre: valores[i][CONFIG.COLUMNAS.NOMBRE_COMPLETO - 1],
          responsable: valores[i][CONFIG.COLUMNAS.RESPONSABLE_ACTUAL - 1],
          fecha_cambio: valores[i][CONFIG.COLUMNAS.FECHA_ULTIMO_CAMBIO - 1]
        });
      }
    }

    return resultado;

  } catch (error) {
    log(`Error obteniendo participantes por estado: ${error}`, "ERROR");
    return [];
  }
}

/**
 * Obtiene participantes por responsable
 */
function obtenerParticipantesPorResponsable(email) {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    const resultado = [];

    for (let i = 1; i < valores.length; i++) {
      if (valores[i][CONFIG.COLUMNAS.RESPONSABLE_ACTUAL - 1] === email) {
        resultado.push({
          id: valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1],
          nombre: valores[i][CONFIG.COLUMNAS.NOMBRE_COMPLETO - 1],
          estado: valores[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1],
          alerta: valores[i][CONFIG.COLUMNAS.ALERTA_MENTORIA - 1]
        });
      }
    }

    return resultado;

  } catch (error) {
    log(`Error obteniendo participantes por responsable: ${error}`, "ERROR");
    return [];
  }
}

/**
 * Registra una sesión de mentoría
 */
function registrarSesionMentoria(idCreamos, fecha = new Date(), notas = "") {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    for (let i = 1; i < valores.length; i++) {
      if (valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1] === idCreamos) {
        const fila = i + 1;

        // Incrementar sesiones
        const sesionesActuales = parseInt(valores[i][CONFIG.COLUMNAS.SESIONES_COMPLETADAS - 1]) || 0;
        hoja.getRange(fila, CONFIG.COLUMNAS.SESIONES_COMPLETADAS)
          .setValue(sesionesActuales + 1);

        // Actualizar última sesión
        hoja.getRange(fila, CONFIG.COLUMNAS.ULTIMA_SESION)
          .setValue(fecha);

        // Agregar nota a expediente si existe
        const expedienteId = valores[i][CONFIG.COLUMNAS.EXPEDIENTE_ID - 1];
        if (expedienteId) {
          agregarNotaExpediente(
            expedienteId,
            `Sesión de mentoría: ${notas}`,
            "Sistema"
          );
        }

        log(`Sesión registrada para ${idCreamos}`, "INFO");
        return true;
      }
    }

    return false;

  } catch (error) {
    log(`Error registrando sesión: ${error}`, "ERROR");
    return false;
  }
}

/**
 * Crea una derivación a otro programa
 */
function crearDerivacion(idCreamos, programaDestino, motivo, responsable) {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    for (let i = 1; i < valores.length; i++) {
      if (valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1] === idCreamos) {
        const fila = i + 1;

        // Marcar como derivación interna
        hoja.getRange(fila, CONFIG.COLUMNAS.TIPO_DERIVACION)
          .setValue("Interna");

        // Indicar programa destino
        hoja.getRange(fila, CONFIG.COLUMNAS.DERIVADO_A)
          .setValue(programaDestino);

        // Cambiar estado
        hoja.getRange(fila, CONFIG.COLUMNAS.ESTADO_ACTUAL)
          .setValue("Derivación");

        // Agregar nota
        const expedienteId = valores[i][CONFIG.COLUMNAS.EXPEDIENTE_ID - 1];
        if (expedienteId) {
          agregarNotaExpediente(
            expedienteId,
            `DERIVACIÓN INTERNA: ${programaDestino}. Motivo: ${motivo}`,
            responsable
          );
        }

        log(`Derivación creada: ${idCreamos} → ${programaDestino}`, "INFO");
        return true;
      }
    }

    return false;

  } catch (error) {
    log(`Error creando derivación: ${error}`, "ERROR");
    return false;
  }
}

/**
 * Completa cierre de participante
 */
function completarCierre(idCreamos, tipoCierre, resultadoFinal, responsable) {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    for (let i = 1; i < valores.length; i++) {
      if (valores[i][CONFIG.COLUMNAS.ID_CREAMOS - 1] === idCreamos) {
        const fila = i + 1;
        const ahora = new Date();

        // Registrar tipo de cierre
        hoja.getRange(fila, CONFIG.COLUMNAS.TIPO_CIERRE)
          .setValue(tipoCierre);

        // Registrar fecha
        hoja.getRange(fila, CONFIG.COLUMNAS.FECHA_CIERRE)
          .setValue(ahora);

        // Registrar resultado
        hoja.getRange(fila, CONFIG.COLUMNAS.RESULTADO_FINAL)
          .setValue(resultadoFinal);

        // Registrar encargado
        hoja.getRange(fila, CONFIG.COLUMNAS.ENCARGADO_CIERRE)
          .setValue(responsable);

        // Cambiar estado a Cierre (final)
        hoja.getRange(fila, CONFIG.COLUMNAS.ESTADO_ACTUAL)
          .setValue("Cierre");

        // Actualizar expediente
        const expedienteId = valores[i][CONFIG.COLUMNAS.EXPEDIENTE_ID - 1];
        if (expedienteId) {
          agregarNotaExpediente(
            expedienteId,
            `CIERRE: Tipo=${tipoCierre}, Resultado=${resultadoFinal}`,
            responsable
          );
        }

        log(`Cierre completado: ${idCreamos}`, "INFO");
        return true;
      }
    }

    return false;

  } catch (error) {
    log(`Error completando cierre: ${error}`, "ERROR");
    return false;
  }
}

/**
 * Genera reporte de participantes activos
 */
function generarReporteActivos() {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    const reporte = {
      total: 0,
      por_estado: {},
      por_responsable: {}
    };

    for (let i = 1; i < valores.length; i++) {
      const estado = valores[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1];
      const responsable = valores[i][CONFIG.COLUMNAS.RESPONSABLE_ACTUAL - 1];

      if (estado && estado !== "Cierre") {
        reporte.total++;

        reporte.por_estado[estado] = (reporte.por_estado[estado] || 0) + 1;
        reporte.por_responsable[responsable] = (reporte.por_responsable[responsable] || 0) + 1;
      }
    }

    return reporte;

  } catch (error) {
    log(`Error generando reporte: ${error}`, "ERROR");
    return null;
  }
}

/**
 * Sincroniza información de participante con expediente
 */
function sincronizarExpediente(idCreamos) {
  try {
    const participante = obtenerParticipante(idCreamos);
    if (!participante) {
      log(`Participante no encontrado: ${idCreamos}`, "WARN");
      return false;
    }

    if (!participante.expediente_id) {
      log(`Sin expediente para ${idCreamos}`, "WARN");
      return false;
    }

    const doc = DocumentApp.openById(participante.expediente_id);
    const body = doc.getBody();

    // Actualizar datos básicos en el expediente
    const texto = body.getText();

    // (Implementar búsqueda y reemplazo de datos en el doc)
    // Por ahora, solo registrar acción

    log(`Expediente sincronizado: ${idCreamos}`, "INFO");
    return true;

  } catch (error) {
    log(`Error sincronizando expediente: ${error}`, "ERROR");
    return false;
  }
}

/**
 * Obtiene estadísticas del sistema
 */
function obtenerEstadisticas() {
  try {
    const hoja = getHojaMatriz();
    const valores = hoja.getDataRange().getValues();

    const stats = {
      total_participantes: valores.length - 1,
      en_orientacion: 0,
      en_mentoria: 0,
      en_formacion: 0,
      en_derivacion: 0,
      cerrados: 0,
      alertas_activas: 0,
      fecha_actualizacion: new Date().toLocaleString("es-GT")
    };

    for (let i = 1; i < valores.length; i++) {
      const estado = valores[i][CONFIG.COLUMNAS.ESTADO_ACTUAL - 1];
      const alerta = valores[i][CONFIG.COLUMNAS.ALERTA_MENTORIA - 1];

      if (estado === "Orientación") stats.en_orientacion++;
      else if (estado === "Mentoría") stats.en_mentoria++;
      else if (estado === "Formación") stats.en_formacion++;
      else if (estado === "Derivación") stats.en_derivacion++;
      else if (estado === "Cierre") stats.cerrados++;

      if (alerta && alerta.includes("ALERTA")) stats.alertas_activas++;
    }

    return stats;

  } catch (error) {
    log(`Error obteniendo estadísticas: ${error}`, "ERROR");
    return null;
  }
}

/**
 * Envía email con estadísticas diarias
 */
function enviarEstadisticasDiarias() {
  try {
    const stats = obtenerEstadisticas();
    if (!stats) return;

    const asunto = "[Paso a Paso] Estadísticas Diarias";

    const cuerpo = `
      ESTADÍSTICAS DEL SISTEMA - ${stats.fecha_actualizacion}

      Total de Participantes: ${stats.total_participantes}

      Por Estado:
      - En Orientación: ${stats.en_orientacion}
      - En Mentoría: ${stats.en_mentoria}
      - En Formación: ${stats.en_formacion}
      - En Derivación: ${stats.en_derivacion}
      - Cerrados: ${stats.cerrados}

      ⚠️ Alertas Activas: ${stats.alertas_activas}

      Accede al sistema para más detalles.

      Sistema Paso a Paso
    `;

    GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, asunto, cuerpo);
    log("Estadísticas enviadas", "INFO");

  } catch (error) {
    log(`Error enviando estadísticas: ${error}`, "ERROR");
  }
}
