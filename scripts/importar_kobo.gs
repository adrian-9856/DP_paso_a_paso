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
