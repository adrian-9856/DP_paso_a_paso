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
