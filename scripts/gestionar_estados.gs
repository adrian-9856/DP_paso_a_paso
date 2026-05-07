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
