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
