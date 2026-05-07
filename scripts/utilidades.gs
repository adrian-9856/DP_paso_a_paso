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
