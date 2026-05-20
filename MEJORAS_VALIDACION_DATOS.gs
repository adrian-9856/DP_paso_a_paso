// ============================================================================
// MEJORAS DE VALIDACIÓN Y DIAGNÓSTICO DE DATOS — v1.0
// Manejo inteligente de sincronización Kobo con validaciones
// ============================================================================

/**
 * MEJORA 1: Inspeccionar dinámicamente campos disponibles en Kobo
 * Muestra exactamente qué campos tiene tu formulario Kobo
 */
function diagnosticoCamposKobo() {
  try {
    const key = getKoboKey();
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json&limit=1',
      { headers: { Authorization: 'Token ' + key }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert('❌ Error al conectar Kobo: ' + resp.getResponseCode());
      return;
    }

    const data = JSON.parse(resp.getContentText());
    const registros = data.results || [];

    if (!registros.length) {
      SpreadsheetApp.getUi().alert('⚠️ No hay respuestas en Kobo todavía');
      return;
    }

    const primerRegistro = registros[0];
    const camposDisponibles = Object.keys(primerRegistro).sort();

    // Mapeo sugerido de campos críticos
    const CAMPOS_CRITICOS = ['nombre', 'dpi', 'telefono', 'edad', 'genero', 'email'];
    const mapeoCriticos = {};

    CAMPOS_CRITICOS.forEach(c => {
      const coincidencia = camposDisponibles.find(k =>
        k.toLowerCase().includes(c) && !k.startsWith('_')
      );
      mapeoCriticos[c] = coincidencia || '❌ NO ENCONTRADO';
    });

    // Crear reporte
    let reporte = '📊 CAMPOS DISPONIBLES EN KOBO\n';
    reporte += '================================\n\n';
    reporte += '🔵 CAMPOS CRÍTICOS (requeridos):\n';
    Object.entries(mapeoCriticos).forEach(([campo, koboField]) => {
      reporte += `  ${campo.toUpperCase().padEnd(15)} → ${koboField}\n`;
    });

    reporte += '\n🟢 TODOS LOS CAMPOS (' + camposDisponibles.length + '):\n';
    camposDisponibles.forEach(c => {
      if (!c.startsWith('_')) {
        reporte += `  • ${c}\n`;
      }
    });

    reporte += '\n📝 INSTRUCCIONES:\n';
    reporte += '1. Copia los nombres exactos de los campos de la lista arriba\n';
    reporte += '2. Ve a ⚙️ CONFIGURACIÓN del Sistema\n';
    reporte += '3. En la sección MAPEO KOBO, actualiza los nombres\n';

    const ui = SpreadsheetApp.getUi();
    const resultado = ui.alert(reporte, ui.ButtonSet.OK_CANCEL);

    if (resultado === ui.Button.OK) {
      guardarDiagnosticoEnLog(reporte, camposDisponibles);
      SpreadsheetApp.getUi().alert('✅ Diagnóstico guardado en la hoja "Log"');
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('diagnosticosCamposKobo', e);
  }
}

// ============================================================================
// MEJORA 2: Validación de Datos Requeridos
// ============================================================================

/**
 * Valida que un registro tenga los campos mínimos requeridos
 * @param {Object} registro - Datos del participante desde Kobo
 * @returns {Object} { valido: boolean, errores: [string] }
 */
function validarDatosRequeridos(registro) {
  const errores = [];
  const advertencias = [];

  // CAMPOS OBLIGATORIOS
  const nombre   = registro['nombre_completo'] || '';
  const dpi      = registro['numero_dpi']      || '';
  const telefono = registro['telefono']        || '';

  if (!nombre || nombre.trim() === '') {
    errores.push('❌ FALTA: Nombre completo');
  }

  if (!dpi || dpi.trim() === '') {
    errores.push('❌ FALTA: DPI');
  }

  if (!telefono || telefono.trim() === '') {
    errores.push('❌ FALTA: Teléfono');
  }

  // CAMPOS IMPORTANTES (advertencias)
  const edad     = registro['edad']                  || '';
  const email    = registro['email']                 || '';
  const educacion = registro['ultimo_grado_aprobado'] || '';

  if (!edad) advertencias.push('⚠️  Sin edad');
  if (!email) advertencias.push('⚠️  Sin email');
  if (!educacion) advertencias.push('⚠️  Sin educación');

  return {
    valido: errores.length === 0,
    errores: errores,
    advertencias: advertencias,
    campos: { nombre, dpi, telefono, edad, email, educacion }
  };
}

/**
 * Obtiene estadísticas de calidad de datos de la última sincronización
 */
function analizarCalidadDatos() {
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);

    if (!hoja || hoja.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('⚠️ No hay datos para analizar');
      return;
    }

    const M = CONFIG.COL;
    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 10).getValues();

    let estadisticas = {
      total: datos.length,
      conNombre: 0,
      conDPI: 0,
      conTelefono: 0,
      conEmail: 0,
      completos: 0,
      incompletos: []
    };

    datos.forEach((fila, idx) => {
      const id = fila[M.ID - 1];
      const nombre = fila[M.NOMBRE - 1];
      const dpi = fila[M.DPI - 1];
      const telefono = fila[M.TELEFONO - 1];
      const email = fila[M.EMAIL - 1];

      if (nombre) estadisticas.conNombre++;
      if (dpi) estadisticas.conDPI++;
      if (telefono) estadisticas.conTelefono++;
      if (email) estadisticas.conEmail++;

      if (nombre && dpi && telefono) {
        estadisticas.completos++;
      } else {
        estadisticas.incompletos.push({
          id: id,
          nombre: nombre,
          fila: idx + 2,
          falta: [!nombre ? 'Nombre' : '', !dpi ? 'DPI' : '', !telefono ? 'Teléfono' : '']
            .filter(x => x).join(', ')
        });
      }
    });

    mostrarPanelDiagnostico(estadisticas);

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('analizarCalidadDatos', e);
  }
}

/**
 * Muestra panel visual del diagnóstico
 */
function mostrarPanelDiagnostico(stats) {
  const porcNombre = ((stats.conNombre / stats.total) * 100).toFixed(1);
  const porcDPI = ((stats.conDPI / stats.total) * 100).toFixed(1);
  const porcTelefono = ((stats.conTelefono / stats.total) * 100).toFixed(1);
  const porcEmail = ((stats.conEmail / stats.total) * 100).toFixed(1);
  const porcCompletos = ((stats.completos / stats.total) * 100).toFixed(1);

  let panel = '📊 DIAGNÓSTICO DE CALIDAD DE DATOS\n';
  panel += '═════════════════════════════════════════\n\n';
  panel += `Total de participantes: ${stats.total}\n\n`;

  panel += '✅ CAMPOS COMPLETADOS:\n';
  panel += `  Nombre ............ ${porcNombre}% (${stats.conNombre}/${stats.total})\n`;
  panel += `  DPI ............... ${porcDPI}% (${stats.conDPI}/${stats.total})\n`;
  panel += `  Teléfono .......... ${porcTelefono}% (${stats.conTelefono}/${stats.total})\n`;
  panel += `  Email ............ ${porcEmail}% (${stats.conEmail}/${stats.total})\n\n`;

  panel += `🟢 REGISTROS COMPLETOS (datos críticos): ${porcCompletos}% (${stats.completos}/${stats.total})\n\n`;

  if (stats.incompletos.length > 0) {
    panel += `🔴 REGISTROS INCOMPLETOS (${stats.incompletos.length}):\n`;
    stats.incompletos.slice(0, 10).forEach(reg => {
      panel += `  Fila ${reg.fila}: ${reg.nombre || '(sin nombre)'} → Falta: ${reg.falta}\n`;
    });
    if (stats.incompletos.length > 10) {
      panel += `  ... y ${stats.incompletos.length - 10} más\n`;
    }
  }

  panel += '\n💡 RECOMENDACIONES:\n';
  if (porcNombre < 100) panel += '  • Verificar que todos tengan nombre completo\n';
  if (porcDPI < 100) panel += '  • Completar DPI faltantes (crítico para identificación)\n';
  if (porcTelefono < 100) panel += '  • Agregar teléfono para contacto\n';
  if (porcCompletos < 80) panel += '  • ⚠️ ALERTA: Menos del 80% de registros completos\n';

  SpreadsheetApp.getUi().alert(panel);
}

// ============================================================================
// MEJORA 3: Sincronización Mejorada con Validaciones
// ============================================================================

/**
 * Sincroniza Kobo con validaciones automáticas
 * @param {boolean} silencioso - Si es true, no muestra alertas
 */
function sincronizarConValidaciones(silencioso) {
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(CONFIG.HOJA);
    if (!hoja) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ Instala primero.');
      return;
    }

    let apiKey;
    try {
      apiKey = getKoboKey();
    } catch (e) {
      if (!silencioso) SpreadsheetApp.getUi().alert('❌ ' + e.message);
      return;
    }

    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json',
      { headers: { Authorization: 'Token ' + apiKey }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      if (!silencioso) {
        SpreadsheetApp.getUi().alert('❌ Error Kobo ' + resp.getResponseCode() +
          '\n\nVerifica tu API Key en ⚙️ Configuración.');
      }
      return;
    }

    const registros = JSON.parse(resp.getContentText()).results || [];
    if (!registros.length) {
      if (!silencioso) SpreadsheetApp.getUi().alert('⚠️ No hay respuestas en Kobo todavía.');
      return;
    }

    const existentes = new Set();
    if (hoja.getLastRow() > 1) {
      hoja.getRange(2, CONFIG.COL.ID, hoja.getLastRow() - 1, 1)
        .getValues().forEach(r => { if (r[0]) existentes.add(String(r[0])); });
    }

    const folderBase = DriveApp.getFolderById(getFolderId());
    const M = CONFIG.COL;
    let agregados = 0;
    let rechazados = 0;
    let advertencias = [];

    registros.forEach((r, idx) => {
      const id = String(r['creamos_id'] || r['_id'] || '');
      if (!id || existentes.has(id)) return;

      // VALIDAR DATOS
      const validacion = validarDatosRequeridos(r);

      if (!validacion.valido) {
        rechazados++;
        advertencias.push({
          id: id,
          errores: validacion.errores
        });
        logError('sincronizarConValidaciones',
          `Registro rechazado [${id}]: ${validacion.errores.join('; ')}`);
        return;
      }

      // CREAR FILA
      const nombre = r['nombre_completo'] || 'Participante ' + id;
      const fila = new Array(26).fill('');
      fila[M.ID - 1]         = id;
      fila[M.FECHA - 1]      = new Date();
      fila[M.NOMBRE - 1]     = nombre;
      fila[M.DPI - 1]        = r['numero_dpi']                  || '';
      fila[M.EDAD - 1]       = r['edad']                        || '';
      fila[M.GENERO - 1]     = r['genero']                      || '';
      fila[M.TELEFONO - 1]   = r['telefono']                    || '';
      fila[M.ZONA - 1]       = r['lugar_residencia']            || '';
      fila[M.EMAIL - 1]      = r['email']                       || '';
      fila[M.EDUCACION - 1]  = r['ultimo_grado_aprobado']       || '';
      fila[M.LABORAL - 1]    = r['situacion_laboral_actual']    || '';
      fila[M.FORTALEZAS - 1] = r['habilidades_especificas']     || '';
      fila[M.OBJETIVO - 1]   = r['objetivo_laboral_especifico'] || '';
      fila[M.PERFIL - 1] = normalizarPerfil(r['perfil_asignado']);
      fila[M.PRIORIDAD - 1] = normalizarPrioridad(r['prioridad_caso']);
      fila[M.PUNTAJE - 1] = Number(r['puntaje_total_60']) || 0;
      fila[M.DIM1 - 1] = Number(r['dimension_1_capital_educativo']) || 0;
      fila[M.DIM2 - 1] = Number(r['dimension_2_capital_laboral']) || 0;
      fila[M.DIM3 - 1] = Number(r['dimension_3_habilidades_digitales']) || 0;
      fila[M.DIM4 - 1] = Number(r['dimension_4_claridad_vocacional']) || 0;
      fila[M.DIM5 - 1] = Number(r['dimension_5_barreras_estructurales']) || 0;
      fila[M.DIM6 - 1] = Number(r['dimension_6_red_apoyo']) || 0;
      fila[M.ESTADO - 1] = 'Orientación';

      // CREAR EXPEDIENTE
      const exp = crearExpediente(folderBase, id, nombre, fila, M);
      fila[M.CARPETA_ID - 1] = exp.carpetaId;
      fila[M.DOC_ID - 1] = exp.docId;
      fila[M.DOC_URL - 1] = exp.docUrl;

      hoja.appendRow(fila);
      colorearFila(hoja, hoja.getLastRow(), fila[M.PERFIL - 1]);
      registrarDerivacionesAutomaticas(ss, id, nombre, fila, M);
      existentes.add(id);
      agregados++;

      // Registrar advertencias si las hay
      if (validacion.advertencias.length > 0) {
        logError('sincronizarConValidaciones',
          `Advertencia [${id}]: ${validacion.advertencias.join('; ')}`);
      }
    });

    // MOSTRAR REPORTE
    let resumen = `✅ Sincronización completada\n`;
    resumen += `📥 ${agregados} nuevos participantes agregados\n`;
    resumen += `📊 ${registros.length} total disponible en Kobo\n`;

    if (rechazados > 0) {
      resumen += `\n🔴 ${rechazados} rechazados (datos incompletos)\n`;
      resumen += `Revisa la hoja "Log" para detalles.`;
    }

    if (!silencioso) SpreadsheetApp.getUi().alert(resumen);

  } catch (e) {
    logError('sincronizarConValidaciones', e);
    if (!silencioso) SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

// ============================================================================
// HELPERS
// ============================================================================

function guardarDiagnosticoEnLog(reporte, campos) {
  try {
    const ss = SpreadsheetApp.getActive();
    let logHoja = ss.getSheetByName('Log');
    if (!logHoja) {
      logHoja = ss.insertSheet('Log');
      logHoja.appendRow(['Timestamp', 'Tipo', 'Detalles', 'Usuario']);
    }
    logHoja.appendRow([
      new Date(),
      '📊 DIAGNÓSTICO',
      reporte,
      Session.getEffectiveUser().getEmail()
    ]);
  } catch (e) {
    Logger.log('Error guardando diagnóstico: ' + e);
  }
}
