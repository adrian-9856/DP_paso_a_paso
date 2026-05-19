// ============================================================================
// INSPECTOR DE DATOS KOBO — Descubre qué información realmente envía tu formulario
// ============================================================================

/**
 * Lee el primer registro de Kobo y muestra TODOS sus datos
 * Esto te permite ver exactamente qué campos tienen información
 */
function inspeccionarPrimerRegistroKobo() {
  try {
    const key = getKoboKey();
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json&limit=1',
      { headers: { Authorization: 'Token ' + key }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert('❌ Error: ' + resp.getResponseCode());
      return;
    }

    const data = JSON.parse(resp.getContentText());
    const registros = data.results || [];

    if (!registros.length) {
      SpreadsheetApp.getUi().alert('⚠️ No hay respuestas en Kobo');
      return;
    }

    const primer = registros[0];

    // Crear reporte detallado
    let reporte = '📋 PRIMER REGISTRO KOBO — DATOS COMPLETOS\n';
    reporte += '═══════════════════════════════════════════════\n\n';

    // Separar campos por si tienen datos
    const conDatos = {};
    const sinDatos = [];

    Object.entries(primer).forEach(([clave, valor]) => {
      if (clave.startsWith('_')) return; // Ignorar meta campos

      if (valor && valor.toString().trim() !== '') {
        conDatos[clave] = valor;
      } else {
        sinDatos.push(clave);
      }
    });

    reporte += '✅ CAMPOS CON DATOS (' + Object.keys(conDatos).length + '):\n';
    reporte += '───────────────────────────────────────────\n';
    Object.entries(conDatos).forEach(([clave, valor]) => {
      const valorTxt = String(valor).substring(0, 50);
      reporte += `• ${clave}\n  → ${valorTxt}\n`;
    });

    reporte += '\n❌ CAMPOS VACÍOS (' + sinDatos.length + '):\n';
    reporte += '───────────────────────────────────────────\n';
    sinDatos.slice(0, 15).forEach(c => {
      reporte += `• ${c}\n`;
    });
    if (sinDatos.length > 15) {
      reporte += `• ... y ${sinDatos.length - 15} más\n`;
    }

    reporte += '\n📊 CAMPOS CRÍTICOS PARA MAPEAR:\n';
    reporte += '───────────────────────────────────────────\n';

    // Buscar campos clave
    const buscar = [
      { criterio: 'nombre', campos: Object.keys(conDatos) },
      { criterio: 'dpi', campos: Object.keys(conDatos) },
      { criterio: 'teléfono', campos: Object.keys(conDatos) },
      { criterio: 'email', campos: Object.keys(conDatos) },
      { criterio: 'edad', campos: Object.keys(conDatos) },
      { criterio: 'género', campos: Object.keys(conDatos) },
      { criterio: 'educación', campos: Object.keys(conDatos) }
    ];

    buscar.forEach(({ criterio, campos }) => {
      const encontrado = campos.find(c => c.toLowerCase().includes(criterio));
      reporte += `${criterio.toUpperCase().padEnd(15)} → ${encontrado || '❌ NO ENCONTRADO'}\n`;
    });

    // Mostrar en alert
    const ui = SpreadsheetApp.getUi();
    const resultado = ui.alert(reporte, ui.ButtonSet.OK_CANCEL);

    if (resultado === ui.Button.OK) {
      guardarInspeccionEnLog(reporte, conDatos);
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('inspeccionarPrimerRegistroKobo', e);
  }
}

/**
 * Compara lo que el código BUSCA vs lo que REALMENTE viene en Kobo
 */
function compararMapeosKobo() {
  try {
    const key = getKoboKey();
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json&limit=1',
      { headers: { Authorization: 'Token ' + key }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert('❌ Error: ' + resp.getResponseCode());
      return;
    }

    const primer = JSON.parse(resp.getContentText()).results[0] || {};
    const camposDisponibles = Object.keys(primer);

    // Campos que el código BUSCA actualmente
    const MAPEOS_ACTUALES = {
      'ID': 'creamos_id',
      'Nombre': 'nombre_completo_del_la_participante',
      'DPI': 'numero_de_dpi_opcional',
      'Edad': 'edad',
      'Género': 'genero',
      'Teléfono': 'numero_de_telefono',
      'Zona': 'lugar_de_residencia',
      'Email': 'correo_electronico_opcional',
      'Educación': 'cual_es_el_ultimo_grado_que_completaste',
      'Laboral': 'cual_es_tu_situacion_laboral_actual',
      'Fortalezas': 'que_sabes_hacer_bien',
      'Objetivo': 'que_tipo_de_empleo_estas_buscando_especificamente',
      'Perfil': 'perfil_asignado',
      'Prioridad': 'prioridad_caso',
      'Puntaje': 'puntaje_total_60',
      'Dim1': 'dimension_1_capital_educativo',
      'Dim2': 'dimension_2_capital_laboral',
      'Dim3': 'dimension_3_habilidades_digitales',
      'Dim4': 'dimension_4_claridad_vocacional',
      'Dim5': 'dimension_5_barreras_estructurales',
      'Dim6': 'dimension_6_red_apoyo'
    };

    let reporte = '🔍 COMPARACIÓN: CÓDIGO vs KOBO REAL\n';
    reporte += '═════════════════════════════════════════════\n\n';

    reporte += '🟢 COINCIDENCIAS (campo existe en Kobo):\n';
    reporte += '─────────────────────────────────────────\n';

    let coincidencias = 0;
    let faltantes = 0;

    Object.entries(MAPEOS_ACTUALES).forEach(([nombreSistema, nombreKobo]) => {
      const existe = camposDisponibles.includes(nombreKobo);
      if (existe) {
        const valor = primer[nombreKobo];
        const dato = valor ? ` → ${String(valor).substring(0, 30)}` : '';
        reporte += `✅ ${nombreSistema.padEnd(12)} = ${nombreKobo}${dato}\n`;
        coincidencias++;
      }
    });

    reporte += '\n🔴 FALTANTES (campo NO existe en Kobo):\n';
    reporte += '─────────────────────────────────────────\n';

    Object.entries(MAPEOS_ACTUALES).forEach(([nombreSistema, nombreKobo]) => {
      const existe = camposDisponibles.includes(nombreKobo);
      if (!existe) {
        reporte += `❌ ${nombreSistema.padEnd(12)} busca: "${nombreKobo}"\n`;
        reporte += `    💡 ¿Existe algo parecido en tu Kobo?\n`;
        faltantes++;
      }
    });

    reporte += '\n📊 ESTADÍSTICA:\n';
    reporte += `✅ ${coincidencias} campos están bien mapeados\n`;
    reporte += `❌ ${faltantes} campos NO existen en tu Kobo\n`;
    reporte += `📋 ${camposDisponibles.length} campos totales en Kobo\n`;

    reporte += '\n💡 SOLUCIÓN:\n';
    reporte += '1. Copia los nombres de los campos que están OK\n';
    reporte += '2. Para los que faltan, verifica el nombre exacto en tu formulario Kobo\n';
    reporte += '3. Actualiza los nombres en Paso_a_Paso.gs línea ~590\n';

    SpreadsheetApp.getUi().alert(reporte);
    guardarInspeccionEnLog(reporte, MAPEOS_ACTUALES);

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('compararMapeosKobo', e);
  }
}

/**
 * Descarga un MUESTRA de registros y crea hoja con los datos reales
 */
function exportarMuestraKoboASheet() {
  try {
    const key = getKoboKey();
    const resp = UrlFetchApp.fetch(
      CONFIG.KOBO_URL + '/assets/' + CONFIG.KOBO_ASSET_ID + '/data/?format=json&limit=10',
      { headers: { Authorization: 'Token ' + key }, muteHttpExceptions: true }
    );

    if (resp.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert('❌ Error: ' + resp.getResponseCode());
      return;
    }

    const registros = JSON.parse(resp.getContentText()).results || [];
    const ss = SpreadsheetApp.getActive();

    // Crear nueva hoja
    let hojaInspection = ss.getSheetByName('INSPECCION_KOBO');
    if (hojaInspection) {
      ss.deleteSheet(hojaInspection);
    }
    hojaInspection = ss.insertSheet('INSPECCION_KOBO');

    // Obtener todos los campos únicos
    const camposUnicos = new Set();
    registros.forEach(r => {
      Object.keys(r).forEach(k => {
        if (!k.startsWith('_')) camposUnicos.add(k);
      });
    });

    const campos = Array.from(camposUnicos).sort();

    // Encabezados
    hojaInspection.appendRow(campos);

    // Datos
    registros.forEach(r => {
      const fila = campos.map(c => r[c] || '');
      hojaInspection.appendRow(fila);
    });

    // Formatear
    hojaInspection.getRange(1, 1, 1, campos.length).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');

    SpreadsheetApp.getUi().alert(
      `✅ Creada hoja "INSPECCION_KOBO"\n\n` +
      `📊 ${registros.length} registros de muestra\n` +
      `📋 ${campos.length} campos distintos\n\n` +
      `Ahora puedes:\n` +
      `1. Ver exactamente qué datos tiene Kobo\n` +
      `2. Identificar nombres reales de campos\n` +
      `3. Copiar nombres exactos para actualizar script`
    );

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    logError('exportarMuestraKoboASheet', e);
  }
}

function guardarInspeccionEnLog(reporte, datos) {
  try {
    const ss = SpreadsheetApp.getActive();
    let logHoja = ss.getSheetByName('Log');
    if (!logHoja) {
      logHoja = ss.insertSheet('Log');
      logHoja.appendRow(['Timestamp', 'Tipo', 'Detalles', 'Usuario']);
    }
    logHoja.appendRow([
      new Date(),
      '🔍 INSPECCION',
      reporte,
      Session.getEffectiveUser().getEmail()
    ]);
  } catch (e) {
    Logger.log('Error guardando inspección: ' + e);
  }
}
