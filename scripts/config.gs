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
