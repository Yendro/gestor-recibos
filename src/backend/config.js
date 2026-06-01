const FORMATO_FECHA_CARPETA = "yyyy-MM-dd";
const NOMBRE_HOJA_CONFIG = "Config";

const COLUMNAS = {
  CLIENTE: 0,
  IMPORTE: 1,
  CONCEPTO: 2,
  FECHA_PAGO: 3,
  IMPORTE_LETRA: 4,
  FOLIO: 5,
  FECHA_CREACION: 6,
  STATUS: 7,
  ARCHIVO: 8,
  ESTADO_FIRMA: 9,
  ESTADO_CORREO: 10,
  DESTINATARIOS: 11,
  ESTATUS_FOTO: 12,
  ELIMINAR: 13,
};

/**
 * Lee la hoja de 'Config' y construye un objeto con las plantillas y correos
 * @returns {Object} { plantillas: Array, directorioCorreos: Array }
 */
function obtenerConfiguracion() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaConfig = spreadSheet.getSheetByName(NOMBRE_HOJA_CONFIG);

  if (!hojaConfig) {
    throw new Error(
      "La hoja de Configuración no existe. Ejecuta 'Inicializar Sistema' desde el menú.",
    );
  }

  // Leer Carpeta Raíz
  const carpetaRaiz =
    hojaConfig.getRange("H2").getValue() || "Recibos-Generales";

  const datosPlantillas = hojaConfig.getRange("A3:D10").getValues();
  const plantillas = [];
  for (const fila of datosPlantillas) {
    if (fila[0] && fila[1]) {
      plantillas.push({
        nombreHoja: fila[0].toString().trim(),
        idPlantilla: fila[1].toString().trim(),
        editores: fila[2]
          ? fila[2]
              .toString()
              .split(",")
              .map((e) => e.trim())
          : [],
        prefijoFolio: fila[3] ? fila[3].toString().trim().toUpperCase() : "REC",
      });
    }
  }

  // 2. Leer Correos (Rango E3:F30)
  const datosCorreos = hojaConfig.getRange("E3:F30").getValues();
  const directorioCorreos = [];

  for (const fila of datosCorreos) {
    if (fila[0] && fila[1]) {
      directorioCorreos.push({
        nombre: fila[0].toString().trim(),
        correo: fila[1].toString().trim(),
      });
    }
  }

  return {
    carpetaRaiz: carpetaRaiz,
    plantillas: plantillas,
    directorioCorreos: directorioCorreos,
  };
}
