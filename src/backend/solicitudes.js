/**
 * Se ejecuta automáticamente mediante el Trigger (onFormSubmit)
 * cada vez que alguien llena el Google Forms vinculado.
 */
function alRecibirFormulario(e) {
  try {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaSolicitudes = spreadSheet.getSheetByName(NOMBRE_HOJA_SOLICITUDES);

    if (!hojaSolicitudes) return;

    // IMPORTANTE: Al ser un Trigger directo de Forms, 'e.values' no existe.
    // Usamos la API nativa de 'e.response' para desempacar los datos.
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();

    // 1. Obtener la Marca Temporal automática
    const marcaTemporal = formResponse.getTimestamp();

    // 2. Variables por defecto
    let cliente = "SIN NOMBRE";
    let correo = "";
    let concepto = "";
    let importeCrudo = "0";

    // 3. Mapear las respuestas en el orden exacto de las preguntas (0 a 3)
    if (itemResponses.length > 0)
      cliente = itemResponses[0].getResponse().toString().toUpperCase().trim();
    if (itemResponses.length > 1)
      correo = itemResponses[1].getResponse().toString().toLowerCase().trim();
    if (itemResponses.length > 2)
      concepto = itemResponses[2].getResponse().toString().toUpperCase().trim();
    if (itemResponses.length > 3)
      importeCrudo = itemResponses[3].getResponse().toString();

    // 4. Limpiar el importe (quitar $, letras o comas accidentalmente escritas)
    let importeNumero = parseFloat(importeCrudo.replace(/[^0-9.-]+/g, ""));
    if (isNaN(importeNumero)) importeNumero = 0;

    // 5. Construir la fila para la hoja de Excel
    const nuevaFila = [
      marcaTemporal,
      cliente,
      correo,
      concepto,
      importeNumero,
      "PENDIENTE", // Confirmación Caja
      "PENDIENTE", // Estado Correo
      "", // Destinatarios Notificados
    ];

    // 6. Escribir en la hoja
    hojaSolicitudes.appendRow(nuevaFila);
  } catch (error) {
    console.error(
      "Error crítico al registrar la solicitud entrante: " + error.message,
    );
  }
}
