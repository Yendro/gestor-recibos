// Recibimos el nombreHoja
function enviarReciboPorCorreo(rowIndex, nombreHoja, destinatariosArr) {
  try {
    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
    const urlDoc = hoja
      .getRange(rowIndex, COLUMNAS.ARCHIVO + 1)
      .getValue()
      .toString();
    const folio = hoja.getRange(rowIndex, COLUMNAS.FOLIO + 1).getValue();
    const cliente = hoja.getRange(rowIndex, COLUMNAS.CLIENTE + 1).getValue();

    const match = urlDoc.match(/[-\w]{25,}/);
    if (!match) throw new Error("No se encontró un documento válido.");
    const docId = match[0];

    // LIMPIEZA FINAL DEL DOCUMENTO
    // Abrimos el doc una última vez para revisar si quedó la etiqueta de foto vacía
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const elementoId = body.findText("{{IdentificacionCliente}}");
    if (elementoId) {
      // Si la etiqueta sigue ahí (el cliente no quiso foto), la borramos para que no salga en el PDF
      elementoId
        .getElement()
        .asText()
        .replaceText("{{IdentificacionCliente}}", "");
    }
    doc.saveAndClose();

    // CONVERSIÓN NATIVA A PDF
    const pdfBlob = DriveApp.getFileById(docId)
      .getAs(MimeType.PDF)
      .setName(`${folio}_${cliente}.pdf`);

    // Preparar la plantilla HTML del correo
    const htmlTemplate = HtmlService.createTemplateFromFile(
      "src/frontend/correo",
    );
    htmlTemplate.cliente = cliente;
    htmlTemplate.folio = folio;
    const bodyCorreo = htmlTemplate.evaluate().getContent();

    const correosDestino = destinatariosArr.join(",");

    MailApp.sendEmail({
      to: correosDestino,
      subject: `Recibo Confirmado - ${folio}`,
      htmlBody: bodyCorreo,
      attachments: [pdfBlob],
    });

    hoja.getRange(rowIndex, COLUMNAS.ESTADO_CORREO + 1).setValue("ENVIADO");
    hoja
      .getRange(rowIndex, COLUMNAS.DESTINATARIOS + 1)
      .setValue(correosDestino);

    return { success: true, message: "Enviado con éxito" };
  } catch (error) {
    console.error("Error en enviarReciboPorCorreo:", error);
    return { success: false, message: error.toString() };
  }
}
