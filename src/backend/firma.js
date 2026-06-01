function guardarFotoIdentificacion(base64Id, fila, nombreHoja) {
  try {
    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
    const urlDoc = hoja
      .getRange(fila, COLUMNAS.ARCHIVO + 1)
      .getValue()
      .toString()
      .trim();
    const match = urlDoc.match(/[-\w]{25,}/);
    if (!match) throw new Error("URL de archivo no válida.");
    const docId = match[0];

    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const elementoId = body.findText("{{IdentificacionCliente}}");

    if (!elementoId) {
      throw new Error(
        "La foto ya fue asignada previamente o no existe el campo en la plantilla.",
      );
    }

    if (base64Id) {
      const base64DataId = base64Id.split(",")[1];
      const mimeType = base64Id.substring(
        base64Id.indexOf(":") + 1,
        base64Id.indexOf(";"),
      );
      const blobId = Utilities.newBlob(
        Utilities.base64Decode(base64DataId),
        mimeType,
        "identificacion." + mimeType.split("/")[1],
      );

      const textElementId = elementoId.getElement().asText();
      const parentId = textElementId.getParent();
      textElementId.replaceText("{{IdentificacionCliente}}", "");

      const imgId = parentId.asParagraph().appendInlineImage(blobId);
      const anchoId = 500;
      const propId = imgId.getHeight() / imgId.getWidth();
      imgId.setWidth(anchoId);
      imgId.setHeight(anchoId * propId);
    }

    doc.saveAndClose();

    // Marcamos la foto como capturada en el Sheets
    hoja.getRange(fila, COLUMNAS.ESTATUS_FOTO + 1).setValue("CAPTURADA");

    return { success: true };
  } catch (error) {
    console.error("Error en guardarFotoIdentificacion:", error);
    return { success: false, message: error.toString() };
  }
}

function procesarYGuardarFirmaWebApp(base64Firma, fila, nombreHoja) {
  try {
    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
    const urlDoc = hoja
      .getRange(fila, COLUMNAS.ARCHIVO + 1)
      .getValue()
      .toString()
      .trim();
    const match = urlDoc.match(/[-\w]{25,}/);
    if (!match) throw new Error("URL de archivo no válida en la celda.");
    const docId = match[0];

    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Procesar la Firma
    const base64DataFirma = base64Firma.split(",")[1];
    const imageBlobFirma = Utilities.newBlob(
      Utilities.base64Decode(base64DataFirma),
      "image/png",
      "firma.png",
    );
    const elementoFirma = body.findText("{{Firma}}");
    const anchoFirmaDeseado = 250;

    if (elementoFirma) {
      const textElement = elementoFirma.getElement().asText();
      const parent = textElement.getParent();
      textElement.replaceText("{{Firma}}", "");
      const img = parent.asParagraph().appendInlineImage(imageBlobFirma);
      const proporcion = img.getHeight() / img.getWidth();
      img.setWidth(anchoFirmaDeseado);
      img.setHeight(anchoFirmaDeseado * proporcion);
    }

    doc.saveAndClose();
    hoja.getRange(fila, COLUMNAS.ESTADO_FIRMA + 1).setValue("FIRMADO");

    return { success: true };
  } catch (error) {
    console.error("Error en procesarYGuardarFirmaWebApp:", error);
    return { success: false, message: error.toString() };
  }
}
