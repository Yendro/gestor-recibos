function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Generar Recibos")
    .addItem("Abrir Panel de Recibos", "mostrarBarraLateral")
    .addToUi();
}

// Abrir sidebar
function mostrarBarraLateral() {
  const html = HtmlService.createHtmlOutputFromFile("src/frontend/sidebar")
    .setTitle("Generador de Recibos")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function generarRecibosDesdeMenu() {
  const resultado = generarTodosPendientes();
  const ui = SpreadsheetApp.getUi();
  const exitosos = resultado.filter((r) => r.success).length;
  const fallidos = resultado.length - exitosos;

  ui.alert(
    "Resultado",
    `Recibos generados: ${exitosos}\nErrores: ${fallidos}`,
    ui.ButtonSet.OK,
  );
}

function obtenerEstadoPendientes() {
  const pendientes = obtenerFilasPendientes();
  return pendientes.map((p) => ({
    cliente: p.cliente,
    folio: p.folio,
    rowIndex: p.rowIndex,
  }));
}
