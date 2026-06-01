function crearCapetaRaiz() {
  const config = obtenerConfiguracion();
  const nombreRaiz = config.carpetaRaiz; // Dinámica
  const iter = DriveApp.getFoldersByName(nombreRaiz);
  if (iter.hasNext()) return iter.next();
  return DriveApp.createFolder(nombreRaiz);
}

function crearCarpetaTipoRecibo(nombreHoja) {
  const raiz = crearCapetaRaiz();
  const nombreCarpeta = `Recibos-${nombreHoja}`;
  const iter = raiz.getFoldersByName(nombreCarpeta);
  if (iter.hasNext()) return iter.next();
  return raiz.createFolder(nombreCarpeta);
}

function crearCarpertaPorFecha(fecha, nombreHoja) {
  const nombreCarpetaFecha = Utilities.formatDate(
    fecha,
    Session.getScriptTimeZone(),
    FORMATO_FECHA_CARPETA,
  );
  // Se crea dentro de la carpeta correspondiente a su hoja
  const carpetaTipo = crearCarpetaTipoRecibo(nombreHoja);
  const iter = carpetaTipo.getFoldersByName(nombreCarpetaFecha);
  if (iter.hasNext()) return iter.next();
  return carpetaTipo.createFolder(nombreCarpetaFecha);
}
