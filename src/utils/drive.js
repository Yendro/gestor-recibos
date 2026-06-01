function crearCapetaRaiz() {
  const iter = DriveApp.getFoldersByName(CARPETA_RAIZ);
  if (iter.hasNext()) return iter.next();
  return DriveApp.createFolder(CARPETA_RAIZ);
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
