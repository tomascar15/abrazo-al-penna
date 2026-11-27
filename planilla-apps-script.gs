/**
 * Abrazo al Penna · Recibe las inscripciones de la página y las guarda en esta planilla.
 * Cada formulario va a su propia pestaña (Amigos, Voluntarios, Las Manos, Guía amiga, etc.).
 * Las columnas se crean solas la primera vez que llega un dato nuevo.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var p = e.parameter || {};
    var nombreHoja = String(p.hoja || 'Otros').slice(0, 40);
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = libro.getSheetByName(nombreHoja) || libro.insertSheet(nombreHoja);

    var datos = {};
    Object.keys(p).forEach(function (k) { if (k !== 'hoja') datos[k] = String(p[k]).slice(0, 2000); });
    datos['Fecha'] = new Date();

    var encabezado = hoja.getLastRow() > 0
      ? hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0]
      : ['Fecha'];
    Object.keys(datos).forEach(function (k) { if (encabezado.indexOf(k) === -1) encabezado.push(k); });
    hoja.getRange(1, 1, 1, encabezado.length).setValues([encabezado]).setFontWeight('bold');
    hoja.setFrozenRows(1);

    var fila = encabezado.map(function (k) { return datos[k] !== undefined ? datos[k] : ''; });
    hoja.appendRow(fila);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
