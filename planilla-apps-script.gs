function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var p = e.parameter || {};
    var nombreHoja = String(p.hoja || 'Otros').slice(0, 40);
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = libro.getSheetByName(nombreHoja) || libro.insertSheet(nombreHoja);

    var datos = {};
    Object.keys(p).forEach(function (k) {
      if (k !== 'hoja') datos[k] = String(p[k]).slice(0, 2000);
    });
    datos['Fecha'] = new Date();

    var respuesta = { ok: true };
    if (nombreHoja === 'Amigos') {
      var props = PropertiesService.getScriptProperties();
      var numero = Number(props.getProperty('contadorAmigos') || 0) + 1;
      props.setProperty('contadorAmigos', String(numero));
      datos['N.º de amigo'] = numero;
      respuesta.numero = numero;
    }

    var encabezado = hoja.getLastRow() > 0
      ? hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0]
      : ['Fecha'];
    Object.keys(datos).forEach(function (k) {
      if (encabezado.indexOf(k) === -1) encabezado.push(k);
    });
    hoja.getRange(1, 1, 1, encabezado.length).setValues([encabezado]).setFontWeight('bold');
    hoja.setFrozenRows(1);

    var fila = encabezado.map(function (k) {
      return datos[k] !== undefined ? datos[k] : '';
    });
    hoja.appendRow(fila);

    return ContentService.createTextOutput(JSON.stringify(respuesta))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
