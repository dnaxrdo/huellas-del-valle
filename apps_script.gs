// Pegar en la hoja de Google: Extensiones > Apps Script.
// Implementar > Nueva implementación > Aplicación web
//   Ejecutar como: Yo   |   Quién tiene acceso: Cualquier usuario
// Copiar la URL que termina en /exec y pegarla en CONFIG.API_URL del index.html.
// Las hojas se llaman "Agenda" y "Citas". Formatear fecha y hora como "Texto sin formato".

function hoja_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function enc_(fila) { return fila.map(function (x) { return String(x).trim().toLowerCase(); }); }

// Lee una hoja completa como lista de objetos (clave = encabezado en minúsculas)
function doGet(e) {
  var nombre = e.parameter.hoja === "Agenda" ? "Agenda" : "Citas";
  var v = hoja_(nombre).getDataRange().getDisplayValues();
  var h = enc_(v.shift());
  var filas = v.filter(function (r) { return r.join("") !== ""; }).map(function (r) {
    var o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o;
  });
  return out_(filas);
}

// Cambia el estado de asistencia de una cita (busca por la columna "id")
function doPost(e) {
  var b = JSON.parse(e.postData.contents);
  if (b.action !== "estado" || ["Asistió", "No asistió"].indexOf(b.estado) < 0) return out_({ ok: false, error: "Solicitud no válida" });
  var s = hoja_("Citas"), v = s.getDataRange().getDisplayValues(), h = enc_(v[0]);
  var ci = h.indexOf("id"), ce = h.indexOf("estado");
  if (ci < 0 || ce < 0) return out_({ ok: false, error: "Faltan las columnas id o estado" });
  for (var i = 1; i < v.length; i++) {
    if (v[i][ci] === String(b.id)) { s.getRange(i + 1, ce + 1).setValue(b.estado); return out_({ ok: true }); }
  }
  return out_({ ok: false, error: "No se encontró la cita" });
}
