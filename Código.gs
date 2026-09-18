const SPREADSHEET_ID = "1vH-iOKG9LY2kYuh6lM1AC3BlITwvPdyo";
const SHEET_NAME = "TARJETAS";
const COL_CODIGO = 2, COL_REFERENCIA = 3, COL_TIPO = 4, COL_PRECIO = 5;
const COL_ESTADO = 6, COL_FECHA_USO = 8;

function doGet(e) {
  const codigo = e && e.parameter && e.parameter.codigo ? String(e.parameter.codigo).trim() : "";
  if (codigo) return htmlResultado(validarCodigo(codigo));
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Lector QR ITEGT 2026");
}

function validarCodigo(codigo) {
  codigo = String(codigo || "").trim();
  if (!codigo) return {tipo:"ERROR", titulo:"SIN CÓDIGO", mensaje:"No se recibió ningún código QR."};

  const hoja = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!hoja) return {tipo:"ERROR", titulo:"ERROR DEL SISTEMA", mensaje:"No existe la hoja TARJETAS."};

  const last = hoja.getLastRow();
  const datos = last < 2 ? [] : hoja.getRange(2,1,last-1,10).getValues();
  let fila=-1, reg=null;
  for (let i=0;i<datos.length;i++) {
    if (String(datos[i][COL_CODIGO-1]).trim() === codigo) { fila=i+2; reg=datos[i]; break; }
  }
  if (fila < 0) return {tipo:"FALSA", titulo:"QR FALSO", mensaje:"El código no está registrado.", codigo};

  const referencia=String(reg[COL_REFERENCIA-1]||"").trim();
  const tipo=String(reg[COL_TIPO-1]||"").trim();
  const precio=String(reg[COL_PRECIO-1]||"").trim();

  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const estado=String(hoja.getRange(fila,COL_ESTADO).getValue()).trim().toUpperCase();

    if (estado==="UTILIZADA") {
      return {tipo:"UTILIZADA", titulo:"TARJETA YA UTILIZADA", mensaje:"INGRESO RECHAZADO",
        codigo, referencia, fecha:formatearFecha(hoja.getRange(fila,COL_FECHA_USO).getValue())};
    }
    if (estado==="CANCELADA") return {tipo:"CANCELADA", titulo:"TARJETA CANCELADA", mensaje:"INGRESO RECHAZADO", codigo, referencia};
    if (estado!=="VALIDA") return {tipo:"ERROR", titulo:"ESTADO NO VÁLIDO", mensaje:"Estado: "+estado, codigo, referencia};

    hoja.getRange(fila,COL_ESTADO).setValue("UTILIZADA");
    hoja.getRange(fila,COL_FECHA_USO).setValue(new Date());
    SpreadsheetApp.flush();

    return {tipo:"VALIDA", titulo:"TARJETA VÁLIDA", mensaje:"ACCESO AUTORIZADO",
      codigo, referencia, tipoTarjeta:tipo, precio, fecha:formatearFecha(new Date())};
  } finally { lock.releaseLock(); }
}

function formatearFecha(f) {
  return f instanceof Date && !isNaN(f.getTime())
    ? Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss") : "";
}

function htmlResultado(r) {
  const color=r.tipo==="VALIDA"?"#16803a":r.tipo==="UTILIZADA"?"#e67e00":"#c62828";
  const icon=r.tipo==="VALIDA"?"✓":r.tipo==="UTILIZADA"?"⚠":"✕";
  return HtmlService.createHtmlOutput(
    '<!doctype html><html lang="es"><head><meta name="viewport" content="width=device-width,initial-scale=1"><meta charset="UTF-8">'+
    '<style>body{font-family:Arial;margin:0;background:#f4f4f4;display:flex;align-items:center;justify-content:center;min-height:100vh}.card{background:#fff;width:90%;max-width:520px;padding:35px 25px;border-radius:22px;text-align:center;box-shadow:0 8px 25px #0002}h1{color:'+color+'}.dato{font-size:18px;line-height:1.7;color:#444}</style></head><body><div class="card">'+
    '<div style="font-size:58px">'+icon+'</div><h1>'+esc(r.titulo)+'</h1><h2>'+esc(r.mensaje)+'</h2><div class="dato">'+
    (r.referencia?'Tarjeta: <b>'+esc(r.referencia)+'</b><br>':'')+
    (r.codigo?'Código: <b>'+esc(r.codigo)+'</b><br>':'')+
    (r.tipoTarjeta?'Tipo: '+esc(r.tipoTarjeta)+'<br>':'')+
    (r.precio?'Precio: L '+esc(r.precio)+'<br>':'')+
    (r.fecha?'Fecha: '+esc(r.fecha):'')+
    '</div><hr><small>IV Festival Interinstitucional de la Canción Popular<br>Instituto Técnico Eulogio Galeano Trejo</small></div></body></html>'
  );
}
function esc(s){return String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#039;");}
