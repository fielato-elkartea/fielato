var ADMIN_PIN = "fielatogestioa";

// El grupo "5/6" (con barra) hace que Sheets lo confunda con una fecha, y esto
// pasa incluso escribiendolo como formula de texto si hay mas de uno en la
// misma escritura: la "limpieza automatica de datos" de Sheets "corrige" todas
// las celdas menos la primera, por mucho que el valor ya sea texto literal.
// Cambiar la barra por un guion ("5-6") TAMPOCO basta: Sheets tambien
// reconoce "5-6" como fecha corta (dia-mes del año actual). La proteccion de
// verdad es anteponer una letra ("G5/6"): ninguna fecha empieza asi, asi que
// Sheets no tiene forma de confundirlo con nada y lo guarda como texto plano
// sin mas. Se quita esa "G" de nuevo al leer, de forma transparente para la
// app (que nunca llega a ver "G5/6", solo "5/6" como siempre).
var COLUMNAS_GRUPO = { "Repartos": 2, "Miembros": 3 };
var PREFIJO_GRUPO = "G";

function protegerGrupo(sheetName, values) {
  var colIdx = COLUMNAS_GRUPO[sheetName];
  if (colIdx === undefined) return values;
  return values.map(function(row) {
    var copia = row.slice();
    var v = copia[colIdx];
    if (typeof v === "string" && /^\d{1,2}\/\d{1,2}$/.test(v)) {
      copia[colIdx] = PREFIJO_GRUPO + v;
    }
    return copia;
  });
}

function restaurarGrupo(sheetName, data) {
  var colIdx = COLUMNAS_GRUPO[sheetName];
  if (colIdx === undefined) return data;
  return data.map(function(row) {
    var copia = row.slice();
    var v = copia[colIdx];
    if (typeof v === "string" && v.indexOf(PREFIJO_GRUPO) === 0 && /^\d{1,2}\/\d{1,2}$/.test(v.slice(PREFIJO_GRUPO.length))) {
      copia[colIdx] = v.slice(PREFIJO_GRUPO.length);
    }
    return copia;
  });
}

function doGet(e) {
  // Sin parametro "action": es una visita normal desde el navegador -> servir la app.
  // Con "action": es una llamada de datos de la propia app -> servir JSON como antes.
  if (!e.parameter || !e.parameter.action) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Fielato')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    var ss = SpreadsheetApp.openById("1xAWvZNxxNXH4Ac7yenRUSTJliTteY-TnkMhRzaBVL2s");

    var action = "";
    var sheetName = "";
    var valuesStr = "";
    var pin = "";

    if (e.parameter) {
      action = e.parameter.action || "";
      sheetName = e.parameter.sheet || "";
      valuesStr = e.parameter.values || "";
      pin = e.parameter.pin || "";
    }

    if (e.postData && e.postData.contents) {
      var body = e.postData.contents;
      var contentType = e.postData.type || "";

      if (contentType.indexOf("application/json") >= 0) {
        var parsed = JSON.parse(body);
        action = action || parsed.action || "";
        sheetName = sheetName || parsed.sheet || "";
        valuesStr = valuesStr || (parsed.values ? JSON.stringify(parsed.values) : "");
        pin = pin || parsed.pin || "";
      } else {
        var parts = body.split("&");
        for (var i = 0; i < parts.length; i++) {
          var eq = parts[i].indexOf("=");
          if (eq < 0) continue;
          var k = decodeURIComponent(parts[i].substring(0, eq));
          var v = decodeURIComponent(parts[i].substring(eq + 1).replace(/\+/g, " "));
          if (k === "action") action = v;
          if (k === "sheet") sheetName = v;
          if (k === "values") valuesStr = v;
          if (k === "pin") pin = v;
        }
      }
    }

    if (action === "readAll") {
      var sheetNames = ["Miembros","Calendario","Asistencia","Piezas","Config","Repartos","Tarjetas","Historico","Mancha","BatidaLugar","Capturas","Jornadas"];
      var result = {};
      var tzAll = ss.getSpreadsheetTimeZone();
      sheetNames.forEach(function(name) {
        var sh = ss.getSheetByName(name);
        if (!sh) { result[name] = []; return; }
        var data = sh.getDataRange().getValues();
        var clean = data.map(function(row) {
          return row.map(function(cell) {
            if (cell instanceof Date) return Utilities.formatDate(cell, tzAll, "yyyy-MM-dd");
            if (cell === null || cell === undefined) return "";
            return cell;
          });
        });
        result[name] = restaurarGrupo(name, clean);
      });
      return out(result);
    }

    if (!sheetName) return out({error: "No sheet specified"});
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return out({error: "Sheet not found: " + sheetName});

    if (action === "read") {
      var data = sheet.getDataRange().getValues();
      var tz = ss.getSpreadsheetTimeZone();
      var clean = data.map(function(row) {
        return row.map(function(cell) {
          if (cell instanceof Date) {
            // Use spreadsheet timezone to avoid UTC offset shifting dates
            return Utilities.formatDate(cell, tz, "yyyy-MM-dd");
          }
          if (cell === null || cell === undefined) return "";
          return cell;
        });
      });
      return out(restaurarGrupo(sheetName, clean));
    }

    if (action === "write") {
      if (pin !== ADMIN_PIN) return out({error: "PIN incorrecto"});
      if (!valuesStr) return out({error: "No values provided"});
      var values = JSON.parse(valuesStr);
      sheet.clearContents();
      if (values.length > 0 && values[0].length > 0) {
        var range = sheet.getRange(1, 1, values.length, values[0].length);
        range.setNumberFormat("@");
        range.setValues(protegerGrupo(sheetName, values));
      }
      SpreadsheetApp.flush();
      return out({ok: true, rows: values.length});
    }

    return out({error: "Unknown action: " + action});

  } catch (err) {
    return out({error: err.toString()});
  }
}

function out(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Copia de seguridad semanal: crea una copia completa de la hoja en Drive y
// ademas la manda por correo como archivo Excel adjunto.
// Ejecutar via un trigger de tiempo (todos los miercoles), configurado a mano
// en el editor de Apps Script: icono del reloj -> Add Trigger -> backupSemanal
// -> Time-driven -> Week timer -> Every Wednesday.
function backupSemanal() {
  var SPREADSHEET_ID = "1xAWvZNxxNXH4Ac7yenRUSTJliTteY-TnkMhRzaBVL2s";
  var BACKUP_FOLDER_NAME = "Fielato - Copias de seguridad";
  var EMAIL_DESTINO = "guregestioa@gmail.com";

  var original = DriveApp.getFileById(SPREADSHEET_ID);
  var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);

  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var nombre = "Guregestioa Fielato - backup " + fecha;
  var copia = original.makeCopy(nombre, folder);

  var excelBlob = copia.getAs(MimeType.MICROSOFT_EXCEL).setName(nombre + ".xlsx");
  MailApp.sendEmail({
    to: EMAIL_DESTINO,
    subject: "Fielato - Copia de seguridad semanal (" + fecha + ")",
    body: "Copia de seguridad semanal de la hoja de Fielato del " + fecha + ", adjunta en Excel.\n\nTambien se ha guardado en Drive, en la carpeta \"" + BACKUP_FOLDER_NAME + "\".",
    attachments: [excelBlob]
  });
}

// Funcion temporal de diagnostico: lista los nombres exactos de todas las
// pestanas que el script ve en la hoja. Ejecutar desde el editor (boton Run)
// con "listarHojas" seleccionado en el desplegable de funciones, y mirar
// el resultado en Ver -> Registros (o Ctrl+Enter). Se puede borrar despues.
function listarHojas() {
  var ss = SpreadsheetApp.openById("1xAWvZNxxNXH4Ac7yenRUSTJliTteY-TnkMhRzaBVL2s");
  var sheets = ss.getSheets();
  sheets.forEach(function(s) { Logger.log("[" + s.getName() + "]"); });
}
