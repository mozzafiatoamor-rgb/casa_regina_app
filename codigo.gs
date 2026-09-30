// =============================================
// HOTEL CASA REGINA - Google Apps Script
// =============================================
// Instrucciones:
// 1. Abre tu Google Sheet
// 2. Ve a Extensiones → Apps Script
// 3. Borra todo el código que aparezca
// 4. Pega este código completo
// 5. Clic en Implementar → Administrar implementaciones
// 6. Editar (lápiz) → Versión: Nueva versión → Implementar
// =============================================

// Hojas que se crean automáticamente (con encabezados) si aún no existen
var SHEET_HEADERS = {
  '💸 Gastos': ['ID', 'Fecha', 'Categoria', 'Concepto', 'Monto', 'Tipo', 'Responsable', 'Notas']
};

// === Repositorio de GitHub para subir imágenes (logos/fotos de marcas) ===
// El token se guarda en Propiedades del script (GITHUB_TOKEN); NO se escribe aquí.
// En el editor de Apps Script:
//   Configuración del proyecto (⚙️) → Propiedades del script → Agregar propiedad:
//     GITHUB_TOKEN = <token de GitHub con permiso "Contents: Read and write" en este repo>
// (Opcional) sobreescribe owner/repo/branch con las propiedades GH_OWNER, GH_REPO, GH_BRANCH.
var GH_DEFAULT_OWNER  = 'mozzafiatoamor-rgb';
var GH_DEFAULT_REPO   = 'casa_regina_app';
var GH_DEFAULT_BRANCH = 'main';

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet && SHEET_HEADERS[name]) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(SHEET_HEADERS[name]);
  }
  return sheet;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // === UPLOAD IMAGE (sube una imagen al repo de GitHub y devuelve su URL) ===
    if (data.action === 'uploadImage') {
      return uploadImageToGitHub(data);
    }

    // === APPEND (agregar fila) ===
    if (!data.action || data.action === 'append') {
      var sheet = getOrCreateSheet(ss, data.sheet);
      if (!sheet) {
        return jsonResponse({ success: false, error: 'Hoja no encontrada: ' + data.sheet });
      }
      sheet.appendRow(data.values);
      return jsonResponse({ success: true, action: 'append', sheet: data.sheet });
    }

    // === UPDATE (actualizar celda o rango) ===
    if (data.action === 'update') {
      var sheet = ss.getSheetByName(data.sheet);
      if (!sheet) {
        return jsonResponse({ success: false, error: 'Hoja no encontrada: ' + data.sheet });
      }
      if (data.row && data.col) {
        sheet.getRange(data.row, data.col).setValue(data.value);
        return jsonResponse({ success: true, action: 'update', sheet: data.sheet });
      }
      if (data.row && data.values) {
        var range = sheet.getRange(data.row, 1, 1, data.values.length);
        range.setValues([data.values]);
        return jsonResponse({ success: true, action: 'update', sheet: data.sheet });
      }
      return jsonResponse({ success: false, error: 'Faltan parámetros para update' });
    }

    // === UPDATE BY ID (buscar por ID y actualizar) ===
    if (data.action === 'updateById') {
      var sheet = ss.getSheetByName(data.sheet);
      if (!sheet) {
        return jsonResponse({ success: false, error: 'Hoja no encontrada: ' + data.sheet });
      }
      var dataRange = sheet.getDataRange().getValues();
      for (var i = 1; i < dataRange.length; i++) {
        if (dataRange[i][0] === data.id) {
          var range = sheet.getRange(i + 1, 1, 1, data.values.length);
          range.setValues([data.values]);
          return jsonResponse({ success: true, action: 'updateById', row: i + 1 });
        }
      }
      return jsonResponse({ success: false, error: 'ID no encontrado: ' + data.id });
    }

    // === DELETE (borrar fila) ===
    if (data.action === 'delete') {
      var sheet = ss.getSheetByName(data.sheet);
      if (!sheet) {
        return jsonResponse({ success: false, error: 'Hoja no encontrada: ' + data.sheet });
      }
      sheet.deleteRow(data.row);
      return jsonResponse({ success: true, action: 'delete', sheet: data.sheet, row: data.row });
    }

    // === DELETE BY ID (buscar por ID y borrar) ===
    if (data.action === 'deleteById') {
      var sheet = ss.getSheetByName(data.sheet);
      if (!sheet) {
        return jsonResponse({ success: false, error: 'Hoja no encontrada: ' + data.sheet });
      }
      var dataRange = sheet.getDataRange().getValues();
      for (var i = dataRange.length - 1; i >= 1; i--) {
        if (dataRange[i][0] === data.id) {
          sheet.deleteRow(i + 1);
          return jsonResponse({ success: true, action: 'deleteById', row: i + 1 });
        }
      }
      return jsonResponse({ success: false, error: 'ID no encontrado: ' + data.id });
    }

    // === BATCH APPEND (agregar múltiples filas) ===
    if (data.action === 'batchAppend') {
      var sheet = ss.getSheetByName(data.sheet);
      if (!sheet) {
        return jsonResponse({ success: false, error: 'Hoja no encontrada: ' + data.sheet });
      }
      var rows = data.rows || [];
      for (var i = 0; i < rows.length; i++) {
        sheet.appendRow(rows[i]);
      }
      return jsonResponse({ success: true, action: 'batchAppend', count: rows.length });
    }

    // === MULTI (múltiples operaciones en una llamada) ===
    if (data.action === 'multi') {
      var ops = data.operations || [];
      var results = [];
      for (var oi = 0; oi < ops.length; oi++) {
        var op = ops[oi];
        try {
          var opSheet = getOrCreateSheet(ss, op.sheet);
          if (!opSheet) { results.push({ success: false, error: 'Hoja no encontrada: ' + op.sheet }); continue; }
          if (op.type === 'append') {
            opSheet.appendRow(op.values);
            results.push({ success: true });
          } else if (op.type === 'updateById') {
            var opData = opSheet.getDataRange().getValues();
            var found = false;
            for (var oj = 1; oj < opData.length; oj++) {
              if (opData[oj][0] === op.id) {
                opSheet.getRange(oj + 1, 1, 1, op.values.length).setValues([op.values]);
                found = true; break;
              }
            }
            results.push({ success: found, error: found ? null : 'ID no encontrado' });
          } else {
            results.push({ success: false, error: 'Tipo no reconocido: ' + op.type });
          }
        } catch (opErr) {
          results.push({ success: false, error: opErr.toString() });
        }
      }
      return jsonResponse({ success: true, action: 'multi', results: results });
    }

    return jsonResponse({ success: false, error: 'Acción no reconocida: ' + data.action });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  // Si no hay parámetros, responder con status
  if (!e || !e.parameter || !e.parameter.action) {
    return jsonResponse({ success: true, message: 'Hotel PMS - API activa' });
  }

  var action = e.parameter.action;

  // === READ (leer datos de una hoja — usado por páginas públicas sin API key) ===
  if (action === 'read') {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetName = e.parameter.sheet;
      var range = e.parameter.range;

      if (!sheetName || !range) {
        return jsonResponse({ error: 'Faltan parámetros: sheet y range' });
      }

      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        return jsonResponse({ error: 'Hoja no encontrada: ' + sheetName });
      }

      var data = sheet.getRange(range).getValues();
      // Filtrar filas vacías (igual que el API de Google Sheets)
      data = data.filter(function(row) {
        return row.some(function(cell) { return cell !== '' && cell !== null; });
      });
      // Convertir Dates a strings YYYY-MM-DD (el API de Sheets las devuelve como texto)
      data = data.map(function(row) {
        return row.map(function(cell) {
          if (cell instanceof Date) {
            var y = cell.getFullYear();
            var m = ('0' + (cell.getMonth() + 1)).slice(-2);
            var d = ('0' + cell.getDate()).slice(-2);
            return y + '-' + m + '-' + d;
          }
          return cell;
        });
      });
      return ContentService
        .createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return jsonResponse({ error: err.toString() });
    }
  }

  return jsonResponse({ error: 'Acción no reconocida: ' + action });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Sube una imagen (base64) a la carpeta assets/marcas del repo de GitHub y devuelve su URL directa.
function uploadImageToGitHub(data) {
  try {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('GITHUB_TOKEN');
    if (!token) {
      return jsonResponse({ success: false, error: 'Falta GITHUB_TOKEN en Propiedades del script. Agrégalo en Configuración del proyecto → Propiedades del script.' });
    }
    var owner  = props.getProperty('GH_OWNER')  || GH_DEFAULT_OWNER;
    var repo   = props.getProperty('GH_REPO')   || GH_DEFAULT_REPO;
    var branch = props.getProperty('GH_BRANCH') || GH_DEFAULT_BRANCH;

    // Limpia el base64 (quita el prefijo data:image/...;base64, si viene incluido)
    var b64 = String(data.base64 || '');
    var comma = b64.indexOf(',');
    if (b64.indexOf('data:') === 0 && comma !== -1) b64 = b64.substring(comma + 1);
    b64 = b64.replace(/\s/g, '');
    if (!b64) return jsonResponse({ success: false, error: 'Imagen vacía' });

    var ext = String(data.ext || 'jpg').replace(/[^a-z0-9]/gi, '').toLowerCase() || 'jpg';
    var safe = String(data.name || 'img').toLowerCase()
      .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').substring(0, 40) || 'img';
    var sub = String(data.folder || '').replace(/[^a-z0-9_-]/gi, '');
    var dir = 'assets/marcas' + (sub ? ('/' + sub) : '');
    var rand = Utilities.getUuid().substring(0, 8);
    var path = dir + '/' + safe + '-' + Date.now() + '-' + rand + '.' + ext;

    var apiUrl = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + encodeURI(path);
    var payload = { message: 'PMS: subir ' + path, content: b64, branch: branch };
    var resp = UrlFetchApp.fetch(apiUrl, {
      method: 'put',
      contentType: 'application/json',
      headers: {
        'Authorization': 'token ' + token,
        'Accept': 'application/vnd.github+json',
        'User-Agent': 'CasaReginaPMS'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = {};
    try { body = JSON.parse(resp.getContentText() || '{}'); } catch (e) {}
    if (code >= 200 && code < 300 && body.content) {
      var rawUrl = body.content.download_url ||
        ('https://raw.githubusercontent.com/' + owner + '/' + repo + '/' + branch + '/' + path);
      return jsonResponse({ success: true, url: rawUrl, path: path });
    }
    return jsonResponse({ success: false, error: 'GitHub ' + code + ': ' + (body.message || resp.getContentText()) });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}
