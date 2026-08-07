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

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // === APPEND (agregar fila) ===
    if (!data.action || data.action === 'append') {
      var sheet = ss.getSheetByName(data.sheet);
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
          var opSheet = ss.getSheetByName(op.sheet);
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
