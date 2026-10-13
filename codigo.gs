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

// ⚠️ EJECUTA ESTA FUNCIÓN UNA VEZ desde el editor (botón ▶ Ejecutar, con "autorizar" seleccionada)
// para conceder el permiso de conexión externa (UrlFetchApp). Acepta el diálogo de permisos que aparece.
// Sin esto, subir logos/fotos falla con "No cuentas con el permiso para llamar a UrlFetchApp.fetch".
function autorizar() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('GITHUB_TOKEN');
  // Hace una llamada externa mínima para forzar el consentimiento del permiso.
  var resp = UrlFetchApp.fetch('https://api.github.com/rate_limit', {
    method: 'get',
    headers: token ? { 'Authorization': 'token ' + token, 'User-Agent': 'CasaReginaPMS' } : { 'User-Agent': 'CasaReginaPMS' },
    muteHttpExceptions: true
  });
  Logger.log('Permiso OK. GitHub respondió: ' + resp.getResponseCode());
  Logger.log(token ? 'GITHUB_TOKEN presente ✅' : 'FALTA GITHUB_TOKEN ❌ (agrégalo en Propiedades del script)');
}

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

    // === PANEL ADMIN (gestión de personal; valida que sea admin) ===
    if (data.action === 'staffAdmin') {
      return staffAdmin(data);
    }

    // === APPEND (agregar fila) ===
    if (!data.action || data.action === 'append') {
      var sheet = getOrCreateSheet(ss, data.sheet);
      if (!sheet) {
        return jsonResponse({ success: false, error: 'Hoja no encontrada: ' + data.sheet });
      }
      sheet.appendRow(data.values);
      _supaTry(function () {
        var v = data.values;
        var rows = (v && v.length && Array.isArray(v[0])) ? v : [v]; // acepta 1 fila o varias
        _supaMirrorRows(data.sheet, rows);
      });
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
          _supaTry(function () { _supaMirrorRows(data.sheet, [data.values]); });
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
          _supaTry(function () { _supaMirrorDelete(data.sheet, data.id); });
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
      _supaTry(function () { _supaMirrorRows(data.sheet, rows); });
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
      _supaTry(function () { _supaMirrorMulti(ops); });
      return jsonResponse({ success: true, action: 'multi', results: results });
    }

    return jsonResponse({ success: false, error: 'Acción no reconocida: ' + data.action });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// Columnas (en orden de la hoja) para leer de Supabase y devolver como arreglos.
var SUPA_READ_COLS = {
  '🏨 Habitaciones': { table: 'habitaciones', cols: ['id', 'numero', 'tipo', 'tarifa', 'estado', 'notas'] },
  '📋 Reservas': { table: 'reservas', cols: ['id', 'habitacion_num', 'huesped', 'telefono', 'email', 'fecha_entrada', 'fecha_salida', 'adultos', 'ninos', 'tarifa', 'tarifa_total', 'estado', 'notas', 'creado_por', 'fecha_creacion'] },
  '🧹 Housekeeping': { table: 'housekeeping', cols: ['id', 'fecha', 'habitacion_num', 'tarea', 'prioridad', 'estado', 'asignado_a', 'notas', 'hora_inicio', 'hora_fin', 'verificado', 'checklist'] },
  '📜 Bitácora': { table: 'bitacora', cols: ['fecha', 'hora', 'usuario', 'accion', 'detalle', 'tipo'], order: 'id.asc' },
  '👤 Usuarios': { table: 'usuarios', cols: ['id', 'usuario', 'password', 'nombre', 'rol'] },
  '💰 Cobros': { table: 'cobros', cols: ['id', 'reserva_id', 'fecha', 'hora', 'monto', 'metodo_pago', 'concepto', 'notas', 'responsable'] },
  '👥 Clientes': { table: 'clientes', cols: ['id', 'nombre', 'telefono', 'email', 'notas', 'primera_visita', 'ultima_visita', 'visitas', 'total_gastado'] },
  '📦 Inventario': { table: 'inventario', cols: ['id', 'categoria', 'producto', 'stock', 'minimo', 'unidad', 'notas', 'ultima_actualizacion'] },
  '⚙️ Config': { table: 'config', cols: ['config_key', 'config_value'] },
  '🍳 Desayunos': { table: 'desayunos', cols: ['codigo', 'reserva_id', 'huesped', 'habitacion', 'fecha', 'personas', 'estado', 'validado_por', 'hora_validacion', 'reg_id'] },
  '📝 Cotizaciones': { table: 'cotizaciones', cols: ['id', 'numero', 'fecha', 'nombre', 'telefono', 'email', 'checkin', 'checkout', 'adultos', 'ninos', 'habitaciones', 'opciones', 'subtotal', 'iva', 'total', 'estado'] },
  '🔑 Accesos': { table: 'accesos', cols: ['reserva_id', 'codigo', 'notas'] },
  '📝 Registro': { table: 'registro', cols: ['reg_id', 'res_id', 'nombre', 'habitacion_num', 'fecha', 'tipo_id', 'num_id', 'nacionalidad', 'procedencia', 'motivo', 'observaciones', 'ts', 'telefono'] },
  '💸 Gastos': { table: 'gastos', cols: ['id', 'fecha', 'categoria', 'concepto', 'monto', 'tipo', 'responsable', 'notas'] }
};
function _supaReadRowsServer(table, cols, order) {
  var c = _supaCfg();
  var sel = cols.join(',');
  var jsonCols = (table === 'cotizaciones') ? { habitaciones: 1, opciones: 1 } : {};
  var all = [], page = 0, size = 1000;
  while (page <= 50) {
    var url = c.url + '/rest/v1/' + table + '?select=' + sel + (order ? ('&order=' + order) : '');
    var resp = UrlFetchApp.fetch(url, { method: 'get', headers: { apikey: c.key, Authorization: 'Bearer ' + c.key, 'Accept-Profile': SUPA_SCHEMA, 'Range-Unit': 'items', 'Range': (page * size) + '-' + (page * size + size - 1) }, muteHttpExceptions: true });
    if (resp.getResponseCode() >= 300) throw new Error('supa ' + resp.getResponseCode());
    var rows = JSON.parse(resp.getContentText() || '[]');
    all = all.concat(rows);
    if (!rows.length || rows.length < size) break;
    page++;
  }
  return all.map(function (o) {
    return cols.map(function (col) {
      var v = o[col];
      if (jsonCols[col]) { return JSON.stringify(v == null ? (col === 'opciones' ? {} : []) : v); }
      return (v == null) ? '' : v;
    });
  });
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
      var sheetName = e.parameter.sheet;
      var range = e.parameter.range;
      if (!sheetName) return jsonResponse({ error: 'Falta el parámetro sheet' });

      // Leer de Supabase si está mapeada y configurada (fuente real de datos)
      if (SUPA_READ_COLS[sheetName] && _supaCfg().ok) {
        try {
          var mc = SUPA_READ_COLS[sheetName];
          var sdata = _supaReadRowsServer(mc.table, mc.cols, mc.order);
          return ContentService.createTextOutput(JSON.stringify(sdata)).setMimeType(ContentService.MimeType.JSON);
        } catch (eSupa) { /* si falla, cae a Sheets abajo */ }
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!range) return jsonResponse({ error: 'Falta el parámetro range' });

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

// =====================================================================
// SUPABASE — Migración inicial (Fase 1) y base para guardado en paralelo
// =====================================================================
// Guarda estas propiedades en: Configuración del proyecto (⚙️) → Propiedades del script
//   SUPABASE_URL          = https://TUPROYECTO.supabase.co   (tu Project URL)
//   SUPABASE_SERVICE_KEY  = <la llave service_role SECRETA>   (NUNCA la pongas en el HTML/chat)
var SUPA_SCHEMA = 'casa_regina';

function _supaCfg() {
  var p = PropertiesService.getScriptProperties();
  var url = (p.getProperty('SUPABASE_URL') || '').replace(/\/+$/, '');
  var key = p.getProperty('SUPABASE_SERVICE_KEY') || '';
  return { url: url, key: key, ok: !!(url && key) };
}

// Prueba de conexión. Ejecútala UNA VEZ desde el editor y mira "Registro de ejecución".
function probarSupabase() {
  var c = _supaCfg();
  if (!c.ok) { Logger.log('❌ Falta SUPABASE_URL o SUPABASE_SERVICE_KEY en Propiedades del script'); return; }
  var resp = UrlFetchApp.fetch(c.url + '/rest/v1/habitaciones?select=id&limit=1', {
    method: 'get',
    headers: { 'apikey': c.key, 'Authorization': 'Bearer ' + c.key, 'Accept-Profile': SUPA_SCHEMA },
    muteHttpExceptions: true
  });
  Logger.log('Código HTTP: ' + resp.getResponseCode());
  Logger.log('Respuesta: ' + resp.getContentText());
  Logger.log(resp.getResponseCode() === 200 ? '✅ Conexión OK con el esquema ' + SUPA_SCHEMA : '❌ Revisa URL / llave / esquema expuesto');
}

// Upsert por llave primaria (o solo insertar si onConflict es null).
function supabaseUpsert(table, rows, onConflict) {
  var c = _supaCfg();
  if (!c.ok) throw new Error('Falta SUPABASE_URL / SUPABASE_SERVICE_KEY');
  if (!rows || !rows.length) return { inserted: 0, dup: 0 };
  // Quitar IDs repetidos (se queda con la PRIMERA aparición). Postgres no permite
  // actualizar la misma fila dos veces en una sola operación de upsert.
  var dup = 0;
  if (onConflict) {
    var seen = {}, uniq = [];
    for (var k = 0; k < rows.length; k++) {
      var kk = String(rows[k][onConflict]);
      if (seen[kk]) { dup++; continue; }
      seen[kk] = 1; uniq.push(rows[k]);
    }
    rows = uniq;
  }
  var inserted = 0, BATCH = 200;
  for (var i = 0; i < rows.length; i += BATCH) {
    var chunk = rows.slice(i, i + BATCH);
    var url = c.url + '/rest/v1/' + table + (onConflict ? ('?on_conflict=' + onConflict) : '');
    var prefer = onConflict ? 'resolution=merge-duplicates,return=minimal' : 'return=minimal';
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'apikey': c.key, 'Authorization': 'Bearer ' + c.key, 'Content-Profile': SUPA_SCHEMA, 'Prefer': prefer },
      payload: JSON.stringify(chunk),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code < 200 || code >= 300) throw new Error(table + ' → HTTP ' + code + ': ' + resp.getContentText());
    inserted += chunk.length;
  }
  return { inserted: inserted, dup: dup };
}

function _sNum(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
function _sInt(v) { var n = parseInt(v, 10); return isNaN(n) ? 0 : n; }
function _sStr(v) { return (v === null || v === undefined) ? '' : String(v); }
function _sJson(v, def) { try { return JSON.parse(v || def); } catch (e) { return JSON.parse(def); } }

function _migrarHoja(nombreHoja, tabla, pk, mapFn) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nombreHoja);
  if (!sh) { Logger.log('— ' + nombreHoja + ': hoja no encontrada, se omite'); return 0; }
  var values = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var obj = mapFn(values[i]);
    if (!obj) continue;
    if (pk && !_sStr(obj[pk]).trim()) continue;
    rows.push(obj);
  }
  var r = supabaseUpsert(tabla, rows, pk);
  Logger.log('✅ ' + nombreHoja + ' → ' + tabla + ': ' + r.inserted + ' fila(s)' +
    (r.dup ? ('  ⚠️ (' + r.dup + ' con id repetido omitidas)') : ''));
  return r.inserted;
}

// =====================================================================
// MAPA ÚNICO Hoja → Tabla de Supabase.
// Lo usan TANTO la migración inicial COMO el guardado en paralelo (doPost),
// así la conversión es siempre idéntica.
//   table  = nombre de la tabla en casa_regina
//   pk     = columna llave (null = sin llave, p.ej. bitácora → solo insertar)
//   toObj  = convierte una fila (arreglo de valores) en el objeto para Supabase
// =====================================================================
var SUPA_MAP = {
  '🏨 Habitaciones': { table: 'habitaciones', pk: 'id', toObj: function (r) {
    return { id: _sStr(r[0]), numero: _sStr(r[1]), tipo: _sStr(r[2]), tarifa: _sNum(r[3]), estado: _sStr(r[4]), notas: _sStr(r[5]) };
  } },
  '📋 Reservas': { table: 'reservas', pk: 'id', toObj: function (r) {
    return { id: _sStr(r[0]), habitacion_num: _sStr(r[1]), huesped: _sStr(r[2]), telefono: _sStr(r[3]), email: _sStr(r[4]),
      fecha_entrada: _toYMD(r[5]), fecha_salida: _toYMD(r[6]), adultos: _sStr(r[7]), ninos: _sStr(r[8]),
      tarifa: _sNum(r[9]), tarifa_total: _sNum(r[10]), estado: _sStr(r[11]), notas: _sStr(r[12]),
      creado_por: _sStr(r[13]), fecha_creacion: _sStr(r[14]) };
  } },
  '📝 Registro': { table: 'registro', pk: 'reg_id', toObj: function (r) {
    return { reg_id: _sStr(r[0]), res_id: _sStr(r[1]), nombre: _sStr(r[2]), habitacion_num: _sStr(r[3]), fecha: _toYMD(r[4]),
      tipo_id: _sStr(r[5]), num_id: _sStr(r[6]), nacionalidad: _sStr(r[7]), procedencia: _sStr(r[8]),
      motivo: _sStr(r[9]), observaciones: _sStr(r[10]), ts: _sStr(r[11]), telefono: _sStr(r[12]) };
  } },
  '🔑 Accesos': { table: 'accesos', pk: 'reserva_id', toObj: function (r) {
    return { reserva_id: _sStr(r[0]), codigo: _sStr(r[1]).replace(/^'/, ''), notas: _sStr(r[2]) };
  } },
  '🧹 Housekeeping': { table: 'housekeeping', pk: 'id', toObj: function (r) {
    return { id: _sStr(r[0]), fecha: _toYMD(r[1]), habitacion_num: _sStr(r[2]), tarea: _sStr(r[3]), prioridad: _sStr(r[4]),
      estado: _sStr(r[5]), asignado_a: _sStr(r[6]), notas: _sStr(r[7]), hora_inicio: _sStr(r[8]),
      hora_fin: _sStr(r[9]), verificado: _sStr(r[10]), checklist: _sStr(r[11]) };
  } },
  '📜 Bitácora': { table: 'bitacora', pk: null, toObj: function (r) {
    if (!_sStr(r[0]) && !_sStr(r[3]) && !_sStr(r[4])) return null;
    return { fecha: _toYMD(r[0]), hora: _sStr(r[1]), usuario: _sStr(r[2]), accion: _sStr(r[3]), detalle: _sStr(r[4]), tipo: _sStr(r[5]) };
  } },
  '👤 Usuarios': { table: 'usuarios', pk: 'id', toObj: function (r) {
    if (!_sStr(r[0])) return null;
    return { id: _sStr(r[0]), usuario: _sStr(r[1]), password: _sStr(r[2]), nombre: _sStr(r[3]), rol: _sStr(r[4]) };
  } },
  '💰 Cobros': { table: 'cobros', pk: 'id', toObj: function (r) {
    return { id: _sStr(r[0]), reserva_id: _sStr(r[1]), fecha: _toYMD(r[2]), hora: _sStr(r[3]), monto: _sNum(r[4]),
      metodo_pago: _sStr(r[5]), concepto: _sStr(r[6]), notas: _sStr(r[7]), responsable: _sStr(r[8]) };
  } },
  '👥 Clientes': { table: 'clientes', pk: 'id', toObj: function (r) {
    return { id: _sStr(r[0]), nombre: _sStr(r[1]), telefono: _sStr(r[2]), email: _sStr(r[3]), notas: _sStr(r[4]),
      primera_visita: _toYMD(r[5]), ultima_visita: _toYMD(r[6]), visitas: _sInt(r[7]), total_gastado: _sNum(r[8]) };
  } },
  '📦 Inventario': { table: 'inventario', pk: 'id', toObj: function (r) {
    if (_sStr(r[1]) === '__DELETED__') return null;
    return { id: _sStr(r[0]), categoria: _sStr(r[1]), producto: _sStr(r[2]), stock: _sInt(r[3]), minimo: _sInt(r[4]),
      unidad: _sStr(r[5]) || 'pzas', notas: _sStr(r[6]), ultima_actualizacion: _sStr(r[7]) };
  } },
  '🍳 Desayunos': { table: 'desayunos', pk: 'codigo', toObj: function (r) {
    return { codigo: _sStr(r[0]), reserva_id: _sStr(r[1]), huesped: _sStr(r[2]), habitacion: _sStr(r[3]), fecha: _toYMD(r[4]),
      personas: _sInt(r[5]) || 1, estado: _sStr(r[6]) || 'Pendiente', validado_por: _sStr(r[7]),
      hora_validacion: _sStr(r[8]), reg_id: _sStr(r[9]) };
  } },
  '📝 Cotizaciones': { table: 'cotizaciones', pk: 'id', toObj: function (r) {
    return { id: _sStr(r[0]), numero: _sStr(r[1]), fecha: _toYMD(r[2]), nombre: _sStr(r[3]), telefono: _sStr(r[4]), email: _sStr(r[5]),
      checkin: _toYMD(r[6]), checkout: _toYMD(r[7]), adultos: _sInt(r[8]) || 1, ninos: _sInt(r[9]),
      habitaciones: _sJson(r[10], '[]'), opciones: _sJson(r[11], '{}'),
      subtotal: _sNum(r[12]), iva: _sNum(r[13]), total: _sNum(r[14]), estado: _sStr(r[15]) || 'Vigente' };
  } },
  '💸 Gastos': { table: 'gastos', pk: 'id', toObj: function (r) {
    return { id: _sStr(r[0]), fecha: _toYMD(r[1]), categoria: _sStr(r[2]) || 'Otros', concepto: _sStr(r[3]),
      monto: _sNum(r[4]), tipo: _sStr(r[5]), responsable: _sStr(r[6]), notas: _sStr(r[7]) };
  } },
  '⚙️ Config': { table: 'config', pk: 'config_key', toObj: function (r) {
    if (!_sStr(r[0])) return null;
    return { config_key: _sStr(r[0]), config_value: _sStr(r[1]) };
  } }
};

// ⭐ MIGRACIÓN COMPLETA — ejecútala UNA VEZ desde el editor.
// Idempotente para tablas con llave (upsert): re-correrla ACTUALIZA, no duplica.
function migrarTodoASupabase() {
  var c = _supaCfg();
  if (!c.ok) { Logger.log('❌ Falta SUPABASE_URL o SUPABASE_SERVICE_KEY en Propiedades del script'); return; }
  var total = 0;
  Object.keys(SUPA_MAP).forEach(function (sheet) {
    var m = SUPA_MAP[sheet];
    total += _migrarHoja(sheet, m.table, m.pk, m.toObj);
  });
  Logger.log('=============================================');
  Logger.log('🎉 MIGRACIÓN TERMINADA — ' + total + ' fila(s) en total');
  Logger.log('Compara los conteos de arriba con tus hojas.');
}

// =====================================================================
// GUARDADO EN PARALELO (Fase 2): espejo de cada escritura hacia Supabase.
// Se llama desde doPost DESPUÉS de escribir en Sheets. Si Supabase falla,
// NO rompe nada: se registra el error y la operación en Sheets sigue válida.
// =====================================================================
function _supaReady() { return _supaCfg().ok; }

// Borra una fila por llave en Supabase.
function supabaseDelete(table, pk, id) {
  if (!pk) return;
  var c = _supaCfg();
  if (!c.ok) return;
  var url = c.url + '/rest/v1/' + table + '?' + pk + '=eq.' + encodeURIComponent(id);
  UrlFetchApp.fetch(url, {
    method: 'delete',
    headers: { 'apikey': c.key, 'Authorization': 'Bearer ' + c.key, 'Content-Profile': SUPA_SCHEMA, 'Prefer': 'return=minimal' },
    muteHttpExceptions: true
  });
}

// Refleja un conjunto de filas (valores posicionales) de una hoja hacia su tabla.
function _supaMirrorRows(sheetName, valuesArray) {
  var m = SUPA_MAP[sheetName];
  if (!m) return; // hoja no mapeada → no se refleja
  var objs = [];
  for (var i = 0; i < valuesArray.length; i++) {
    var o = m.toObj(valuesArray[i]);
    if (o) objs.push(o);
  }
  if (objs.length) supabaseUpsert(m.table, objs, m.pk);
}

// Refleja un borrado por id.
function _supaMirrorDelete(sheetName, id) {
  var m = SUPA_MAP[sheetName];
  if (!m || !m.pk) return;
  supabaseDelete(m.table, m.pk, id);
}

// Refleja un lote 'multi' (append/updateById): agrupa por hoja y hace un upsert por tabla.
function _supaMirrorMulti(ops) {
  var bySheet = {};
  for (var i = 0; i < ops.length; i++) {
    var op = ops[i];
    if (!op || !op.values) continue;
    if (op.type === 'append' || op.type === 'updateById') {
      if (!bySheet[op.sheet]) bySheet[op.sheet] = [];
      bySheet[op.sheet].push(op.values);
    }
  }
  Object.keys(bySheet).forEach(function (sheet) { _supaMirrorRows(sheet, bySheet[sheet]); });
}

// Envuelve el espejo para que NUNCA tumbe la respuesta principal.
function _supaTry(fn) {
  if (!_supaReady()) return;
  try { fn(); } catch (e) { Logger.log('⚠️ Espejo Supabase falló (no crítico): ' + e); }
}

// =====================================================================
// DIAGNÓSTICO DE CLIENTES (solo lectura) — Fase de limpieza
// Ejecútala desde el editor y lee el "Registro de ejecución".
// NO modifica nada: solo analiza la hoja 👥 Clientes.
// =====================================================================
function diagnosticarClientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('👥 Clientes');
  if (!sh) { Logger.log('No se encontró la hoja 👥 Clientes'); return; }
  var vals = sh.getDataRange().getValues();

  function norm(s){ return String(s==null?'':s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/\s+/g,' ').trim(); }

  var totalFilas = 0, idsVacios = 0;
  var byId = {};   // id -> [ {fila, nombre, tel} ]
  var byName = {}; // nombre_norm -> [ {fila, id, tel} ]

  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    var id = String(r[0]==null?'':r[0]).trim();
    var nombre = String(r[1]==null?'':r[1]).trim();
    var tel = String(r[2]==null?'':r[2]).trim();
    // fila "vacía" de verdad: sin id y sin nombre → se ignora
    if (!id && !nombre) continue;
    totalFilas++;
    if (!id) idsVacios++;
    var fila = i + 1;
    if (id) { if (!byId[id]) byId[id] = []; byId[id].push({ fila: fila, nombre: nombre, tel: tel }); }
    var nn = norm(nombre);
    if (nn) { if (!byName[nn]) byName[nn] = []; byName[nn].push({ fila: fila, id: id, tel: tel }); }
  }

  // IDs repetidos
  var idsRepetidos = [], colisiones = [], mismaPersona = [];
  Object.keys(byId).forEach(function(id){
    var filas = byId[id];
    if (filas.length < 2) return;
    idsRepetidos.push(id);
    var nombresDistintos = {};
    filas.forEach(function(f){ nombresDistintos[norm(f.nombre)] = f.nombre; });
    if (Object.keys(nombresDistintos).length > 1) colisiones.push({ id: id, nombres: Object.keys(nombresDistintos).map(function(k){return nombresDistintos[k];}) });
    else mismaPersona.push(id);
  });

  // Nombres repetidos (misma persona en varias filas)
  var nombresRepetidos = [];
  Object.keys(byName).forEach(function(nn){
    var filas = byName[nn];
    if (filas.length < 2) return;
    nombresRepetidos.push({ nombre: nn, filas: filas.length, ids: filas.map(function(f){return f.id || '(vacío)';}) });
  });

  Logger.log('========== DIAGNÓSTICO CLIENTES ==========');
  Logger.log('Filas con datos: ' + totalFilas);
  Logger.log('Filas con ID vacío: ' + idsVacios);
  Logger.log('IDs únicos: ' + Object.keys(byId).length);
  Logger.log('');
  Logger.log('— IDs REPETIDOS —');
  Logger.log('IDs que se repiten: ' + idsRepetidos.length);
  Logger.log('  · misma persona (mismo nombre): ' + mismaPersona.length);
  Logger.log('  · COLISIÓN entre personas distintas: ' + colisiones.length);
  if (colisiones.length) {
    Logger.log('  Ejemplos de colisión (id → nombres distintos):');
    colisiones.slice(0, 8).forEach(function(c){ Logger.log('    ' + c.id + ' → ' + c.nombres.join(' | ')); });
  }
  Logger.log('');
  Logger.log('— NOMBRES REPETIDOS (misma persona, varias filas) —');
  Logger.log('Nombres que aparecen más de una vez: ' + nombresRepetidos.length);
  nombresRepetidos.slice(0, 12).forEach(function(n){
    Logger.log('    "' + n.nombre + '" → ' + n.filas + ' filas (ids: ' + n.ids.join(', ') + ')');
  });
  Logger.log('==========================================');
  Logger.log('Es solo lectura: no se cambió nada.');
}

// =====================================================================
// LIMPIEZA DE CLIENTES — reconstruye una fila por persona (por nombre),
// recalcula estadísticas desde Reservas, reasigna IDs únicos y sincroniza Supabase.
// ⚠️ MODIFICA la hoja 👥 Clientes, PERO crea un respaldo automático antes.
// =====================================================================
function supabaseDeleteAll(table, pk) {
  var c = _supaCfg(); if (!c.ok) return;
  var url = c.url + '/rest/v1/' + table + '?' + pk + '=neq.__none__';
  UrlFetchApp.fetch(url, { method: 'delete',
    headers: { 'apikey': c.key, 'Authorization': 'Bearer ' + c.key, 'Content-Profile': SUPA_SCHEMA, 'Prefer': 'return=minimal' },
    muteHttpExceptions: true });
}

// Convierte cualquier valor de fecha (objeto Date, "YYYY-MM-DD", o texto largo tipo
// "Sat Jan 10 2026 ...") SIEMPRE a "YYYY-MM-DD" (o '' si no se puede).
function _toYMD(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  var d = new Date(s);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function limpiarClientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('👥 Clientes');
  if (!sh) { Logger.log('No se encontró la hoja 👥 Clientes'); return; }
  function norm(s){ return String(s==null?'':s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/\s+/g,' ').trim(); }
  function best(sp){ var b='',bc=-1; Object.keys(sp||{}).forEach(function(k){ if(k && sp[k]>bc){bc=sp[k];b=k;} }); return b; }

  // 1) RESPALDO
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  var backup = sh.copyTo(ss).setName('👥 Clientes_RESPALDO_' + stamp);
  Logger.log('🛟 Respaldo creado en la pestaña: ' + backup.getName());

  // 2) Clientes viejos (notas y respaldo de datos para gente sin reservas)
  var cVals = sh.getDataRange().getValues();
  var oldByName = {};
  for (var i = 1; i < cVals.length; i++) {
    var r = cVals[i]; var nm = norm(r[1]); if (!nm) continue;
    if (!oldByName[nm]) oldByName[nm] = { spellings:{}, telefono:'', email:'', notas:'', visitas:0, total:0, primera:'', ultima:'' };
    var o = oldByName[nm];
    o.spellings[String(r[1]||'')] = (o.spellings[String(r[1]||'')]||0)+1;
    if (!o.telefono && String(r[2]||'').trim()) o.telefono = String(r[2]).trim();
    if (!o.email && String(r[3]||'').trim()) o.email = String(r[3]).trim();
    if (!o.notas && String(r[4]||'').trim()) o.notas = String(r[4]).trim();
    var v = parseInt(r[7])||0, g = parseFloat(r[8])||0;
    if (v > o.visitas) o.visitas = v;
    if (g > o.total) o.total = g;
    var op5 = _toYMD(r[5]), ou6 = _toYMD(r[6]);
    if (op5 && (!o.primera || op5 < o.primera)) o.primera = op5;
    if (ou6 && (!o.ultima || ou6 > o.ultima)) o.ultima = ou6;
  }

  // 3) Estadísticas desde Reservas (fuente de la verdad)
  var rsh = ss.getSheetByName('📋 Reservas');
  var rVals = rsh ? rsh.getDataRange().getValues() : [];
  var resByName = {};
  for (var j = 1; j < rVals.length; j++) {
    var rr = rVals[j]; var nm2 = norm(rr[2]); if (!nm2) continue;
    if (!resByName[nm2]) resByName[nm2] = { spellings:{}, tel:'', telDate:'', email:'', emailDate:'', visitas:0, total:0, primera:'', ultima:'' };
    var o2 = resByName[nm2];
    o2.spellings[String(rr[2]||'')] = (o2.spellings[String(rr[2]||'')]||0)+1;
    var tel = String(rr[3]||'').trim(), email = String(rr[4]||'').trim(), fe = _toYMD(rr[5]);
    var estado = String(rr[11]||''), total = parseFloat(rr[10])||0;
    if (tel && (!o2.telDate || fe >= o2.telDate)) { o2.tel = tel; o2.telDate = fe; }
    if (email && (!o2.emailDate || fe >= o2.emailDate)) { o2.email = email; o2.emailDate = fe; }
    if (estado === 'Check-in' || estado === 'Check-out') o2.visitas++;
    if (estado !== 'Cancelada') o2.total += total;
    if (fe && (!o2.primera || fe < o2.primera)) o2.primera = fe;
    if (fe && (!o2.ultima || fe > o2.ultima)) o2.ultima = fe;
  }

  // 4) Unir personas de ambas fuentes
  var allNames = {};
  Object.keys(oldByName).forEach(function(n){ allNames[n]=1; });
  Object.keys(resByName).forEach(function(n){ allNames[n]=1; });
  var names = Object.keys(allNames).sort();

  // 5) Construir filas limpias
  var out = [], idx = 0;
  names.forEach(function(nm3){
    idx++;
    var rb = resByName[nm3], ob = oldByName[nm3];
    var nombre = (rb && best(rb.spellings)) || (ob && best(ob.spellings)) || nm3;
    var telefono = (rb && rb.tel) || (ob && ob.telefono) || '';
    var email = (rb && rb.email) || (ob && ob.email) || '';
    var notas = (ob && ob.notas) || '';
    var visitas = rb ? rb.visitas : (ob ? ob.visitas : 0);
    var total = rb ? rb.total : (ob ? ob.total : 0);
    var primera = (rb && rb.primera) || (ob && ob.primera) || '';
    var ultima = (rb && rb.ultima) || (ob && ob.ultima) || '';
    var id = 'CL' + String(idx).padStart(4, '0');
    out.push([id, nombre, telefono, email, notas, primera, ultima, visitas, total]);
  });

  // 6) Reescribir la hoja (conservando el encabezado de la fila 1)
  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 9).clearContent();
  if (out.length) {
    // Columnas F y G (primera/última visita) como TEXTO, para que queden YYYY-MM-DD
    // y la app las lea tal cual (evita el error "Invalid time value").
    sh.getRange(2, 6, out.length, 2).setNumberFormat('@');
    sh.getRange(2, 1, out.length, 9).setValues(out);
  }
  Logger.log('✅ Clientes limpios: ' + out.length + ' personas (antes: ' + (cVals.length - 1) + ' filas)');

  // 7) Sincronizar Supabase: borrar los revueltos y subir la lista limpia
  if (_supaReady()) {
    try {
      supabaseDeleteAll('clientes', 'id');
      var objs = out.map(function(r){ return SUPA_MAP['👥 Clientes'].toObj(r); });
      supabaseUpsert('clientes', objs, 'id');
      Logger.log('☁️ Supabase clientes actualizado: ' + objs.length + ' filas');
    } catch (e) { Logger.log('⚠️ No se pudo sincronizar Supabase (reversa con migrarTodoASupabase): ' + e); }
  }
  Logger.log('🛟 Respaldo intacto en: ' + backup.getName() + ' (bórralo cuando confirmes que todo quedó bien)');
}

// =====================================================================
// PANEL DE ADMINISTRADOR — gestión de personal (crear, cambiar PIN, rol, borrar)
// Protegido: solo funciona si quien lo pide es un admin (se valida su sesión).
// La llave service_role nunca sale del backend.
// =====================================================================
function _slug(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, '').substring(0, 20) || 'user';
}
// Devuelve el usuario de Supabase dueño de un token (o null)
function _supaAuthUser(token) {
  var c = _supaCfg(); if (!c.ok || !token) return null;
  var resp = UrlFetchApp.fetch(c.url + '/auth/v1/user', {
    method: 'get', headers: { apikey: c.key, Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) return null;
  try { return JSON.parse(resp.getContentText()); } catch (e) { return null; }
}
// ¿El token pertenece a un ADMIN del hotel? Devuelve su user_id o null.
function _supaIsAdmin(token) {
  var u = _supaAuthUser(token); if (!u || !u.id) return null;
  var c = _supaCfg();
  var resp = UrlFetchApp.fetch(c.url + '/rest/v1/staff?user_id=eq.' + u.id + '&select=rol', {
    headers: { apikey: c.key, Authorization: 'Bearer ' + c.key, 'Accept-Profile': SUPA_SCHEMA }, muteHttpExceptions: true
  });
  try { var rows = JSON.parse(resp.getContentText() || '[]'); if (rows[0] && rows[0].rol === 'admin') return u.id; } catch (e) {}
  return null;
}
// ¿El PIN ya lo usa otra persona? (prueba el PIN contra las demás cuentas)
function _supaPinInUse(pin, exceptEmail) {
  var c = _supaCfg();
  var resp = UrlFetchApp.fetch(c.url + '/rest/v1/staff?select=email', { headers: { apikey: c.key, Authorization: 'Bearer ' + c.key, 'Accept-Profile': SUPA_SCHEMA }, muteHttpExceptions: true });
  var rows = []; try { rows = JSON.parse(resp.getContentText() || '[]'); } catch (e) {}
  for (var i = 0; i < rows.length; i++) {
    var em = rows[i].email; if (!em || em === exceptEmail) continue;
    var r = UrlFetchApp.fetch(c.url + '/auth/v1/token?grant_type=password', { method: 'post', contentType: 'application/json', headers: { apikey: c.key }, payload: JSON.stringify({ email: em, password: pin }), muteHttpExceptions: true });
    if (r.getResponseCode() === 200) { try { if (JSON.parse(r.getContentText()).access_token) return true; } catch (e) {} }
  }
  return false;
}
function _uniqueEmail(base) {
  var c = _supaCfg(); var dom = '@casaregina.mx'; var email = base + dom; var n = 1;
  while (n <= 50) {
    var resp = UrlFetchApp.fetch(c.url + '/rest/v1/staff?select=email&email=eq.' + encodeURIComponent(email), { headers: { apikey: c.key, Authorization: 'Bearer ' + c.key, 'Accept-Profile': SUPA_SCHEMA }, muteHttpExceptions: true });
    var rows = []; try { rows = JSON.parse(resp.getContentText() || '[]'); } catch (e) {}
    if (!rows.length) return email;
    n++; email = base + n + dom;
  }
  return base + Date.now() + dom;
}

function staffAdmin(data) {
  var c = _supaCfg();
  if (!c.ok) return jsonResponse({ success: false, error: 'Backend sin Supabase configurado' });
  var adminId = _supaIsAdmin(data.token);
  if (!adminId) return jsonResponse({ success: false, error: 'No autorizado (solo un admin puede gestionar usuarios)' });
  var op = data.op;

  if (op === 'list') {
    var resp = UrlFetchApp.fetch(c.url + '/rest/v1/staff?select=user_id,usuario,nombre,rol,email&order=nombre', { headers: { apikey: c.key, Authorization: 'Bearer ' + c.key, 'Accept-Profile': SUPA_SCHEMA }, muteHttpExceptions: true });
    var rows = []; try { rows = JSON.parse(resp.getContentText() || '[]'); } catch (e) {}
    return jsonResponse({ success: true, staff: rows });
  }

  if (op === 'create') {
    var nombre = String(data.nombre || '').trim();
    var rol = String(data.rol || '').trim();
    var pin = String(data.pin || '').trim();
    var usuario = String(data.usuario || '').trim() || _slug(nombre);
    if (!nombre || !rol) return jsonResponse({ success: false, error: 'Falta nombre o rol' });
    if (pin.length < 6) return jsonResponse({ success: false, error: 'El PIN debe tener al menos 6 dígitos' });
    if (_supaPinInUse(pin, null)) return jsonResponse({ success: false, error: 'Ese PIN ya lo usa otra persona. Elige uno distinto.' });
    var email = _uniqueEmail(_slug(usuario));
    var create = UrlFetchApp.fetch(c.url + '/auth/v1/admin/users', { method: 'post', contentType: 'application/json', headers: { apikey: c.key, Authorization: 'Bearer ' + c.key }, payload: JSON.stringify({ email: email, password: pin, email_confirm: true }), muteHttpExceptions: true });
    var cc = create.getResponseCode(); var cb = {}; try { cb = JSON.parse(create.getContentText()); } catch (e) {}
    if (cc < 200 || cc >= 300 || !cb.id) return jsonResponse({ success: false, error: 'No se pudo crear la cuenta: ' + (cb.msg || cb.error_description || create.getContentText()) });
    supabaseUpsert('staff', [{ user_id: cb.id, usuario: usuario, nombre: nombre, rol: rol, email: email }], 'user_id');
    return jsonResponse({ success: true, user_id: cb.id, email: email });
  }

  if (op === 'setPin') {
    var uid = String(data.user_id || '').trim();
    var pin2 = String(data.pin || '').trim();
    if (!uid) return jsonResponse({ success: false, error: 'Falta user_id' });
    if (pin2.length < 6) return jsonResponse({ success: false, error: 'El PIN debe tener al menos 6 dígitos' });
    // correo de este usuario (para excluirlo del chequeo de duplicado)
    var er = UrlFetchApp.fetch(c.url + '/rest/v1/staff?user_id=eq.' + uid + '&select=email', { headers: { apikey: c.key, Authorization: 'Bearer ' + c.key, 'Accept-Profile': SUPA_SCHEMA }, muteHttpExceptions: true });
    var erows = []; try { erows = JSON.parse(er.getContentText() || '[]'); } catch (e) {}
    var myEmail = erows[0] && erows[0].email;
    if (_supaPinInUse(pin2, myEmail)) return jsonResponse({ success: false, error: 'Ese PIN ya lo usa otra persona. Elige uno distinto.' });
    var up = UrlFetchApp.fetch(c.url + '/auth/v1/admin/users/' + uid, { method: 'put', contentType: 'application/json', headers: { apikey: c.key, Authorization: 'Bearer ' + c.key }, payload: JSON.stringify({ password: pin2 }), muteHttpExceptions: true });
    if (up.getResponseCode() < 200 || up.getResponseCode() >= 300) return jsonResponse({ success: false, error: 'No se pudo cambiar el PIN: ' + up.getContentText() });
    return jsonResponse({ success: true });
  }

  if (op === 'update') {
    var uid3 = String(data.user_id || '').trim();
    if (!uid3) return jsonResponse({ success: false, error: 'Falta user_id' });
    var row = { user_id: uid3 };
    if (data.nombre != null) row.nombre = String(data.nombre).trim();
    if (data.rol != null) row.rol = String(data.rol).trim();
    if (data.usuario != null) row.usuario = String(data.usuario).trim();
    supabaseUpsert('staff', [row], 'user_id');
    return jsonResponse({ success: true });
  }

  if (op === 'delete') {
    var uid4 = String(data.user_id || '').trim();
    if (!uid4) return jsonResponse({ success: false, error: 'Falta user_id' });
    if (uid4 === adminId) return jsonResponse({ success: false, error: 'No puedes eliminar tu propia cuenta de admin' });
    UrlFetchApp.fetch(c.url + '/auth/v1/admin/users/' + uid4, { method: 'delete', headers: { apikey: c.key, Authorization: 'Bearer ' + c.key }, muteHttpExceptions: true });
    supabaseDelete('staff', 'user_id', uid4);
    return jsonResponse({ success: true });
  }

  return jsonResponse({ success: false, error: 'Operación no reconocida: ' + op });
}
