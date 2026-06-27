/**
 * Google Apps Script for All in One POS - sync 2 chiều.
 *
 * Nếu sheet vẫn trắng sau khi đồng bộ:
 * - Copy ID của Google Sheet trong URL và dán vào SPREADSHEET_ID bên dưới.
 *   Ví dụ URL: https://docs.google.com/spreadsheets/d/ABC123/edit
 *   ID là: ABC123
 */
var SPREADSHEET_ID = '1dFBhs7rzX6jksY-61sCHfG2M9eGRnXVckK4LiqcGCpU'; // Khuyến nghị: dán ID Google Sheet vào đây nếu Web App không ghi dữ liệu.

function doPost(e) {
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var ss = getSpreadsheet_();
    writeStore_(ss, data);
    writeProducts_(ss, data.products || []);
    writeInvoices_(ss, data.invoices || []);
    writeInvoiceItems_(ss, data.invoices || []);
    return json_({ ok: true, message: 'Da ghi du lieu', syncedAt: new Date().toISOString(), spreadsheetUrl: ss.getUrl() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet(e) {
  try {
    var ss = getSpreadsheet_();
    if (e && e.parameter && e.parameter.action === 'ping') {
      var pingData = { ok: true, message: 'Ket noi Apps Script thanh cong', spreadsheetUrl: ss.getUrl(), sheets: ss.getSheets().map(function(s) { return s.getName(); }) };
      return e.parameter.callback ? jsonp_(e.parameter.callback, pingData) : json_(pingData);
    }
    var data = readAll_(ss);
    return e && e.parameter && e.parameter.callback ? jsonp_(e.parameter.callback, data) : json_(data);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim()) {
    return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Khong tim thay Google Sheet. Hay dien SPREADSHEET_ID trong Apps Script.');
  return ss;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(callback, obj) {
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(obj) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function sheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#172226').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

function writeStore_(ss, data) {
  var sh = sheet_(ss, 'Cửa hàng', ['Tên', 'Địa chỉ', 'Điện thoại', 'Đồng bộ lúc']);
  var s = data.store || {};
  sh.getRange(2, 1, 1, 4).setValues([[s.name || '', s.address || '', s.phone || '', data.syncedAt || '']]);
}

function writeProducts_(ss, products) {
  var headers = ['ID', 'Tên sản phẩm', 'Danh mục', 'Giá mặc định', 'Giá vốn', 'Tồn kho', 'Cảnh báo', 'Theo IMEI', 'IMEI/SN còn tồn'];
  var sh = sheet_(ss, 'Tồn kho', headers);
  var rows = products.map(function(p) {
    var serials = (p.serialItems || []).map(function(x) { return x.code + (x.price != null ? ' | ' + x.price : ''); }).join('\n');
    return [p.id, p.name, p.category, p.price, p.cost, p.stock, p.min, p.trackSerials ? 'Có' : 'Không', serials];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.autoResizeColumns(1, headers.length);
}

function writeInvoices_(ss, invoices) {
  var headers = ['ID', 'Mã hóa đơn', 'Ngày lập', 'Cập nhật', 'Khách hàng', 'Điện thoại', 'Đơn vị giao hàng', 'Mã vận đơn', 'Trạng thái giao hàng', 'Ghi chú giao hàng', 'Cập nhật giao hàng', 'Tạm tính', 'Giảm giá', 'Thuế suất', 'Tiền thuế', 'Tổng cộng', 'JSON items'];
  var sh = sheet_(ss, 'Hóa đơn', headers);
  var rows = invoices.map(function(i) {
    var sh = i.shipping || {};return [i.id, i.code, i.date, i.updatedAt || '', (i.customer || {}).name || '', (i.customer || {}).phone || '', sh.carrier || '', sh.tracking || '', sh.status || 'Chưa giao', sh.note || '', sh.updatedAt || '', i.subtotal, i.discount, i.taxRate, i.tax, i.total, JSON.stringify(i.items || [])];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.autoResizeColumns(1, headers.length);
}

function writeInvoiceItems_(ss, invoices) {
  var headers = ['Mã hóa đơn', 'Sản phẩm', 'IMEI/SN', 'Số lượng', 'Đơn giá', 'Thành tiền'];
  var sh = sheet_(ss, 'Chi tiết hóa đơn', headers);
  var rows = [];
  invoices.forEach(function(inv) {
    (inv.items || []).forEach(function(item) {
      if (item.serialItems && item.serialItems.length) {
        item.serialItems.forEach(function(serial) { rows.push([inv.code, item.name, serial.code, 1, serial.price, serial.price]); });
      } else {
        rows.push([inv.code, item.name, (item.serials || []).join(', '), item.qty, item.price, (item.price || 0) * (item.qty || 0)]);
      }
    });
  });
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.autoResizeColumns(1, headers.length);
}

function readAll_(ss) {
  return { ok: true, store: readStore_(ss), products: readProducts_(ss), invoices: readInvoices_(ss) };
}

function values_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  return values.slice(1).filter(function(row) { return row.some(function(v) { return v !== ''; }); }).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function readStore_(ss) {
  var rows = values_(ss, 'Cửa hàng');
  var r = rows[0] || {};
  return { name: r['Tên'] || r['Ten'] || '', address: r['Địa chỉ'] || r['Dia chi'] || '', phone: r['Điện thoại'] || r['Dien thoai'] || '' };
}

function readProducts_(ss) {
  return values_(ss, 'Tồn kho').map(function(r) {
    var serialItems = parseSerialItems_(r['IMEI/SN còn tồn'] || r['IMEI/SN con ton'], Number(r['Giá mặc định'] || r['Gia mac dinh']) || 0);
    return {
      id: String(r['ID'] || Utilities.getUuid()),
      name: String(r['Tên sản phẩm'] || ''),
      category: String(r['Danh mục'] || ''),
      price: Number(r['Giá mặc định']) || 0,
      cost: Number(r['Giá vốn']) || 0,
      stock: serialItems.length || Number(r['Tồn kho']) || 0,
      min: valueOrBlank_(r['Cảnh báo'] || r['Canh bao']),
      trackSerials: String(r['Theo IMEI']).toLowerCase() === 'co' || serialItems.length > 0,
      serialItems: serialItems,
      serials: serialItems.map(function(x) { return x.code; }),
      serial: serialItems[0] ? serialItems[0].code : ''
    };
  });
}

function readInvoices_(ss) {
  return values_(ss, 'Hóa đơn').map(function(r) {
    var items = [];
    try { items = JSON.parse(r['JSON items'] || '[]'); } catch (err) { items = []; }
    return {
      id: String(r['ID'] || Utilities.getUuid()),
      code: String(r['Mã hóa đơn'] || r['Ma hoa don'] || ''),
      date: iso_(r['Ngày lập'] || r['Ngay lap']),
      updatedAt: (r['Cập nhật'] || r['Cap nhat']) ? iso_(r['Cập nhật'] || r['Cap nhat']) : '',
      customer: { name: String(r['Khách hàng'] || r['Khach hang'] || ''), phone: String(r['Điện thoại'] || r['Dien thoai'] || '') },
      shipping: { carrier: String(r['Đơn vị giao hàng'] || ''), tracking: String(r['Mã vận đơn'] || ''), status: String(r['Trạng thái giao hàng'] || 'Chưa giao'), note: String(r['Ghi chú giao hàng'] || ''), updatedAt: r['Cập nhật giao hàng'] ? iso_(r['Cập nhật giao hàng']) : '' },
      subtotal: Number(r['Tạm tính'] || r['Tam tinh']) || 0,
      discount: Number(r['Giảm giá'] || r['Giam gia']) || 0,
      taxRate: Number(r['Thuế suất'] || r['Thue suat']) || 0,
      tax: Number(r['Tiền thuế'] || r['Tien thue']) || 0,
      total: Number(r['Tổng cộng'] || r['Tong cong']) || 0,
      items: items
    };
  });
}

function valueOrBlank_(value) {
  if (value === '' || value == null) return '';
  var n = Number(value);
  return isNaN(n) ? '' : n;
}

function parseSerialItems_(text, defaultPrice) {
  if (!text) return [];
  var seen = {};
  return String(text).split(/\r?\n/).map(function(line) {
    line = line.trim();
    if (!line) return null;
    var parts = line.split('|');
    var code = parts[0].trim();
    var price = parts.length > 1 ? Number(String(parts.slice(1).join('|')).replace(/[^0-9.]/g, '')) : defaultPrice;
    if (!code || seen[code.toLowerCase()]) return null;
    seen[code.toLowerCase()] = true;
    return { code: code, price: price || defaultPrice || 0 };
  }).filter(Boolean);
}

function iso_(value) {
  if (!value) return new Date().toISOString();
  if (Object.prototype.toString.call(value) === '[object Date]') return value.toISOString();
  return new Date(value).toISOString();
}
