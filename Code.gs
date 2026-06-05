var RECORD_SHEET = '電工工作紀錄';
var AREA_SHEET   = '區域清單';
var CONFIG_SHEET = '人員設定';

function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : '';

  var cb = e.parameter.callback || '';

  function wrap(obj) {
    var json = JSON.stringify(obj);
    return ContentService
      .createTextOutput(cb ? cb + '(' + json + ')' : json)
      .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }

  if (action === 'config') {
    return wrap(getConfig());
  }

  if (action === 'submit') {
    try {
      var data = JSON.parse(decodeURIComponent(e.parameter.data));
      submitRecord(data);
      return wrap({ success: true });
    } catch (err) {
      return wrap({ success: false, error: err.message });
    }
  }

  return wrap({ status: 'ok' });
}

function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var areas = [];
  var areaSheet = ss.getSheetByName(AREA_SHEET);
  if (areaSheet) {
    areaSheet.getDataRange().getValues().forEach(function(row) {
      row.forEach(function(cell) {
        if (cell) areas.push(String(cell).trim());
      });
    });
  }

  var departments = [], electricians = [], equipments = [];
  var configSheet = ss.getSheetByName(CONFIG_SHEET);
  if (configSheet) {
    configSheet.getDataRange().getValues().forEach(function(row) {
      if (row[0]) departments.push(String(row[0]).trim());
      if (row[1]) electricians.push(String(row[1]).trim());
      if (row[2]) equipments.push(String(row[2]).trim());
    });
  }

  return { areas: areas, departments: departments, electricians: electricians, equipments: equipments };
}

function submitRecord(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RECORD_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(RECORD_SHEET);
    sheet.appendRow([
      '時間戳記','部門','填表人','設備名稱','區域','設備號碼',
      '報修日期時間','維修人員','工作類型','完成日期時間','結果','備註說明'
    ]);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(),
    data.department,
    data.filler,
    data.equipment,
    data.area,
    data.equipmentNo,
    data.reportTime,
    data.electrician,
    data.workType,
    data.completionTime,
    data.result,
    data.remarks
  ]);
}
