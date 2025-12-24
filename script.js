function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var rows = sheet.getDataRange().getDisplayValues();
  
  // 1. 讀取行程資料
  var events = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue;
    events.push({
      date: row[0],
      time: row[1],
      title: row[2],
      type: row[3],
      duration: row[4],
      location: row[5],
      id: row[6],
      order: row[7] ? parseFloat(row[7]) : 0
    });
  }
  
  // 2. 讀取試算表名稱作為標題
  var tripTitle = ss.getName();

  // 3. 回傳包裝好的物件
  var response = {
    title: tripTitle,
    data: events
  };
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (action === 'add') {
      var newId = params.id ? params.id : Utilities.getUuid();
      // 新增時，order 預設給一個很大的數字，讓它排在最後面，或者由前端決定
      var newOrder = params.order ? params.order : Date.now(); 
      sheet.appendRow([
        params.date, "'" + params.time, params.title, params.type, params.duration, params.location, newId, newOrder
      ]);
      
    } else if (action === 'edit' || action === 'delete') {
      var data = sheet.getDataRange().getDisplayValues();
      var rowIndex = -1;
      
      // 搜尋 ID (第 7 欄, index 6)
      for (var i = 1; i < data.length; i++) {
        if (data[i][6] == params.id) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex > 0) {
        if (action === 'delete') {
          sheet.deleteRow(rowIndex);
        } else if (action === 'edit') {
          // 這裡非常關鍵：我們只更新傳進來的欄位
          // 為了避免覆蓋掉沒傳的資料，我們先讀取舊資料
          var range = sheet.getRange(rowIndex, 1, 1, 8); // 讀取 8 欄
          var oldValues = range.getValues()[0];
          
          // 如果 params 有值就用 params，沒有就用舊的 (oldValues)
          var newDate = params.date !== undefined ? params.date : oldValues[0];
          var newTime = params.time !== undefined ? "'" + params.time : (oldValues[1] ? "'" + oldValues[1] : "");
          var newTitle = params.title !== undefined ? params.title : oldValues[2];
          var newType = params.type !== undefined ? params.type : oldValues[3];
          var newDuration = params.duration !== undefined ? params.duration : oldValues[4];
          var newLoc = params.location !== undefined ? params.location : oldValues[5];
          // ID (oldValues[6]) 不變
          var newOrder = params.order !== undefined ? params.order : oldValues[7];

          range.setValues([[newDate, newTime, newTitle, newType, newDuration, newLoc, oldValues[6], newOrder]]);
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', msg: e.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
