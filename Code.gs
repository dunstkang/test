// 工作表名稱（要與試算表左下角的分頁名稱完全一致）
const SHEET_NAME = 'my_pwa_test_log';

function getSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('找不到工作表：' + SHEET_NAME);
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);               // 避免多人同時寫入互相覆蓋
  try {
    const { id, text } = JSON.parse(e.postData.contents);
    getSheet().appendRow([new Date(), id, text]);
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  const rows = getSheet().getDataRange().getValues();
  return out(rows.map(r => ({ t: r[0], id: r[1], text: r[2] })));
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
