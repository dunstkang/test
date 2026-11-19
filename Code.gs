// ===== 設定 =====
const VERSION = 3;               // 前端會檢查這個數字，用來判斷後端是不是舊版
const MIN_PW = 4;                // 密碼最短字元數（留空代表不使用密碼）
const MAX_FAILS = 5;             // 連續輸錯幾次密碼後暫時鎖定
const LOCK_SECONDS = 900;        // 鎖定 15 分鐘
const LIST_LIMIT = 200;          // 畫面最多顯示最近幾筆
const HASH_ROUNDS = 200;         // 雜湊重複次數（增加破解難度）

// ===== 試算表 =====
// 資料分頁：第一個「不叫 users」的分頁，不論它叫什麼名字
function logSheet() {
  return SpreadsheetApp.getActive().getSheets().find(s => s.getName() !== 'users');
}

// id 分頁：叫 users，不存在時會自動建立
// 欄位：id、salt、密碼雜湊、建立時間（雜湊是空的 = 這個 id 沒有設定密碼）
function userSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('users');
  if (!sh) {
    sh = ss.insertSheet('users', ss.getNumSheets());
    sh.appendRow(['id', 'salt', 'hash', 'created']);
  }
  return sh;
}

// ===== 密碼雜湊（SHA-256 + 每個 id 獨立 salt）=====
function hashPw(salt, pw) {
  let h = salt + pw;
  for (let i = 0; i < HASH_ROUNDS; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8)
      .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  }
  return h;
}

function findUser(id) {
  const rows = userSheet().getDataRange().getValues();
  const key = id.toLowerCase();
  for (let i = 1; i < rows.length; i++) {           // 第 1 列是標題
    if (String(rows[i][0]).toLowerCase() === key) {
      return { row: i + 1, id: String(rows[i][0]), salt: String(rows[i][1]), hash: String(rows[i][2]) };
    }
  }
  return null;
}

// ===== 動作 =====
function list() {
  const sh = logSheet();
  const last = sh.getLastRow();
  let rows = [];
  if (last >= 1) {
    const start = Math.max(1, last - LIST_LIMIT + 1);
    rows = sh.getRange(start, 1, last - start + 1, 3).getValues()
      .filter(r => r[1] !== '' || r[2] !== '')
      .map(r => ({ t: r[0], id: r[1], text: r[2] }));
  }
  const users = userSheet().getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '')
    .map(r => ({ id: String(r[0]), locked: String(r[2]) !== '' }));
  return { ok: true, rows: rows, users: users };
}

function add(req) {
  const id = String(req.id || '').trim();
  const pw = String(req.password || '');
  const text = String(req.text || '').trim();

  if (!/^[A-Za-z0-9_\u4e00-\u9fff]{1,20}$/.test(id)) {
    return { error: 'id 需為 1～20 個字，只能用中英文、數字、底線' };
  }
  if (pw && (pw.length < MIN_PW || pw.length > 64)) {
    return { error: '密碼需為 ' + MIN_PW + '～64 個字元（或留空表示不使用密碼）' };
  }
  if (!text) return { error: '請輸入內容' };
  if (text.length > 500) return { error: '內容最多 500 字' };

  const cache = CacheService.getScriptCache();
  const failKey = 'f_' + id.toLowerCase();
  if (Number(cache.get(failKey) || 0) >= MAX_FAILS) {
    return { error: '密碼輸入錯誤次數過多，請 15 分鐘後再試' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);               // 避免多人同時寫入互相覆蓋
  try {
    let account = id, created = false, pwSet = false;
    const user = findUser(id);

    if (!user) {
      // 沒看過的 id → 自動註冊
      const salt = pw ? Utilities.getUuid() : '';
      userSheet().appendRow([id, salt, pw ? hashPw(salt, pw) : '', new Date()]);
      created = true;
      pwSet = !!pw;
    } else {
      account = user.id;
      if (user.hash) {
        // 這個 id 已有密碼 → 必須對
        if (hashPw(user.salt, pw) !== user.hash) {
          cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), LOCK_SECONDS);
          return { error: '密碼錯誤' };
        }
        cache.remove(failKey);
      } else if (pw) {
        // 這個 id 還沒有密碼，這次第一次給密碼 → 從此固定
        const salt = Utilities.getUuid();
        userSheet().getRange(user.row, 2, 1, 2).setValues([[salt, hashPw(salt, pw)]]);
        pwSet = true;
      }
    }

    const safe = /^[=+\-@]/.test(text) ? "'" + text : text;   // 避免被試算表當成公式
    logSheet().appendRow([new Date(), account, safe]);
    return { ok: true, created: created, pwSet: pwSet };
  } finally {
    lock.releaseLock();
  }
}

// ===== 入口 =====
function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return out({ error: '請求格式錯誤' });
  }
  try {
    switch (req.action) {
      case 'list': return out(list());
      case 'add':  return out(add(req));
      default:     return out({ error: '未知的動作' });
    }
  } catch (err) {
    console.error(err);
    return out({ error: '伺服器發生錯誤' });
  }
}

// 用瀏覽器直接開網址，看到 version 是 3 就代表新版已部署
function doGet() {
  return out({ ok: true, message: 'API is running' });
}

function out(obj) {
  obj.v = VERSION;
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
