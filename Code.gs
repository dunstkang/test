// ===== 設定 =====
const INVITE_CODE = '';          // 空字串 = 任何人都能註冊；填入文字 = 註冊時必須輸入這個邀請碼
const SESSION_SECONDS = 21600;   // 登入有效時間 6 小時（CacheService 上限），每次操作會自動延長
const MAX_FAILS = 5;             // 連續輸錯幾次密碼後暫時鎖定
const LOCK_SECONDS = 900;        // 鎖定 15 分鐘
const LIST_LIMIT = 200;          // 畫面最多顯示最近幾筆
const HASH_ROUNDS = 200;         // 雜湊重複次數（增加破解難度）

// ===== 試算表 =====
// 資料分頁：第一個「不叫 users」的分頁，不論它叫什麼名字
function logSheet() {
  return SpreadsheetApp.getActive().getSheets().find(s => s.getName() !== 'users');
}

// 帳號分頁：叫 users，不存在時第一次註冊會自動建立
function userSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('users');
  if (!sh) {
    sh = ss.insertSheet('users', ss.getNumSheets());
    sh.appendRow(['account', 'salt', 'hash', 'created']);
  }
  return sh;
}

// ===== 密碼雜湊（SHA-256 + 每人獨立 salt）=====
function hashPw(salt, pw) {
  let h = salt + pw;
  for (let i = 0; i < HASH_ROUNDS; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8)
      .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  }
  return h;
}

function findUser(account) {
  const rows = userSheet().getDataRange().getValues();
  const key = account.toLowerCase();
  for (let i = 1; i < rows.length; i++) {           // 第 1 列是標題
    if (String(rows[i][0]).toLowerCase() === key) {
      return { account: String(rows[i][0]), salt: String(rows[i][1]), hash: String(rows[i][2]) };
    }
  }
  return null;
}

// ===== 登入狀態 =====
function newSession(account) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('s_' + token, account, SESSION_SECONDS);
  return { ok: true, token: token, account: account };
}

function auth(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const account = cache.get('s_' + token);
  if (!account) return null;
  cache.put('s_' + token, account, SESSION_SECONDS);   // 有操作就延長
  return account;
}

// ===== 動作 =====
function register(req) {
  const account = String(req.account || '').trim();
  const password = String(req.password || '');
  if (!/^[A-Za-z0-9_\u4e00-\u9fff]{2,20}$/.test(account)) {
    return { error: '帳號需為 2～20 個字，只能用中英文、數字、底線' };
  }
  if (password.length < 6 || password.length > 64) return { error: '密碼需為 6～64 個字元' };
  if (INVITE_CODE && String(req.invite || '') !== INVITE_CODE) return { error: '邀請碼錯誤' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);               // 避免兩人同時註冊同一個帳號
  try {
    if (findUser(account)) return { error: '這個帳號已經有人使用' };
    const salt = Utilities.getUuid();
    userSheet().appendRow([account, salt, hashPw(salt, password), new Date()]);
  } finally {
    lock.releaseLock();
  }
  return newSession(account);
}

function login(req) {
  const account = String(req.account || '').trim();
  const password = String(req.password || '');
  const cache = CacheService.getScriptCache();
  const failKey = 'f_' + account.toLowerCase();
  if (Number(cache.get(failKey) || 0) >= MAX_FAILS) {
    return { error: '輸入錯誤次數過多，請 15 分鐘後再試' };
  }
  const user = account ? findUser(account) : null;
  if (!user || hashPw(user.salt, password) !== user.hash) {
    cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), LOCK_SECONDS);
    return { error: '帳號或密碼錯誤' };
  }
  cache.remove(failKey);
  return newSession(user.account);
}

function list(req) {
  if (!auth(req.token)) return { error: 'unauthorized' };
  const sh = logSheet();
  const last = sh.getLastRow();
  if (last < 1) return { ok: true, rows: [] };
  const start = Math.max(1, last - LIST_LIMIT + 1);
  const vals = sh.getRange(start, 1, last - start + 1, 3).getValues();
  const rows = vals
    .filter(r => r[1] !== '' || r[2] !== '')
    .map(r => ({ t: r[0], id: r[1], text: r[2] }));
  return { ok: true, rows: rows };
}

function add(req) {
  const account = auth(req.token);
  if (!account) return { error: 'unauthorized' };
  const text = String(req.text || '').trim();
  if (!text) return { error: '請輸入內容' };
  if (text.length > 500) return { error: '內容最多 500 字' };
  const safe = /^[=+\-@]/.test(text) ? "'" + text : text;   // 避免被試算表當成公式
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    logSheet().appendRow([new Date(), account, safe]);       // id 由後端依登入帳號寫入，無法冒充
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
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
      case 'register': return out(register(req));
      case 'login':    return out(login(req));
      case 'list':     return out(list(req));
      case 'add':      return out(add(req));
      case 'logout':
        CacheService.getScriptCache().remove('s_' + req.token);
        return out({ ok: true });
      default:         return out({ error: '未知的動作' });
    }
  } catch (err) {
    console.error(err);
    return out({ error: '伺服器發生錯誤' });
  }
}

// 直接用瀏覽器開網址只會看到這行，不再洩漏任何資料
function doGet() {
  return out({ ok: true, message: 'API is running' });
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
