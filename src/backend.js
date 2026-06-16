// ==========================================
// CONFIG
// ==========================================
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const GEMINI_API_URL = `https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;

const SHEET_NAMES = {
  MEMBER:        'メンバー',
  HISTORY:       '成績履歴',
  SCHEDULE:      '設定',
  OFFICIAL_HDCP: '公式HDCP履歴',
  ACCESS_LOG:    'アクセスログ',
};

const SESSION_TTL_SECONDS = 21600; // 6時間
const MAX_RETRIES         = 3;
const RETRY_BASE_MS       = 1000;
const LOGIN_FAIL_LIMIT    = 5;     // ユーザー単位の連続失敗上限
const GLOBAL_FAIL_LIMIT   = 15;    // システム全体の連続失敗上限
const LOGIN_LOCK_SECONDS  = 600;   // ロック時間 10分


// ==========================================
// UTILS
// ==========================================
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function validateSession(token) {
  if (!token) return null;
  const cached = CacheService.getScriptCache().get(token);
  return cached ? JSON.parse(cached) : null;
}

function sha256(text) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8
  );
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/** 名前の空白を除去して正規化（比較用） */
function normalizeName(v) {
  return String(v).replace(/\s+/g, '');
}

/** パスワードハッシュの形式検証（64文字の16進数かチェック） */
function isValidHash(hash) {
  return typeof hash === 'string' && /^[a-f0-9]{64}$/i.test(hash);
}

/** アクセスログ書き込み */
function writeAccessLog(name, event, result, note) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAMES.ACCESS_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAMES.ACCESS_LOG);
      sheet.appendRow(['日時', '名前', 'イベント', '結果', '備考']);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 120);
      sheet.setColumnWidth(3, 80);
      sheet.setColumnWidth(4, 80);
      sheet.setColumnWidth(5, 200);
    }
    const now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
    sheet.appendRow([now, name, event, result, note || '-']);
  } catch (e) {
    Logger.log('ログ書き込みエラー: ' + e.toString());
  }
}

/** ログイン失敗カウント管理（ユーザー単位 + グローバル2段構え） */
function checkAndIncrementFailCount(id) {
  const cache          = CacheService.getScriptCache();
  const globalLockKey  = 'login_lock_global';
  const globalCountKey = 'login_fail_global';
  const userLockKey    = `login_lock_${id}`;
  const userCountKey   = `login_fail_${id}`;

  // ロックチェック
  if (cache.get(globalLockKey)) {
    throw new Error('システムへの異常なアクセスを検知したため、一時的にログインを制限しています。10分後に再試行してください。');
  }
  if (cache.get(userLockKey)) {
    throw new Error('アカウントが一時ロックされています。10分後に再試行してください。');
  }

  // カウントアップ
  const userCount   = Number(cache.get(userCountKey)   || 0) + 1;
  const globalCount = Number(cache.get(globalCountKey) || 0) + 1;

  if (globalCount >= GLOBAL_FAIL_LIMIT) {
    cache.put(globalLockKey, '1', LOGIN_LOCK_SECONDS);
    cache.remove(globalCountKey);
    throw new Error('システムへの異常なアクセスを検知したため、一時的にログインを制限しています。10分後に再試行してください。');
  }

  if (userCount >= LOGIN_FAIL_LIMIT) {
    cache.put(userLockKey, '1', LOGIN_LOCK_SECONDS);
    cache.remove(userCountKey);
    throw new Error('ログイン試行回数が上限を超えました。10分後に再試行してください。');
  }

  cache.put(userCountKey,   String(userCount),   LOGIN_LOCK_SECONDS);
  cache.put(globalCountKey, String(globalCount), LOGIN_LOCK_SECONDS);
}

function clearFailCount(id) {
  const cache = CacheService.getScriptCache();
  cache.remove(`login_fail_${id}`);
  cache.remove(`login_lock_${id}`);
  cache.remove('login_fail_global'); // 正規ログイン成功でグローバルカウントも緩和
}


// ==========================================
// CACHE MANAGEMENT
// ==========================================
function clearDataCache() {
  const cache = CacheService.getScriptCache();
  // 個別削除ではなく配列で一括削除
  cache.removeAll([
    'cache_score_history', 
    'cache_official_hdcp_history', 
    'cache_member_list',
    'cache_dashboard_data'
  ]);
}

function forceClearCache() {
  clearDataCache();
  return { success: true, message: 'システムキャッシュを正常に破棄しました。' };
}


// ==========================================
// ROUTER
// ==========================================
const PUBLIC_ACTIONS = {
  checkLogin: (req) => checkLogin(req.id, req.password),
};

const PRIVATE_ACTIONS = {
  getScoreHistory:           ()             => getScoreHistory(),
  getAllOfficialHdcpHistory: ()             => getAllOfficialHdcpHistory(), 
  getMemberList:             ()             => getMemberList(),
  changeMyPassword:          (req)          => changeMyPassword(req),
  logout:                    (req)          => logout(req.token),
};

const ADMIN_ACTIONS = {
  uploadAndProcessImage: (req)          => uploadAndProcessImage(req.base64Data, req.mimeType),
  resetMemberPassword:   (req)          => resetMemberPassword(req.targetId, req.newPassword),
  clearSystemCache:      (req)          => forceClearCache(),
  toggleMaintenance:     (req)          => toggleMaintenanceMode(req.state),
};

// メンテナンスモードの切替処理
function toggleMaintenanceMode(stateStr) {
  // stateStr は 'true' または 'false' が文字列として渡される想定
  PropertiesService.getScriptProperties().setProperty('MAINTENANCE_MODE', stateStr);
  
  // 状態が変わったためキャッシュも念のため破棄
  clearDataCache(); 
  
  const isMaintenance = stateStr === 'true';
  writeAccessLog('SYSTEM', 'MAINTENANCE', isMaintenance ? 'ON' : 'OFF');
  return { success: true, maintenanceMode: isMaintenance };
}

function doPost(e) {
  try {
    const req    = JSON.parse(e.postData.contents);
    const action = req.action;

    if (PUBLIC_ACTIONS[action]) return jsonResponse(PUBLIC_ACTIONS[action](req));

    const session = validateSession(req.token);
    // セッション切れの場合はフロントに強制ログアウトを指示
    if (!session) return jsonResponse({ success: false, error: 'セッションが無効または期限切れです。', forceLogout: true });

    // すでにログイン済みのユーザーに対するメンテナンスモード強制シャットアウト
    if (action !== 'logout') {
      const isMaintenance = PropertiesService.getScriptProperties().getProperty('MAINTENANCE_MODE') === 'true';
      if (isMaintenance && session.role !== '管理者') {
        // 管理者以外は強制ログアウトフラグを立てて弾く
        return jsonResponse({ success: false, error: 'システムがメンテナンスモードに入りました。', forceLogout: true });
      }
    }

    if (ADMIN_ACTIONS[action]) {
      if (session.role !== '管理者') throw new Error('管理者権限がありません。');
      return jsonResponse(ADMIN_ACTIONS[action](req, session));
    }
    if (PRIVATE_ACTIONS[action]) return jsonResponse(PRIVATE_ACTIONS[action](req, session));

    throw new Error('未定義のアクション: ' + action);
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ==========================================
// AUTH & DASHBOARD DATA  ログイン処理を軽くするための共通ダッシュボードデータ取得 （キャッシュ対応）
// ==========================================
function getCommonDashboardData() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('cache_dashboard_data');
  if (cachedData) return JSON.parse(cachedData);

  const futureEvents = [];
  const schedSheet   = getSheet(SHEET_NAMES.SCHEDULE);
  if (schedSheet) {
    const sData       = schedSheet.getDataRange().getValues();
    const currentTime = new Date().getTime();
    const days        = ['日', '月', '火', '水', '木', '金', '土'];
    for (let j = 1; j < sData.length; j++) {
      const dVal = sData[j][0];
      if (!(dVal instanceof Date)) continue;
      const threshold = new Date(dVal.getTime());
      threshold.setHours(10, 0, 0, 0);
      if (threshold.getTime() <= currentTime) continue;

      const tVal = sData[j][1];
      const cVal = sData[j][2];
      let timeStr = '時間未定';
      if (tVal) {
        if (tVal instanceof Date) {
          timeStr = `${tVal.getHours()}時${String(tVal.getMinutes()).padStart(2, '0')}分スタート`;
        } else {
          const tStr = String(tVal).trim();
          timeStr = (tStr.includes('スタート') || tStr.includes('未定')) ? tStr : `${tStr}スタート`;
        }
      }
      futureEvents.push({
        rawDate:  dVal.getTime(),
        dateText: `${dVal.getMonth() + 1}月${dVal.getDate()}日（${days[dVal.getDay()]}曜日）`,
        timeText: timeStr,
        clubCompe: cVal ? String(cVal).trim() : '',
      });
    }
    futureEvents.sort((a, b) => a.rawDate - b.rawDate);
  }

  let eagleUpdateDateStr = '未確定';
  const histSheet = getSheet(SHEET_NAMES.HISTORY);
  if (histSheet) {
    const hData = histSheet.getDataRange().getValues();
    if (hData.length > 1) {
      const lastDate = new Date(hData[hData.length - 1][0]);
      if (!isNaN(lastDate)) eagleUpdateDateStr = Utilities.formatDate(lastDate, 'JST', 'yyyy/MM/dd');
    }
  }

  let officialUpdateDateStr = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/01');
  const ohSheet = getSheet(SHEET_NAMES.OFFICIAL_HDCP);
  if (ohSheet) {
    const ohData = ohSheet.getDataRange().getValues();
    if (ohData.length > 1) {
      const lastDate = new Date(ohData[ohData.length - 1][0]);
      if (!isNaN(lastDate)) officialUpdateDateStr = Utilities.formatDate(lastDate, 'JST', 'yyyy/MM/dd');
    }
  }

  const result = {
    nextCompe: futureEvents[0] || null,
    futureSchedules: futureEvents.slice(1),
    eagleUpdateDateStr,
    officialUpdateDateStr
  };
  
  cache.put('cache_dashboard_data', JSON.stringify(result), SESSION_TTL_SECONDS);
  return result;
}

function checkLogin(id, password) {
  try {
    checkAndIncrementFailCount(id);
  } catch (lockErr) {
    const isGlobal = lockErr.message.includes('システムへ');
    writeAccessLog(id, 'LOGIN', 'LOCKED', isGlobal ? 'システム全体ロック' : '試行回数上限');
    return { success: false, error: lockErr.message };
  }

  const memberSheet = getSheet(SHEET_NAMES.MEMBER);
  if (!memberSheet) return { success: false, error: 'システムエラー' };

  const memberData = memberSheet.getDataRange().getValues();
  const inputHash  = password; // 受信値はすでにSHA-256ハッシュ済み

  for (let i = 1; i < memberData.length; i++) {
    if (memberData[i][0] !== id || memberData[i][1] !== inputHash) continue;

    clearFailCount(id);

    const name         = memberData[i][2];
    const eagleHdcp    = memberData[i][3];
    const officialHdcp = memberData[i][4];
    const role         = memberData[i][5];
    // 管理者ログイン時のメール通知機能
   if (role === '管理者') {
      try {
        const adminEmail = 'wadaijp@gmail.com'; 
        
        const timeStr = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
        const subject = '【イーグル会】管理者ログイン通知';
        const body = `イーグル会システムに管理者権限でのログインがありました。\n\n`
                   + `・日時: ${timeStr}\n`
                   + `・ログインID: ${id}\n`
                   + `・ユーザー名: ${name}\n\n`
                   + `※もしご自身でのログインでない場合、直ちにスプレッドシートから対象者のパスワードを変更してください。`;
        
        MailApp.sendEmail(adminEmail, subject, body);
      } catch (e) {
        console.error('メール送信エラー: ', e);
      }
    }
    const eagleTee     = memberData[i][6] ? String(memberData[i][6]).trim() : '未設定';
    const genderRaw    = memberData[i][7] ? String(memberData[i][7]).trim() : '男';
    const gender       = genderRaw === '女' ? 'female' : 'male';

    const isMaintenance = PropertiesService.getScriptProperties().getProperty('MAINTENANCE_MODE') === 'true';
    if (isMaintenance && role !== '管理者') {
      writeAccessLog(name, 'LOGIN', 'BLOCKED', 'メンテナンス中');
      return { success: false, error: '現在システムメンテナンス中です。終了までしばらくお待ちください。' };
    }

    writeAccessLog(name, 'LOGIN', 'SUCCESS');

    // 全員共通のダッシュボードデータ（スケジュール・更新日）をキャッシュ関数から取得
    const dashboard = getCommonDashboardData();

    const token = Utilities.getUuid();
    CacheService.getScriptCache().put(
      token, JSON.stringify({ name, role }), SESSION_TTL_SECONDS
    );

    return {
      success: true, name, role, token, eagleHdcp, officialHdcp, eagleTee, gender,
      eagleUpdateDate:    dashboard.eagleUpdateDateStr,
      officialUpdateDate: dashboard.officialUpdateDateStr,
      nextCompe:          dashboard.nextCompe,
      futureSchedules:    dashboard.futureSchedules,
      maintenanceMode:    isMaintenance,
    };
  }

  const cache          = CacheService.getScriptCache();
  const failCount      = cache.get(`login_fail_${id}`) || cache.get('login_fail_global') || '1';
  writeAccessLog(id, 'LOGIN', 'FAILED', `試行回数:${failCount}`);
  return { success: false, error: 'IDまたはパスワードが違います' };
}

function logout(token) {
  const session = token ? validateSession(token) : null;
  const name    = session ? session.name : '不明';
  writeAccessLog(name, 'LOGOUT', 'SUCCESS');
  if (token) CacheService.getScriptCache().remove(token);
  return { success: true };
}

function resetMemberPassword(targetId, newPassword) {
  if (!targetId || !newPassword) {
    return { success: false, error: 'IDと新パスワードは必須です。' };
  }
  
  // API直接アクセスによる不正な文字列（1文字など）の登録を防ぐ
  if (!isValidHash(newPassword)) {
    return { success: false, error: 'パスワードの処理に失敗しました。アプリを再読み込みしてもう一度お試しください。' };
  }
  
  const sheet = getSheet(SHEET_NAMES.MEMBER);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizeName(data[i][0]) !== normalizeName(targetId)) continue;
    sheet.getRange(i + 1, 2).setValue(newPassword);
    return { success: true, message: `${data[i][2]} さんのPWをリセットしました。` };
  }
  return { success: false, error: `ID "${targetId}" が見つかりません。` };
}

// ==========================================
// OCR & データ登録
// ==========================================
function uploadAndProcessImage(base64Data, mimeType) {
  try {
    const memberData  = getSheet(SHEET_NAMES.MEMBER).getDataRange().getValues();
    const memberNames = memberData.slice(1).map(r => String(r[2]).trim()).filter(Boolean);
    const todayStr    = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');

    const prompt = `添付画像はゴルフコンペの成績表です。以下のJSON配列のみを返却してください。
【重要】「名前」は以下のリストと照合・補正してください。見つからない場合のみ画像のまま出力。
【リスト】${memberNames.join(', ')}
【日付】画像右上付近の開催日（例:2025年9月23日）を「YYYY/MM/DD」に変換。なければ本日の日付（${todayStr}）を使用。
【フォーマット】[{"日付":"YYYY/MM/DD","順位":1,"名前":"〇〇 〇〇","OUT":45,"IN":42,"GROSS":87,"HDCP":14,"NET":73}]`;

    const payload  = { contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mimeType, data: base64Data } }] }] };
    const response = UrlFetchApp.fetch(GEMINI_API_URL, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true,
    });

    const replyText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
    const scoreData = JSON.parse(replyText.replace(/^```json\s*/i, '').replace(/```\s*$/, '').trim());

    const historySheet = getSheet(SHEET_NAMES.HISTORY);
    const rows = scoreData.map(r => [r['日付'], r['順位'], r['名前'], r['OUT'], r['IN'], r['GROSS'], r['HDCP'], r['NET']]);
    historySheet.getRange(historySheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);

    updateMemberHandicaps(scoreData);
    
    // 成績データ更新後にシステム全体のキャッシュを破棄
    clearDataCache();
    
    return { success: true, count: scoreData.length };
  } catch (e) {
    const msg = e.toString();
    let userMsg = 'データ処理エラーが発生しました。';
    if (msg.includes('quota')) userMsg = 'API利用上限に達しました。しばらく待って再試行してください。';
    else if (msg.includes('JSON')) userMsg = '画像の解析結果が不正でした。別の画像で再試行してください。';
    return { success: false, error: userMsg, detail: msg };
  }
}

function updateMemberHandicaps(scoreData) {
  const memberSheet = getSheet(SHEET_NAMES.MEMBER);
  const memberData  = memberSheet.getDataRange().getValues();

  const compePlayers = {};
  scoreData.forEach(r => {
    const name = normalizeName(r['名前']);
    if (name) compePlayers[name] = { rank: Number(r['順位']), hdcp: Number(r['HDCP']) };
  });

  const newHdcpValues = [];
  for (let i = 1; i < memberData.length; i++) {
    const name = normalizeName(memberData[i][2]);
    let newHdcp = !isNaN(memberData[i][3]) ? Number(memberData[i][3])
                : (!isNaN(memberData[i][4]) ? Number(memberData[i][4]) : 0);

    if (compePlayers[name]) {
      const p   = compePlayers[name];
      const cut = p.rank === 1 ? 0.80 : p.rank === 2 ? 0.90 : p.rank === 3 ? 0.95 : 1.00;
      newHdcp   = Math.round(p.hdcp * cut * 10) / 10;
    }
    newHdcpValues.push([newHdcp]);
  }
  if (newHdcpValues.length > 0) {
    memberSheet.getRange(2, 4, newHdcpValues.length, 1).setValues(newHdcpValues);
  }
}

// ==========================================
// 公式HDCP PDF一括処理
// ==========================================
function processOfficialHandicapPDF() {
  const folderName        = 'イーグル会公式HDCP';
  const archiveFolderName = '処理済み_公式HDCP';

  const getOrCreate = (name) =>
    DriveApp.getFoldersByName(name).hasNext()
      ? DriveApp.getFoldersByName(name).next()
      : DriveApp.createFolder(name);

  const folder        = getOrCreate(folderName);
  const archiveFolder = getOrCreate(archiveFolderName);
  const files         = folder.getFilesByType(MimeType.PDF);
  if (!files.hasNext()) return Logger.log('解析対象のPDFファイルがありません。');

  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet   = getSheet(SHEET_NAMES.OFFICIAL_HDCP);
  if (!historySheet) {
    historySheet = ss.insertSheet(SHEET_NAMES.OFFICIAL_HDCP);
    historySheet.appendRow(['日付', '名前', '公式HDCP']);
  }

  const memberSheet = getSheet(SHEET_NAMES.MEMBER);
  const memberData  = memberSheet.getDataRange().getValues();
  const memberNames = memberData.slice(1).map(r => String(r[2]).trim()).filter(Boolean);

  const startTime = Date.now();
  while (files.hasNext()) {
    if (Date.now() - startTime > 255000) {
      const msg = '実行時間制限が迫ったため処理を中断。残りは次回実行してください。';
      Logger.log('⚠️ ' + msg);
      writeAccessLog('SYSTEM', 'PDF_PROCESS', 'INTERRUPTED', msg);
      break;
    }

    const file   = files.next();
    const prompt = `添付PDFはゴルフ倶楽部のハンディインデックス表です。以下のJSONのみを出力してください。
1. 右上付近の開催日を「YYYY/MM/DD」形式で抽出
2. 【リスト】の全員のインデックスを抽出（カンマはピリオドに修正）。いない場合は"不明"とする。
【リスト】${memberNames.join(', ')}
【出力】{"日付":"YYYY/MM/DD","データ":[{"名前":"〇〇","インデックス":16.7}]}`;

    const payload = {
      contents: [{ parts: [
        { text: prompt },
        { inline_data: { mime_type: MimeType.PDF, data: Utilities.base64Encode(file.getBlob().getBytes()) } },
      ]}],
    };
    const options = {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true,
    };

    let resultObj = null;
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      try {
        const res  = UrlFetchApp.fetch(GEMINI_API_URL, options);
        const json = JSON.parse(res.getContentText());
        if (!json.error && json.candidates) {
          resultObj = JSON.parse(json.candidates[0].content.parts[0].text.replace(/^```json\s*/i, '').replace(/```\s*$/, '').trim());
          break;
        }
      } catch (e) { /* retry */ }
      Utilities.sleep(Math.min(Math.pow(2, attempt) * RETRY_BASE_MS, 30000));
    }

    if (!resultObj || !resultObj['日付']) {
      Logger.log('❌ 解析失敗スキップ: ' + file.getName());
      continue;
    }

    const hdcpMap = {};
    resultObj['データ'].forEach(item => {
      hdcpMap[normalizeName(item['名前'])] = item['インデックス'];
    });

    const rows = [];
    memberData.slice(1).forEach(r => {
      const key = normalizeName(r[2]);
      if (hdcpMap[key] !== undefined) rows.push([resultObj['日付'], r[2], hdcpMap[key]]);
    });
    if (rows.length > 0) {
      historySheet.getRange(historySheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    }

    file.moveTo(archiveFolder);
    Utilities.sleep(2000);
  }

  const latestMap = {};
  historySheet.getDataRange().getValues().slice(1).forEach(r => {
    const d = new Date(r[0]);
    const n = normalizeName(r[1]);
    if (!isNaN(d) && (!latestMap[n] || d > latestMap[n].date)) {
      latestMap[n] = { date: d, value: r[2] };
    }
  });

  const syncValues = memberData.slice(1).map(r => {
    const n = normalizeName(r[2]);
    return [latestMap[n] !== undefined ? latestMap[n].value : r[4]];
  });
  if (syncValues.length > 0) {
    memberSheet.getRange(2, 5, syncValues.length, 1).setValues(syncValues);
  }
  
  // 公式HDCP更新後にシステム全体のキャッシュを破棄
  clearDataCache();
  
  Logger.log('✅ 公式HDCPの全同期完了');
}

// ==========================================
// フロントエンド用データ取得 (キャッシュ対応済)
// ==========================================
function getScoreHistory() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('cache_score_history');
  if (cachedData) return JSON.parse(cachedData);

  try {
    const data = getSheet(SHEET_NAMES.HISTORY).getDataRange().getValues();
    if (data.length <= 1) return [];
    const headers = data[0];
    const result = data.slice(1).map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = (r[i] instanceof Date)
          ? Utilities.formatDate(r[i], 'JST', 'yyyy/MM/dd')
          : r[i];
      });
      return obj;
    });
    // エラーがなく正常に取得できた場合のみキャッシュに保存
    cache.put('cache_score_history', JSON.stringify(result), SESSION_TTL_SECONDS);
    return result;
  } catch { return []; }
}

function getAllOfficialHdcpHistory() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('cache_official_hdcp_history');
  if (cachedData) return JSON.parse(cachedData);

  try {
    const data      = getSheet(SHEET_NAMES.OFFICIAL_HDCP).getDataRange().getValues();
    const threshold = new Date();
    threshold.setFullYear(threshold.getFullYear() - 1);

    const result = data.slice(1)
      .filter(r => {
        const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
        return !isNaN(d) && d >= threshold;
      })
      .map(r => ({
        date: Utilities.formatDate(r[0] instanceof Date ? r[0] : new Date(r[0]), 'JST', 'yyyy/MM/dd'),
        name: normalizeName(r[1]),
        hdcp: r[2],
      }))
      .sort((a, b) => new Date(a.date) - new Date(b.date));
      
    // エラーがなく正常に取得できた場合のみキャッシュに保存
    cache.put('cache_official_hdcp_history', JSON.stringify(result), SESSION_TTL_SECONDS);
    return result;
  } catch { return []; }
}

function getMemberList() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('cache_member_list');
  if (cachedData) return JSON.parse(cachedData);

  try {
    const data = getSheet(SHEET_NAMES.MEMBER).getDataRange().getValues();
    const result = data.slice(1)
      .filter(r => String(r[2]).trim())
      .map(r => {
        const eagleHdcpRaw = parseFloat(r[3]); 
        const hdcp = parseFloat(r[4]);         
        return { 
          name: String(r[2]).trim(), 
          eagleHdcp: isNaN(eagleHdcpRaw) ? '-' : eagleHdcpRaw,
          hdcp: isNaN(hdcp) ? 999 : hdcp 
        };
      })
      .sort((a, b) => a.hdcp - b.hdcp);
      
    // エラーがなく正常に取得できた場合のみキャッシュに保存
    cache.put('cache_member_list', JSON.stringify(result), SESSION_TTL_SECONDS);
    return result;
  } catch { return []; }
}

function changeMyPassword(req) {
  try {
    if (!req.currentPassword || !req.newPassword) {
      return { success: false, error: 'パスワード情報が不足しています。' };
    }

    // API直接アクセスによる不正な文字列の登録を防ぐ
    if (!isValidHash(req.currentPassword) || !isValidHash(req.newPassword)) {
      return { success: false, error: 'パスワードの処理に失敗しました。アプリを再読み込みしてもう一度お試しください。' };
    }

    const sheet = getSheet(SHEET_NAMES.MEMBER);
    const data = sheet.getDataRange().getValues();
    const id = req.loginId;
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id) {
        if (String(data[i][1]) !== req.currentPassword) {
          return { success: false, error: '現在のパスワードが間違っています。' };
        }
        sheet.getRange(i + 1, 2).setValue(req.newPassword);
        return { success: true, message: 'パスワードを変更しました。' };
      }
    }
    return { success: false, error: 'ユーザーが見つかりません。' };
  } catch (e) {
    return { success: false, error: 'システムエラーが発生しました。' };
  }
}

// ==========================================
// SYSTEM CHECK TOOLS (管理者・開発者用)
// ==========================================
function checkCacheSize() {
  const data = getSheet(SHEET_NAMES.HISTORY).getDataRange().getValues();
  // JSON文字列化した際のバイト数を概算計算
  const sizeBytes = Utilities.newBlob(JSON.stringify(data)).getBytes().length;
  Logger.log(`成績履歴JSONサイズ: ${sizeBytes} bytes (CacheService上限: 約100,000 bytes)`);
}