/**
 * 退会アンケート → 匿名化 → Discord 投稿（Google Apps Script）
 *
 * スプレッドシートの「拡張機能 → Apps Script」に貼り付けて使う。
 * Discord の Webhook URL は「プロジェクトの設定 → スクリプト プロパティ」に
 *   DISCORD_WEBHOOK_URL = https://discord.com/api/webhooks/....
 * として登録する（コードやフォームには書かない）。
 */

const CONFIG = {
  // 'batch'  : 回答をため、1日1回まとめて順番をシャッフルして投稿（退会直後に投稿されて誰か特定されるのを防ぐ・推奨）
  // 'instant': 回答が届いたらすぐ投稿
  MODE: 'batch',
  BATCH_HOUR: 21,          // まとめ投稿の時刻（日本時間）
  BATCH_MIN_ITEMS: 1,      // これ未満しか溜まっていなければ次回に持ち越す（2 にすると1件だけの投稿を避けられる）
  SHEET_NAME: '回答',
  BOT_NAME: '退会アンケート',
  TITLE: '退会アンケート（匿名）',
  EMBED_COLOR: 0xE2A52B,
  MAX_MESSAGE: 5000,
  MAX_ANSWERS: 15,
  MAX_VALUE: 300,
  MIN_ELAPSED_MS: 3000,    // これより速い送信はボットとみなして捨てる
};

// 集計シートでグラフにしない列（自由記述）
const DASH_SKIP = ['差し支えなければ、他に参加されているコミュニティーを教えて頂けないでしょうか？'];
const DASH_SHEET = '集計';

const COL = { DATE: 1, ANON: 2, STATUS: 3, MESSAGE: 4, FIRST_Q: 5 };
const BASE_HEADERS = ['受付日', '匿名ID', '状態', '想い'];

/* ---------------- 受付 ---------------- */

function doGet() {
  return out({ ok: true, message: '退会アンケート受付中' });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // ボット対策：見えない入力欄に何か入っている／速すぎる送信は成功扱いで捨てる
    if (body.hp) return out({ ok: true });
    if (typeof body.elapsedMs === 'number' && body.elapsedMs < CONFIG.MIN_ELAPSED_MS) return out({ ok: true });

    const answers = (Array.isArray(body.answers) ? body.answers : [])
      .slice(0, CONFIG.MAX_ANSWERS)
      .map(a => ({
        title: clip(anonymizeText(a && a.title), 100),
        value: clip(anonymizeText(Array.isArray(a && a.value) ? a.value.join(' / ') : a && a.value), CONFIG.MAX_VALUE),
      }))
      .filter(a => a.title);
    const message = clip(anonymizeText(body.message), CONFIG.MAX_MESSAGE);

    if (!answers.some(a => a.value) && !message) return out({ ok: false, error: 'empty' });

    const rec = { anonId: makeAnonId(), answers, message };
    const row = appendRecord(rec);

    if (CONFIG.MODE === 'instant') {
      try {
        postRecord(rec);
        getSheet().getRange(row, COL.STATUS).setValue('投稿済');
      } catch (err) {
        console.error(err); // 未投稿のまま残るので flushQueue で再送できる
      }
    }
    return out({ ok: true });
  } catch (err) {
    console.error(err);
    return out({ ok: false, error: 'server' });
  }
}

/* ---------------- 匿名化 ---------------- */

/** 本文から個人を特定しうる情報を伏せ字にする（フォーム側のプレビューと同じルール） */
function anonymizeText(s) {
  return String(s == null ? '' : s)
    .replace(/\r\n?/g, '\n')
    .replace(/<@[!&]?\d+>/g, '[メンション省略]')
    .replace(/<#\d+>/g, '[チャンネル省略]')
    .replace(/https?:\/\/\S+/gi, '[URL省略]')
    .replace(/discord\.gg\/\S+/gi, '[URL省略]')
    .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, '[メール省略]')
    .replace(/\b0x[a-fA-F0-9]{40}\b/g, '[ウォレット省略]')
    .replace(/(?<!\d)\d{17,20}(?!\d)/g, '[ID省略]')
    .replace(/(?<!\d)0\d{1,4}-\d{1,4}-\d{3,4}(?!\d)/g, '[電話番号省略]')
    .replace(/(?<!\d)0[5-9]0\d{8}(?!\d)/g, '[電話番号省略]')
    .replace(/[^\s#＃]{2,32}[#＃]\d{4}\b/g, '[ユーザー名省略]')
    .replace(/[@＠][^\s@＠、。,.!?！？「」()（）]{1,32}/g, '@匿名')
    .trim();
}

/** 投稿ごとに使い捨ての匿名ID（誰とも紐づかない乱数） */
function makeAnonId() {
  return '匿名メンバー #' + Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
}

/* ---------------- スプレッドシート ---------------- */

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_NAME);
    sh.appendRow(BASE_HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function appendRecord(rec) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet();
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    rec.answers.forEach(a => {
      if (headers.indexOf(a.title) === -1) {
        headers.push(a.title);
        sh.getRange(1, headers.length).setValue(a.title);
      }
    });
    const row = headers.map(() => '');
    // 正確な時刻は残さず「日付だけ」記録する
    row[COL.DATE - 1] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    row[COL.ANON - 1] = rec.anonId;
    row[COL.STATUS - 1] = '未投稿';
    row[COL.MESSAGE - 1] = safeCell(rec.message);
    rec.answers.forEach(a => { row[headers.indexOf(a.title)] = safeCell(a.value); });
    sh.appendRow(row);
    return sh.getLastRow();
  } finally {
    lock.releaseLock();
  }
}

/** まとめ投稿：未投稿の回答をシャッフルして Discord へ（時間トリガーから呼ばれる） */
function dailyJob() {
  if (CONFIG.MODE === 'batch') flushQueue();
  buildDashboard();
}

function flushQueue() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let sh, values, pending;
  try {
    sh = getSheet();
    const last = sh.getLastRow();
    if (last < 2) return;
    values = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
    pending = [];
    for (let i = 1; i < values.length; i++) {
      if (values[i][COL.STATUS - 1] === '未投稿') pending.push(i + 1);
    }
    if (pending.length < CONFIG.BATCH_MIN_ITEMS) return;
    pending.forEach(r => sh.getRange(r, COL.STATUS).setValue('投稿中'));
  } finally {
    lock.releaseLock();
  }

  const headers = values[0].map(String);
  shuffle(pending).forEach(r => {
    const v = values[r - 1];
    const rec = {
      anonId: String(v[COL.ANON - 1]),
      message: String(v[COL.MESSAGE - 1] || ''),
      answers: headers.slice(COL.FIRST_Q - 1).map((h, j) => ({ title: h, value: String(v[COL.FIRST_Q - 1 + j] || '') })),
    };
    try {
      postRecord(rec);
      sh.getRange(r, COL.STATUS).setValue('投稿済');
    } catch (err) {
      console.error(err);
      sh.getRange(r, COL.STATUS).setValue('未投稿');
    }
    Utilities.sleep(1500);
  });
}

/* ---------------- Discord ---------------- */

function postRecord(rec) {
  const url = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK_URL');
  if (!url) throw new Error('スクリプトプロパティ DISCORD_WEBHOOK_URL が未設定です');

  const fields = rec.answers
    .filter(a => a.value)
    .map(a => ({ name: clip(a.title, 256), value: clip(a.value, 1024), inline: a.value.length <= 24 }));
  const fieldChars = fields.reduce((n, f) => n + f.name.length + f.value.length, 0);

  // 1 つの embed は合計 6000 字まで。1 通目に入りきらない想いは続きとして分けて送る
  const firstLen = Math.max(0, Math.min(3800, 5500 - fieldChars));
  const msg = rec.message || '';
  const chunks = msg ? [msg.slice(0, firstLen)].concat(chunk(msg.slice(firstLen), 3800)).filter(Boolean) : [];
  const total = chunks.length;

  send(url, {
    embeds: [{
      title: '📮 ' + CONFIG.TITLE,
      color: CONFIG.EMBED_COLOR,
      author: { name: rec.anonId },
      fields,
      description: chunks[0] ? '**💬 想い**\n' + chunks[0] : '（自由記述なし）',
      footer: { text: total > 1 ? `想い 1/${total} ・投稿者は自動で匿名化されています` : '投稿者は自動で匿名化されています' },
    }],
  });
  for (let i = 1; i < total; i++) {
    Utilities.sleep(800);
    send(url, {
      embeds: [{
        color: CONFIG.EMBED_COLOR,
        author: { name: rec.anonId + '（続き）' },
        description: chunks[i],
        footer: { text: `想い ${i + 1}/${total}` },
      }],
    });
  }
}

function send(url, payload, retried) {
  payload.username = CONFIG.BOT_NAME;
  payload.allowed_mentions = { parse: [] }; // @everyone 等で誰かを呼び出さない
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code === 429 && !retried) {
    const wait = (JSON.parse(res.getContentText() || '{}').retry_after || 2) * 1000;
    Utilities.sleep(Math.min(wait + 200, 30000));
    return send(url, payload, true);
  }
  if (code >= 300) throw new Error('Discord ' + code + ': ' + res.getContentText());
}

/* ---------------- 管理メニュー・集計 ---------------- */

/** スプレッドシートを開くと上部に「退会アンケート」メニューが出る */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('退会アンケート')
    .addItem('📊 集計グラフを更新', 'buildDashboard')
    .addSeparator()
    .addItem('Discord にテスト投稿', 'testDiscord')
    .addItem('溜まっている回答を今すぐ投稿', 'flushQueue')
    .addItem('毎日の自動処理を設定（最初に1回）', 'setupBatchTrigger')
    .addToUi();
}

/** 「回答」シートを数えて「集計」シートに表とグラフを作り直す */
function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = getSheet();
  const last = src.getLastRow();
  const values = last >= 2 ? src.getRange(1, 1, last, src.getLastColumn()).getValues() : [src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0]];
  const headers = values[0].map(String);
  const rows = values.slice(1);

  let sh = ss.getSheetByName(DASH_SHEET);
  if (!sh) sh = ss.insertSheet(DASH_SHEET, 0);
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.clear();
  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 70);

  sh.getRange(1, 1, 1, 2).setValues([['退会アンケート 集計', '']]).setFontWeight('bold').setFontSize(14);
  sh.getRange(2, 1, 1, 2).setValues([['回答数（合計）', rows.length]]);
  sh.getRange(3, 1).setValue('更新：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '（メニュー「退会アンケート → 集計グラフを更新」で最新に）').setFontColor('#888888');

  let r = 5;
  const addBlock = (title, pairs, type) => {
    if (!pairs.length) return;
    sh.getRange(r, 1, 1, 2).setValues([[title, '件数']]).setFontWeight('bold').setBackground('#FCF3DE');
    sh.getRange(r + 1, 1, pairs.length, 2).setValues(pairs);
    const chart = sh.newChart()
      .setChartType(type)
      .addRange(sh.getRange(r, 1, pairs.length + 1, 2))
      .setNumHeaders(1)
      .setPosition(r, 4, 0, 0)
      .setOption('title', title)
      .setOption('legend', { position: type === Charts.ChartType.PIE ? 'right' : 'none' })
      .setOption('colors', ['#E2A52B'])
      .setOption('width', 620)
      .setOption('height', Math.max(240, pairs.length * 34 + 80))
      .build();
    sh.insertChart(chart);
    r += Math.max(pairs.length + 3, Math.ceil(Math.max(240, pairs.length * 34 + 80) / 21) + 2);
  };

  // 月ごとの回答数
  const months = {};
  rows.forEach(v => {
    const d = v[COL.DATE - 1];
    const m = d instanceof Date ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM') : String(d).slice(0, 7);
    if (m) months[m] = (months[m] || 0) + 1;
  });
  addBlock('月ごとの回答数', Object.keys(months).sort().map(k => [k, months[k]]), Charts.ChartType.COLUMN);

  // 選択式の質問ごと
  headers.slice(COL.FIRST_Q - 1).forEach((h, j) => {
    if (!h || DASH_SKIP.indexOf(h) !== -1) return;
    const counts = {};
    rows.forEach(v => {
      String(v[COL.FIRST_Q - 1 + j] || '').split(' / ').map(x => x.trim()).filter(Boolean).forEach(x => {
        const key = x.indexOf('その他：') === 0 ? 'その他' : x;
        counts[key] = (counts[key] || 0) + 1;
      });
    });
    let pairs = Object.keys(counts).map(k => [k, counts[k]]);
    if (/満足度/.test(h)) pairs.sort((a, b) => String(b[0]).localeCompare(String(a[0])));
    else pairs.sort((a, b) => b[1] - a[1]);
    // 1 件ずつしかない自由記述っぽい列は表だけ（グラフにしない）
    if (pairs.length > 15) return;
    addBlock(h, pairs, pairs.length <= 5 ? Charts.ChartType.PIE : Charts.ChartType.BAR);
  });

  // 他に参加しているコミュニティー（自由記述）は一覧で
  DASH_SKIP.forEach(h => {
    const j = headers.indexOf(h);
    if (j === -1) return;
    const list = rows.map(v => String(v[j] || '').trim()).filter(Boolean);
    sh.getRange(r, 1).setValue('他に参加しているコミュニティー（自由記述）').setFontWeight('bold').setBackground('#FCF3DE');
    if (list.length) sh.getRange(r + 1, 1, list.length, 1).setValues(list.map(x => [x]));
    r += list.length + 3;
  });
}

/* ---------------- 初期設定用（エディタから手動で実行） ---------------- */

/** Discord にテスト投稿する（最初の権限承認もこれで行う） */
function testDiscord() {
  postRecord({
    anonId: makeAnonId(),
    answers: [
      { title: '在籍期間', value: '3〜6ヶ月' },
      { title: '退会の理由', value: '料金が高い / 時間がとれない' },
      { title: 'いくらなら続けていましたか？', value: '3,000〜3,999円' },
    ],
    message: anonymizeText('これはテスト投稿です。@taro さんや https://example.com は自動で伏せ字になります。'),
  });
}

/** 毎日の自動処理（まとめ投稿＋集計グラフ更新）のトリガーを作る。最初に 1 回だけ実行 */
function setupBatchTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['flushQueue', 'dailyJob'].indexOf(t.getHandlerFunction()) !== -1)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(CONFIG.BATCH_HOUR).inTimezone('Asia/Tokyo').create();
}

/* ---------------- 小物 ---------------- */

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function clip(s, n) {
  s = String(s == null ? '' : s);
  return s.length > n ? s.slice(0, n - 1) + '…' : s;
}
function chunk(s, n) {
  const res = [];
  for (let i = 0; i < s.length; i += n) res.push(s.slice(i, i + n));
  return res;
}
function shuffle(a) {
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}
/** スプレッドシートで数式として解釈されないようにする */
function safeCell(v) {
  v = String(v == null ? '' : v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
