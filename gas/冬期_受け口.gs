/**
 * 冬期講習 個人予定の受け口（Google Apps Script）
 *
 * student.html から次を受け付けます。
 *   GET  ?action=ping              動作確認
 *   POST {action:"get", t:<生徒番号>, name:<氏名>}    番号と氏名が名簿と合えば、学年と提出済みの内容を返す
 *   POST {action:"save", t:<生徒番号>, name:<氏名>, ...}  合えば予定を保存する（同じ番号なら上書き）
 * 氏名はURLに載せないため、どちらもPOSTで受けます。
 *
 * このスクリプトは、回答を溜めるスプレッドシートの「拡張機能 > Apps Script」に貼ります。
 * シートは「名簿」と「回答」の2枚。最初に一度だけ 初期設定() を実行してください。
 * 生徒番号・氏名・回答はシートにだけ置き、このファイルには入れません（リポジトリは公開です）。
 */

const CONFIG = {
  DEADLINE: '2026-12-05T23:59:59+09:00', // 提出の締切。過ぎても保存でき、「回答」シートの「締切後」列に印が付く
  WEEKS: 3,
  DOWS: [1, 2, 3, 4, 5, 6],               // 月〜土
  KOMAS: [1, 2, 3, 4, 5, 6, 7, 8],
  SUBJECTS: ['英', '国', '数', '理', '社', '小論文', 'その他'],
  MEMO_MAX: 500,
  MAX_BODY: 30000,
  ROSTER: '名簿',   // 列: 合言葉（=生徒番号） / 生徒ID / 氏名 / 学年
  ANSWERS: '回答',   // 生徒ごとの最新の1行（提出し直すと上書き）
  HISTORY: '履歴',   // 提出のたびに1行足す（上書きしない）
  OVERVIEW: '一覧',  // 全員の予定を1枚の表にしたもの。提出のたびに作り直す（手で書き込まない）
  START: '2026-12-20', // 1週目の日曜日（日付の見出し用）
};
const ANSWER_HEADERS = ['合言葉', '氏名', '学年', '提出日時', '科目', '予定(JSON)', '補足', '保護者確認', '提出回数', '締切後'];
const HISTORY_HEADERS = ['提出日時', '合言葉', '氏名', '学年', '科目', '予定(JSON)', '補足', '保護者確認', '締切後'];
const ROSTER_HEADERS = ['合言葉', '生徒ID', '氏名', '学年'];
const TOKEN_RE = /^[0-9]{4,10}$/;   // 生徒番号（数字のみ）
const TEACHER = '講師';   // 名簿の「学年」がこの値なら講師（科目と保護者確認は要らない）
const MARK_VALUES = ['x', 't', 'q'];   // x=×  t=△  q=？

/* ---------- 初期設定（最初に一度だけ実行） ---------- */
function 初期設定() {
  const ss = SpreadsheetApp.getActive();
  [[CONFIG.ROSTER, ROSTER_HEADERS], [CONFIG.ANSWERS, ANSWER_HEADERS], [CONFIG.HISTORY, HISTORY_HEADERS]].forEach(function (pair) {
    const sh = ss.getSheetByName(pair[0]) || ss.insertSheet(pair[0]);
    sh.getRange(1, 1, 1, pair[1].length).setValues([pair[1]]);
    sh.setFrozenRows(1);
    // 合言葉・日時などを数値や日付に変えられないよう、列を文字列にしておく
    sh.getRange(1, 1, sh.getMaxRows(), pair[1].length).setNumberFormat('@');
  });
}

/* ---------- 受け付ける処理 ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'ping') return json_({ ok: true });
  return json_({ ok: false, error: 'bad_request' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const raw = (e && e.postData && e.postData.contents) || '';
    if (raw.length > CONFIG.MAX_BODY) return json_({ ok: false, error: 'too_large' });
    let body;
    try { body = JSON.parse(raw); } catch (_) { return json_({ ok: false, error: 'bad_request' }); }
    if (!body || typeof body !== 'object') return json_({ ok: false, error: 'bad_request' });
    const token = normalize_(body.t);
    const member = TOKEN_RE.test(token) ? findMember_(token, body.name) : null;
    if (!member) return json_({ ok: false, error: 'not_found' });

    const now = new Date();
    if (body.action === 'get') {
      const ans = findAnswer_(token);
      const out = { ok: true, name: member.name, grade: member.grade, role: member.grade === TEACHER ? 'teacher' : 'student', data: null };
      if (ans) { out.submittedAt = ans.submittedAt; out.data = ans.data; }
      return json_(out);
    }
    if (body.action !== 'save') return json_({ ok: false, error: 'bad_request' });

    const v = validateSubmission_(body, member, now);
    if (!v.ok) return json_(v);

    const when = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    saveAnswer_(token, member, v.record, when, isClosed_(now));
    rebuildOverview_();
    return json_({ ok: true, submittedAt: when });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ---------- 検査（スプレッドシートに触れない関数。node でも試せる） ---------- */
function isClosed_(now) {
  return now.getTime() > new Date(CONFIG.DEADLINE).getTime();
}

function validateSubmission_(body, member, now) {
  if (!body || typeof body !== 'object') return fail_('bad_request');

  const marks = body.marks;
  if (!marks || typeof marks !== 'object' || Array.isArray(marks)) return fail_('bad_request');
  const clean = {};
  const keys = Object.keys(marks);
  if (keys.length > CONFIG.WEEKS * CONFIG.DOWS.length * CONFIG.KOMAS.length) return fail_('bad_request');
  for (let i = 0; i < keys.length; i++) {
    const m = /^(\d+)_(\d+)_(\d+)$/.exec(keys[i]);
    if (!m) return fail_('bad_request');
    const w = Number(m[1]), d = Number(m[2]), k = Number(m[3]);
    if (w >= CONFIG.WEEKS || CONFIG.DOWS.indexOf(d) < 0 || CONFIG.KOMAS.indexOf(k) < 0) return fail_('bad_request');
    if (MARK_VALUES.indexOf(marks[keys[i]]) < 0) return fail_('bad_request');
    clean[keys[i]] = marks[keys[i]];
  }

  const isTeacher = member.grade === TEACHER;
  let subjects = body.subjects;
  if (isTeacher) {
    subjects = [];   // 講師は科目を持たない
  } else {
    if (!Array.isArray(subjects) || subjects.length < 1 || subjects.length > CONFIG.SUBJECTS.length) return fail_('need_subjects');
    const seen = {};
    for (let i = 0; i < subjects.length; i++) {
      if (CONFIG.SUBJECTS.indexOf(subjects[i]) < 0 || seen[subjects[i]]) return fail_('bad_request');
      seen[subjects[i]] = true;
    }
  }

  const memo = typeof body.memo === 'string' ? body.memo.trim() : '';
  if (memo.length > CONFIG.MEMO_MAX) return fail_('memo_too_long');

  const parentOk = body.parentOk === true;
  if (!isTeacher && /^(小|中)/.test(member.grade) && !parentOk) return fail_('need_parent');

  return { ok: true, record: { subjects: subjects, marks: clean, memo: memo, parentOk: parentOk } };
}

function fail_(code) { return { ok: false, error: code }; }

/* ---------- シートの読み書き ---------- */
// 全角・半角をそろえ、空白をすべて除く（「山田 太郎」と「山田太郎」を同じにする）
function normalize_(v) {
  return String(v == null ? '' : v).normalize('NFKC').replace(/\s+/g, '');
}

// 番号と氏名の両方が名簿と合うときだけ、その生徒を返す
function findMember_(token, name) {
  const want = normalize_(name);
  if (!want) return null;
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.ROSTER);
  const last = sh ? sh.getLastRow() : 0;
  if (last < 2) return null;
  const rows = sh.getRange(2, 1, last - 1, ROSTER_HEADERS.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (normalize_(rows[i][0]) === token && normalize_(rows[i][2]) === want) {
      return { name: String(rows[i][2]).trim(), grade: String(rows[i][3]).trim() };
    }
  }
  return null;
}

function findAnswerRow_(sh, token) {
  const last = sh.getLastRow();
  if (last < 2) return null;
  const rows = sh.getRange(2, 1, last - 1, ANSWER_HEADERS.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === token) return { rowIndex: i + 2, values: rows[i] };
  }
  return null;
}

function findAnswer_(token) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.ANSWERS);
  if (!sh) return null;
  const hit = findAnswerRow_(sh, token);
  if (!hit) return null;
  const v = hit.values;
  let marks = {};
  try { marks = JSON.parse(String(v[5] || '{}')); } catch (_) {}
  return {
    submittedAt: String(v[3]),
    data: {
      subjects: String(v[4] || '') ? String(v[4]).split('・') : [],
      marks: marks,
      memo: String(v[6] || ''),
      parentOk: String(v[7]) === '済',
    },
  };
}

function saveAnswer_(token, member, rec, when, late) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.ANSWERS);
  const hit = findAnswerRow_(sh, token);
  const count = hit ? Number(hit.values[8] || 0) + 1 : 1;
  const row = [token, member.name, member.grade, when, rec.subjects.join('・'),
    JSON.stringify(rec.marks), rec.memo, rec.parentOk ? '済' : '', count, late ? '締切後' : ''];
  if (hit) sh.getRange(hit.rowIndex, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  appendHistory_([when, token, member.name, member.grade, rec.subjects.join('・'),
    JSON.stringify(rec.marks), rec.memo, rec.parentOk ? '済' : '', late ? '締切後' : '']);
}

// 提出のたびに「履歴」シートの末尾へ1行足す。足せなくても、提出そのものは成功させる
function appendHistory_(row) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CONFIG.HISTORY);
    if (sh) sh.appendRow(row);
  } catch (_) {}
}

/* ---------- 一覧シート（全員の予定を1枚で見る） ---------- */
// 手で実行しても作り直せる。失敗しても提出は成功させる
function 一覧を更新() { rebuildOverview_(true); }

function rebuildOverview_(loud) {
  try {
    const ss = SpreadsheetApp.getActive();
    const ros = ss.getSheetByName(CONFIG.ROSTER);
    const ansSh = ss.getSheetByName(CONFIG.ANSWERS);
    if (!ros || ros.getLastRow() < 2) return;
    const roster = ros.getRange(2, 1, ros.getLastRow() - 1, ROSTER_HEADERS.length).getValues();
    const answers = {};
    if (ansSh && ansSh.getLastRow() >= 2) {
      ansSh.getRange(2, 1, ansSh.getLastRow() - 1, ANSWER_HEADERS.length).getValues().forEach(function (r) { answers[String(r[0])] = r; });
    }
    const slots = [];
    for (let w = 0; w < CONFIG.WEEKS; w++) CONFIG.DOWS.forEach(function (d) { CONFIG.KOMAS.forEach(function (k) { slots.push([w, d, k]); }); });
    const FIX = ['氏名', '学年', '状況', '提出日時', '来られる枠', '科目', '補足'];
    const nCols = FIX.length + slots.length;
    const startMs = new Date(CONFIG.START + 'T00:00:00+09:00').getTime();
    const WD = '日月火水木金土';
    const head1 = FIX.map(function () { return ''; }), head2 = FIX.slice();
    slots.forEach(function (s) {
      const ms = startMs + (7 * s[0] + s[1]) * 86400000;
      head1.push(s[2] === CONFIG.KOMAS[0] ? Utilities.formatDate(new Date(ms), 'Asia/Tokyo', 'M/d') + '(' + WD[s[1]] + ')' : '');
      head2.push(String(s[2]));
    });
    const values = [], colors = [];
    const GREEN = '#e6f4e6', RED = '#f4b6b6', YEL = '#fde7a6', BLUE = '#cfe3f7', GRAY = '#eeeeee';
    const MARK = { x: ['×', RED], t: ['△', YEL], q: ['？', BLUE] };
    roster.forEach(function (rr) {
      const token = String(rr[0]), a = answers[token];
      const row = [String(rr[2]), String(rr[3])], bg = [];
      for (let i = 0; i < FIX.length; i++) bg.push(a ? '#ffffff' : GRAY);
      if (!a) {
        row.push('未提出', '', '', '', '');
        slots.forEach(function () { row.push(''); bg.push(GRAY); });
      } else {
        let marks = {};
        try { marks = JSON.parse(String(a[5] || '{}')); } catch (_) {}
        let nx = 0;
        const cells = slots.map(function (s) {
          const m = MARK[marks[s[0] + '_' + s[1] + '_' + s[2]]];
          if (m && m[0] === '×') nx++;
          return m || ['', GREEN];
        });
        row.push(String(a[9] || '') ? '提出済（締切後）' : '提出済', String(a[3]), (slots.length - nx) + '/' + slots.length, String(a[4]), String(a[6]));
        cells.forEach(function (c) { row.push(c[0]); bg.push(c[1]); });
      }
      values.push(row); colors.push(bg);
    });
    // 講師に見せる場合は、別のスプレッドシート（番号の入った名簿を含まないもの）に書く。
    // そのIDはコードに置かず、スクリプトのプロパティ「SHARE_ID」に入れる（リポジトリは公開のため）
    const shareId = PropertiesService.getScriptProperties().getProperty('SHARE_ID');
    const target = shareId ? SpreadsheetApp.openById(shareId) : ss;
    let sh = target.getSheetByName(CONFIG.OVERVIEW) || target.insertSheet(CONFIG.OVERVIEW);
    sh.clear();
    if (sh.getMaxColumns() < nCols) sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());
    sh.getRange(1, 1, 2, nCols).setValues([head1, head2]).setBackground('#dddddd').setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(3, 1, values.length, nCols).setValues(values).setBackgrounds(colors);
    sh.getRange(3, FIX.length + 1, values.length, slots.length).setHorizontalAlignment('center');
    sh.setColumnWidths(FIX.length + 1, slots.length, 26);
    sh.setColumnWidth(1, 90); sh.setColumnWidth(3, 110); sh.setColumnWidth(4, 130); sh.setColumnWidth(7, 160);
    sh.setFrozenRows(2); sh.setFrozenColumns(3);
    // 曜日の区切りに縦線を入れる
    for (let i = 0; i < slots.length; i += CONFIG.KOMAS.length) {
      sh.getRange(1, FIX.length + 1 + i, values.length + 2, 1).setBorder(null, true, null, null, null, null, '#888888', SpreadsheetApp.BorderStyle.SOLID);
    }
  } catch (err) {
    if (loud) throw err;
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
