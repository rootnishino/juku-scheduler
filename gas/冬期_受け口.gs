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
  ANSWERS: '回答',
};
const ANSWER_HEADERS = ['合言葉', '氏名', '学年', '提出日時', '科目', '予定(JSON)', '補足', '保護者確認', '提出回数', '締切後'];
const ROSTER_HEADERS = ['合言葉', '生徒ID', '氏名', '学年'];
const TOKEN_RE = /^[0-9]{4,10}$/;   // 生徒番号（数字のみ）
const TEACHER = '講師';   // 名簿の「学年」がこの値なら講師（科目と保護者確認は要らない）
const MARK_VALUES = ['x', 't', 'q'];   // x=×  t=△  q=？

/* ---------- 初期設定（最初に一度だけ実行） ---------- */
function 初期設定() {
  const ss = SpreadsheetApp.getActive();
  [[CONFIG.ROSTER, ROSTER_HEADERS], [CONFIG.ANSWERS, ANSWER_HEADERS]].forEach(function (pair) {
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
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
