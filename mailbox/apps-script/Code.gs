/**
 * 느린 우체통 — 구글 시트 연동 스크립트
 *
 * - '멤버' 탭에 적은 이름이 우체통 목록이 됩니다.
 * - 편지는 '편지' 탭에 저장됩니다.
 * - 여는 날(REVEAL_AT) 전에는 몇 통인지만 웹사이트로 전달하고, 내용은 보내지 않습니다.
 * - 학생회 암호(PASSCODE)를 아는 사람만 들어올 수 있습니다.
 *
 * - 편지 내용과 보낸 사람은 시트에 봉인(암호화)된 채로 저장됩니다.
 *   시트를 열어도 'SEAL:…' 같은 알 수 없는 글자로 보이고, 여는 날 이후 웹사이트에서만 풀립니다.
 *
 * ⚠ 아래 두 값은 Apps Script 편집기 안에서만 바꾸세요. GitHub에 올리면 누구나 암호를 볼 수 있어요.
 */
const PASSCODE = '';                              // 학생회 암호. 예: 'inmun2026'
const REVEAL_AT = '2026-12-19T18:00:00+09:00';    // 여는 날과 시각 (한국 시간)

const LETTERS = '편지';
const MEMBERS = '멤버';
const ALL = '__all__';
const COLORS = ['butter', 'mint', 'peach', 'sky', 'lilac'];
const MAX_SENDS_PER_10MIN = 60;

/* ---------- 최초 설정 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const lt = ss.getSheetByName(LETTERS) || ss.insertSheet(LETTERS);
  lt.getRange(1, 1, 1, 6).setValues([['ID', '보낸 시각', '받는 사람', '보낸 사람 (봉인)', '내용 (봉인)', '편지지']]).setFontWeight('bold').setBackground('#ffe7e3');
  key_();
  lt.setFrozenRows(1);
  lt.setColumnWidth(5, 420);
  lt.getRange(2, 2, 1000, 1).setNumberFormat('yyyy.MM.dd HH:mm');

  const mb = ss.getSheetByName(MEMBERS) || ss.insertSheet(MEMBERS, 0);
  mb.getRange(1, 1).setValue('이름').setFontWeight('bold').setBackground('#ffe7e3');
  mb.getRange(1, 2).setValue('← 아래로 집행부원 이름을 한 칸에 한 명씩 적어주세요. 순서대로 우체통에 보입니다.').setFontColor('#8d94a6');
  mb.setFrozenRows(1);
  mb.setColumnWidth(1, 140);

  if (!PASSCODE) throw new Error('설정은 끝났어요. 이제 맨 위 PASSCODE 에 학생회 암호를 적고 저장해주세요.');
}

/* ---------- 공통 ---------- */
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function safe_(s, max) {
  s = String(s == null ? '' : s).replace(/\r\n/g, '\n').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function unsafe_(s) { return String(s == null ? '' : s).replace(/^'/, ''); }
function revealed_() { return Date.now() >= new Date(REVEAL_AT).getTime(); }
function checkCode_(code) { return !!PASSCODE && String(code || '') === PASSCODE; }

function members_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(MEMBERS);
  if (!sh || sh.getLastRow() < 2) return [];
  const seen = {};
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(n => n && !seen[n] && (seen[n] = true));
}

/* ---------- 봉인 (시트에서 내용이 보이지 않게) ---------- */
// 봉인 열쇠는 스크립트 속성에 자동으로 만들어 둡니다. 지우면 기존 편지를 풀 수 없으니 건드리지 마세요.
function key_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('SEAL_KEY');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('SEAL_KEY', k); }
  return k;
}

function xor_(bytes, nonce, key) {
  const out = [];
  let block = [];
  for (let i = 0; i < bytes.length; i++) {
    if (i % 32 === 0) block = Utilities.computeHmacSha256Signature(nonce + ':' + (i / 32), key);
    out.push(((bytes[i] ^ block[i % 32]) << 24) >> 24);
  }
  return out;
}

function seal_(text) {
  text = String(text == null ? '' : text);
  if (!text) return '';
  const nonce = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  const enc = xor_(Utilities.newBlob(text).getBytes(), nonce, key_());
  return 'SEAL:' + nonce + ':' + Utilities.base64Encode(enc);
}

function open_(s) {
  s = String(s == null ? '' : s);
  if (s.indexOf('SEAL:') !== 0) return unsafe_(s);          // 봉인 전 기존 편지
  const parts = s.split(':');
  const dec = xor_(Utilities.base64Decode(parts[2]), parts[1], key_());
  return Utilities.newBlob(dec).getDataAsString();
}

// 이 기능을 넣기 전에 들어온 편지를 봉인합니다. 한 번만 ▶ 실행하면 돼요.
function sealExisting() {
  const sh = SpreadsheetApp.getActive().getSheetByName(LETTERS);
  if (!sh || sh.getLastRow() < 2) return '봉인할 편지가 없어요.';
  const rng = sh.getRange(2, 4, sh.getLastRow() - 1, 2);
  const v = rng.getValues();
  let n = 0;
  v.forEach(r => {
    for (let j = 0; j < 2; j++) {
      const cell = String(r[j]);
      if (cell && cell.indexOf('SEAL:') !== 0) { r[j] = seal_(unsafe_(cell)); n++; }
    }
  });
  rng.setValues(v);
  sh.getRange(1, 4, 1, 2).setValues([['보낸 사람 (봉인)', '내용 (봉인)']]);
  return n + '칸을 봉인했어요.';
}

// 여는 날 이후, 기념으로 시트에 원문을 풀어 남기고 싶을 때 ▶ 실행하세요. 여는 날 전에는 동작하지 않아요.
function unsealAll() {
  if (!revealed_()) throw new Error('아직 여는 날 전이라 풀 수 없어요.');
  const sh = SpreadsheetApp.getActive().getSheetByName(LETTERS);
  if (!sh || sh.getLastRow() < 2) return;
  const rng = sh.getRange(2, 4, sh.getLastRow() - 1, 2);
  rng.setValues(rng.getValues().map(r => [safe_(open_(r[0]), 20), safe_(open_(r[1]), 600)]));
  sh.getRange(1, 4, 1, 2).setValues([['보낸 사람', '내용']]);
}

function letters_(withContent) {
  const sh = SpreadsheetApp.getActive().getSheetByName(LETTERS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .filter(r => r[0])
    .map(r => withContent ? {
      id: String(r[0]),
      at: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
      to: String(r[2]),
      from: open_(r[3]) || '익명',
      body: open_(r[4]),
      color: COLORS.indexOf(String(r[5])) >= 0 ? String(r[5]) : 'butter',
    } : { to: String(r[2]) });
}

/* ---------- 우체통 열기 ---------- */
function doGet(e) {
  try {
    if (!checkCode_(e && e.parameter && e.parameter.code)) return json_({ ok: false, error: 'wrong_code' });
    const members = members_();
    const open = revealed_();
    const all = letters_(open);   // 여는 날 전에는 봉인을 풀지 않음
    const counts = {};
    all.forEach(l => { counts[l.to] = (counts[l.to] || 0) + 1; });
    return json_({
      ok: true,
      members: members,
      revealAt: REVEAL_AT,
      revealed: open,
      counts: counts,
      letters: open ? all : [],   // 여는 날 전에는 내용을 보내지 않음
    });
  } catch (x) {
    return json_({ ok: false, error: String(x.message || x) });
  }
}

/* ---------- 편지 넣기 ---------- */
function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: '요청 형식이 올바르지 않아요.' }); }
  if (!checkCode_(b.code)) return json_({ ok: false, error: '암호가 맞지 않아요. 새로고침 후 다시 입력해주세요.' });
  if (b.action !== 'send') return json_({ ok: false, error: '알 수 없는 요청이에요.' });
  if (b.website) return json_({ ok: true });

  const to = String(b.to || '');
  if (to !== ALL && members_().indexOf(to) < 0) return json_({ ok: false, error: '받는 사람을 찾지 못했어요. 새로고침해주세요.' });
  const body = String(b.body == null ? '' : b.body).replace(/\r\n/g, '\n').trim().slice(0, 600);
  if (!body) return json_({ ok: false, error: '편지 내용을 적어주세요.' });
  const from = String(b.from == null ? '' : b.from).trim().slice(0, 20) || '익명';
  const color = COLORS.indexOf(b.color) >= 0 ? b.color : 'butter';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: '잠시 뒤 다시 시도해주세요.' });
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get('sends') || 0);
    if (n >= MAX_SENDS_PER_10MIN) return json_({ ok: false, error: '지금은 편지가 너무 많이 들어왔어요. 10분 뒤에 다시 넣어주세요.' });
    cache.put('sends', String(n + 1), 600);
    const id = Utilities.getUuid().slice(0, 8);
    SpreadsheetApp.getActive().getSheetByName(LETTERS).appendRow([id, new Date(), to, seal_(from), seal_(body), color]);
    return json_({ ok: true, id: id });
  } finally {
    lock.releaseLock();
  }
}
