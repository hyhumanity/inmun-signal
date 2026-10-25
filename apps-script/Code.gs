/**
 * 인문 신호등 — 구글 시트 연동 스크립트
 *
 * 이 코드는 민원 시트에 붙어서 동작합니다.
 *  - 웹사이트에서 접수한 민원을 '민원' 탭에 한 줄씩 저장
 *  - 공개 가능한 민원만 골라 웹사이트에 전달 (비공개 요청 건은 제외)
 *  - 공감 수 저장
 *  - 학생회가 '처리단계' 칸을 바꾸면 날짜를 자동 기록
 *
 * 처음 한 번: 위쪽 함수 선택에서 setup 을 고르고 ▶ 실행 → 권한 허용
 * 그다음: 배포 → 새 배포 → 유형 '웹 앱' / 실행 사용자 '나' / 액세스 '모든 사용자' → 배포
 */

const SHEET_NAME = '민원';
const HEADERS = ['ID', '접수일시', '분류', '제목', '내용', '비공개요청', '처리단계', '담당국', '답변', '공감', '단계기록', '최종수정'];
const COL = { id: 1, createdAt: 2, category: 3, title: 4, body: 5, private: 6, stage: 7, dept: 8, reply: 9, likes: 10, history: 11, updatedAt: 12 };

const STAGES = { '접수': 'received', '담당 배정': 'assigned', '검토 중': 'review', '학교 전달': 'forwarded', '완료': 'done', '처리 불가': 'rejected' };
const CATS = ['시설', '학사·수업', '복지', '학생회 운영', '기타'];
const DEPTS = ['회장단', '사무국', '기획국', '복지국', '홍보국', '대외협력국'];

// 도배 방지: 10분 동안 전체 접수 최대 건수
const MAX_SUBMITS_PER_10MIN = 30;

/* ---------- 최초 설정 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#e5e9f5');
  sh.setFrozenRows(1);

  const stageRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(STAGES), true).setAllowInvalid(false).build();
  const deptRule = SpreadsheetApp.newDataValidation().requireValueInList(DEPTS, true).setAllowInvalid(false).build();
  sh.getRange(2, COL.stage, 1000, 1).setDataValidation(stageRule);
  sh.getRange(2, COL.dept, 1000, 1).setDataValidation(deptRule);

  sh.setColumnWidth(COL.id, 90);
  sh.setColumnWidth(COL.title, 220);
  sh.setColumnWidth(COL.body, 320);
  sh.setColumnWidth(COL.reply, 320);
  sh.getRange(2, COL.createdAt, 1000, 1).setNumberFormat('yyyy.MM.dd HH:mm');
  sh.getRange(2, COL.updatedAt, 1000, 1).setNumberFormat('yyyy.MM.dd HH:mm');
  sh.getRange(1, COL.history).setNote('자동 기록되는 칸이에요. 직접 고치지 마세요.');
  sh.getRange(1, COL.likes).setNote('웹사이트 공감 수 (자동)');
  sh.getRange(1, COL.private).setNote('Y 이면 웹사이트 목록에 보이지 않습니다.');
}

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('시트 설정이 안 되어 있어요. setup 을 먼저 실행해주세요.');
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 수식으로 해석되지 않도록 (=, +, -, @ 로 시작하는 입력 방지)
function safe_(s, max) {
  s = String(s == null ? '' : s).replace(/\r\n/g, '\n').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function iso_(v) {
  if (v instanceof Date) return v.toISOString();
  const d = new Date(v);
  return isNaN(d) ? '' : d.toISOString();
}

/* ---------- 목록 조회 (웹사이트가 1분마다 호출) ---------- */
function doGet() {
  try {
    const values = sheet_().getDataRange().getValues();
    const items = [];
    for (let i = 1; i < values.length; i++) {
      const r = values[i];
      if (!r[COL.id - 1]) continue;
      if (String(r[COL.private - 1]).trim().toUpperCase() === 'Y') continue;
      let history = [];
      try { history = JSON.parse(r[COL.history - 1] || '[]'); } catch (e) { history = []; }
      items.push({
        id: String(r[COL.id - 1]),
        createdAt: iso_(r[COL.createdAt - 1]),
        category: String(r[COL.category - 1]),
        title: String(r[COL.title - 1]).replace(/^'/, ''),
        body: String(r[COL.body - 1]).replace(/^'/, ''),
        stage: STAGES[String(r[COL.stage - 1]).trim()] || 'received',
        dept: String(r[COL.dept - 1] || ''),
        reply: String(r[COL.reply - 1] || ''),
        likes: Number(r[COL.likes - 1]) || 0,
        history: history,
        updatedAt: iso_(r[COL.updatedAt - 1]),
      });
    }
    return json_({ ok: true, items: items });
  } catch (e) {
    return json_({ ok: false, error: String(e.message || e) });
  }
}

/* ---------- 접수 / 공감 ---------- */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: '요청 형식이 올바르지 않아요.' }); }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: '접속이 몰려 있어요. 잠시 뒤 다시 시도해주세요.' });
  try {
    if (body.action === 'submit') return json_(submit_(body));
    if (body.action === 'like') return json_(like_(body));
    return json_({ ok: false, error: '알 수 없는 요청이에요.' });
  } catch (x) {
    return json_({ ok: false, error: String(x.message || x) });
  } finally {
    lock.releaseLock();
  }
}

function submit_(b) {
  if (b.website) return { ok: true }; // 자동 등록 봇 차단 (사람에게는 보이지 않는 칸)
  const title = safe_(b.title, 60);
  const text = safe_(b.body, 1500);
  const category = CATS.indexOf(b.category) >= 0 ? b.category : '기타';
  if (!title || !text) return { ok: false, error: '제목과 내용을 모두 적어주세요.' };

  const cache = CacheService.getScriptCache();
  const count = Number(cache.get('submits') || 0);
  if (count >= MAX_SUBMITS_PER_10MIN) return { ok: false, error: '지금은 접수가 너무 많아요. 10분 뒤에 다시 시도해주세요.' };
  cache.put('submits', String(count + 1), 600);

  const now = new Date();
  const id = Utilities.getUuid().slice(0, 8);
  sheet_().appendRow([
    id, now, category, title, text, b.private ? 'Y' : '', '접수', '', '', 0,
    JSON.stringify([{ stage: 'received', at: now.toISOString() }]), now,
  ]);
  return { ok: true, id: id };
}

function like_(b) {
  const id = String(b.id || '');
  const delta = b.delta === -1 ? -1 : 1;
  const sh = sheet_();
  const ids = sh.getRange(2, COL.id, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      const cell = sh.getRange(i + 2, COL.likes);
      const n = Math.max(0, (Number(cell.getValue()) || 0) + delta);
      cell.setValue(n);
      return { ok: true, likes: n };
    }
  }
  return { ok: false, error: '해당 민원을 찾지 못했어요.' };
}

/* ---------- 학생회가 처리단계를 바꾸면 날짜 자동 기록 ---------- */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_NAME) return;
  const row = e.range.getRow();
  if (row < 2) return;
  const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
  const now = new Date();

  if (c1 <= COL.stage && COL.stage <= c2) {
    const label = String(sh.getRange(row, COL.stage).getValue()).trim();
    const key = STAGES[label];
    if (key) {
      const hCell = sh.getRange(row, COL.history);
      let h = [];
      try { h = JSON.parse(hCell.getValue() || '[]'); } catch (x) { h = []; }
      if (!h.length || h[h.length - 1].stage !== key) {
        h.push({ stage: key, at: now.toISOString() });
        hCell.setValue(JSON.stringify(h));
      }
    }
  }
  const touched = [COL.stage, COL.dept, COL.reply].some(c => c1 <= c && c <= c2);
  if (touched) sh.getRange(row, COL.updatedAt).setValue(now);
}
