/**
 * 時代華語2 完課測驗 — 成績接收程式（Google Apps Script）第 2 版
 *
 * 學生交卷後，這個程式會：
 *   1. 在「作答紀錄」加一列（每一次交卷一列，含各大題分數、錯題、作答時間）
 *   2. 在「成績冊－班級名稱」寫分數（一班一張，例如「成績冊－國專部G春」）
 *      第 2 列「測驗日期」會在那一課第一個人交卷時自動填上日期
 */

// 成績要寫進的試算表
const SHEET_ID = "1I2iHmCFYk8I2q--oeE5jNKpaohiScpP4YPJLxzGneoc";

// 成績冊要記哪一次的分數："first"＝第一次、"best"＝最高分、"latest"＝最後一次
const SCORE_POLICY = "first";

const LOG_SHEET = "作答紀錄";
const BOOK_PREFIX = "成績冊－";
const LOG_HEAD = ["交卷時間", "課別", "班級", "學號", "姓名", "第幾次", "得分", "滿分", "百分比", "作答分鐘", "各大題", "錯題"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const log = sheet_(ss, LOG_SHEET, LOG_HEAD);
    const secText = (d.sectionOrder || Object.keys(d.sections || {}))
      .map(k => k + " " + d.sections[k]).join("｜");
    log.appendRow([new Date(), d.lesson, d.klass, String(d.seat || ""), d.name, d.attempt,
                   d.score, d.max, d.pct / 100, d.minutes, secText, d.wrong]);
    log.getRange(log.getLastRow(), 9).setNumberFormat("0.0%");
    updateBook_(ss, d);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, msg: "成績接收程式正常運作中（第 2 版）" });
}

function lessons_() {
  const a = [];
  for (let i = 1; i <= 16; i++) a.push("L" + ("0" + i).slice(-2));
  return a;
}

function updateBook_(ss, d) {
  const ls = lessons_();
  const name = BOOK_PREFIX + (d.klass || "未分班");
  let book = ss.getSheetByName(name);
  if (!book) {
    book = ss.insertSheet(name);
    const head = ["班級", "學號", "姓名"].concat(ls).concat(["平均"]);
    book.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold").setHorizontalAlignment("center");
    book.getRange(2, 3).setValue("測驗日期");
    book.getRange(2, 1, 1, head.length).setBackground("#ededed").setFontWeight("bold").setHorizontalAlignment("center");
    book.setFrozenRows(2);
    book.setFrozenColumns(3);
  }
  // 找這一課在哪一欄
  const lastCol = book.getLastColumn();
  const head = book.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const col = head.indexOf(d.lesson) + 1;
  if (col < 1) return;
  const avgCol = head.indexOf("平均") + 1;

  // 測驗日期：這一課第一個人交卷時填上
  const dateCell = book.getRange(2, col);
  if (dateCell.getDisplayValue() === "") {
    dateCell.setNumberFormat("@").setValue(Utilities.formatDate(new Date(), "Asia/Taipei", "MM/dd"));
  }

  // 找學生：有學號就用學號找，沒有學號就用姓名找
  const seat = String(d.seat || "").trim(), sname = String(d.name || "").trim();
  const last = book.getLastRow();
  let row = -1;
  if (last > 2) {
    const vals = book.getRange(3, 2, last - 2, 2).getDisplayValues();
    for (let i = 0; i < vals.length; i++) {
      const hit = seat ? vals[i][0].trim() === seat : vals[i][1].trim() === sname;
      if (hit) { row = i + 3; break; }
    }
  }
  if (row < 0) {
    row = Math.max(last, 2) + 1;
    book.getRange(row, 1, 1, 3).setNumberFormat("@").setValues([[d.klass, seat, sname]]);
    if (avgCol > 0) {
      const a = book.getRange(row, 4).getA1Notation(), b = book.getRange(row, avgCol - 1).getA1Notation();
      book.getRange(row, avgCol).setFormula('=IFERROR(ROUND(AVERAGE(' + a + ':' + b + '),1),"")');
    }
  }
  const cell = book.getRange(row, col);
  const old = cell.getValue();
  const empty = old === "" || old === "缺席";
  const score = Math.round(d.pct);
  if (empty || SCORE_POLICY === "latest" || (SCORE_POLICY === "best" && score > Number(old))) {
    cell.setValue(score).setBackground(null).setFontColor(null).setHorizontalAlignment("center");
  }
  if (d.attempt > 1) cell.setNote("交卷 " + d.attempt + " 次，最新一次 " + score + " 分");
}

function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold").setBackground("#e6e9f6");
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
