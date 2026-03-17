/// <reference path="./types.ts" />
/// <reference path="./utils.ts" />
// Core logic to send reminder emails and log results
// Assumes three sheets: 管理表:原本, 設定, 送信ログ

// Referenced types from types.ts
// type TimingKey = ...
// type RecordRow = ...
// type Settings = ...

const SHEET_NAMES = {
  settings: '設定',
  log: '送信ログ',
  guide: '使い方とワークフロー',
} as const;

const TIMING_COLUMNS: TimingKey[] = [
  '1M-1W', '1M+1W', '1M',
  '3M-2W', '3M+2W', '3M',
  '6M-2W', '6M+2W', '6M',
  '12M-1M', '12M+1M', '12M',
  '18M-1M', '18M+1M', '18M',
  '24M-1M', '24M+1M', '24M',
];

const EVAL_COLUMNS: { evalKey: string; group: '1M' | '3M' | '6M' | '12M' | '18M' | '24M' }[] = [
  { evalKey: '評価日：1M', group: '1M' },
  { evalKey: '評価日：3M', group: '3M' },
  { evalKey: '評価日：6M', group: '6M' },
  { evalKey: '評価日：12M', group: '12M' },
  { evalKey: '評価日：18M', group: '18M' },
  { evalKey: '評価日：24M', group: '24M' },
];

const ALERT_COLOR = '#d5a6bd'; // soft red
const CLEAR_COLOR = '#ffffff'; // white
const GREY_COLOR = '#cccccc'; // sent (success) grey

// Map a group (e.g., '1M') to its three timing columns
function groupTimingsFor(group: '1M' | '3M' | '6M' | '12M' | '18M' | '24M'): string[] {
  const map: Record<'1M' | '3M' | '6M' | '12M' | '18M' | '24M', string[]> = {
    '1M': ['1M-1W', '1M', '1M+1W'],
    '3M': ['3M-2W', '3M', '3M+2W'],
    '6M': ['6M-2W', '6M', '6M+2W'],
    '12M': ['12M-1M', '12M', '12M+1M'],
    '18M': ['18M-1M', '18M', '18M+1M'],
    '24M': ['24M-1M', '24M', '24M+1M'],
  } as const;
  return map[group];
}

function groupLabelFrom(group: '1M' | '3M' | '6M' | '12M' | '18M' | '24M'): string {
  const map: Record<typeof group, string> = {
    '1M': '登録1ヶ月後評価',
    '3M': '登録3ヶ月後評価',
    '6M': '登録6ヶ月後評価',
    '12M': '登録12ヶ月後評価',
    '18M': '登録18ヶ月後評価',
    '24M': '登録24ヶ月後評価',
  } as const;
  return map[group] || group + '評価';
}

function getSheetByName(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function getSettings(): Settings {
  const sh = getSheetByName(SHEET_NAMES.settings);
  const values = sh.getDataRange().getValues();
  const map: Record<string, string> = {};
  if (!values.length) return {} as Settings;
  const headerRow = values[0].map((v) => String(v || '').trim());
  const headerHasMany = headerRow.filter(Boolean).length >= 3; // heuristic for header layout
  // Detect institution mapping table headers
  const looksLikeTable = headerRow.includes('登録機関') && headerRow.includes('登録機関コード') && headerRow.includes('送信先アドレス');
  let recipientsByCode: Record<string, string> | undefined;
  let namesByCode: Record<string, string> | undefined;
  if (headerHasMany && values.length >= 2) {
    // Header-layout: row1 headers, row2 values
    const data = values[1] || [];
    headerRow.forEach((h, i) => {
      if (!h) return;
      map[h] = String((data[i] ?? '')).trim();
    });
    // Additionally, if the sheet looks like a table, parse all rows into mapping
    if (looksLikeTable && values.length >= 2) {
      recipientsByCode = {};
      namesByCode = {};
      const idxName = headerRow.indexOf('登録機関');
      const idxCode = headerRow.indexOf('登録機関コード');
      const idxTo = headerRow.indexOf('送信先アドレス');
      for (let r = 1; r < values.length; r++) {
        const row = values[r];
        const name = String(row[idxName] || '').trim();
        const code = String(row[idxCode] || '').trim().toUpperCase();
        const to = String(row[idxTo] || '').trim();
        if (!code) continue;
        if (name) namesByCode[code] = name;
        if (to) recipientsByCode[code] = to;
      }
    }
  } else {
    // Key-value layout: colA key, colB value
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || '').trim();
      const val = String(values[i][1] || '').trim();
      if (key) map[key] = val;
    }
  }
  // Synonyms / normalization
  // Support backward/forward compatible key aliases to allow header renames in 設定 sheet.
  const aliases: Array<[from: string, to: string]> = [
    ['登録機関', '登録機関名'],
    // Request: allow renaming of subject/body keys
    ['カスタムメール件名', 'メール件名'],
    ['カスタムメール本文（HTML）', 'メール本文（HTML）'],
  ];
  aliases.forEach(([from, to]) => {
    if (map[from] && !map[to]) map[to] = map[from];
  });
  const settings = map as Settings;
  if (recipientsByCode) settings.recipientsByCode = recipientsByCode;
  if (namesByCode) settings.namesByCode = namesByCode;
  return settings;
}

function getHeaderMap(sh: GoogleAppsScript.Spreadsheet.Sheet): Map<string, number> {
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = new Map<string, number>();
  header.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map.set(key, i);
  });
  return map;
}

function readRowsFromSheet(sh: GoogleAppsScript.Spreadsheet.Sheet): { rows: RecordRow[]; headerMap: Map<string, number> } {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return { rows: [], headerMap: new Map() };
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const headerMap = getHeaderMap(sh);
  const rows: RecordRow[] = data.map((row) => {
    const obj: any = {};
    header.forEach((h, idx) => {
      const key = String(h || '').trim();
      obj[key] = row[idx];
    });
    return obj as RecordRow;
  });
  return { rows, headerMap };
}

function appendLog(被験者ID: string, 登録医療機関: string, 送信タイミング: string, 送信先: string, 成功: boolean, message?: string) {
  const sh = getSheetByName(SHEET_NAMES.log);
  ensureLogHeader(sh);
  sh.appendRow([
    被験者ID,
    登録医療機関,
    送信タイミング,
    送信先,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss'),
    成功 ? `成功` : `失敗${message ? '：' + message : ''}`,
  ]);
}

function ensureLogHeader(sh: GoogleAppsScript.Spreadsheet.Sheet) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(['被験者ID', '登録医療機関', '送信タイミング', '送信先', '送信日時', '送信結果']);
  } else {
    const header = sh.getRange(1, 1, 1, Math.max(6, sh.getLastColumn())).getValues()[0];
    if (
      String(header[0]) !== '被験者ID' ||
      String(header[1]) !== '登録医療機関' ||
      String(header[2]) !== '送信タイミング' ||
      String(header[3]) !== '送信先' ||
      String(header[4]) !== '送信日時' ||
      String(header[5]) !== '送信結果'
    ) {
      // Overwrite first row to the expected header to be safe
      sh.getRange(1, 1, 1, 6).setValues([['被験者ID', '登録医療機関', '送信タイミング', '送信先', '送信日時', '送信結果']]);
    }
  }
}

function alreadyLoggedToday(被験者ID: string, 送信タイミング候補: string | string[]): boolean {
  const sh = getSheetByName(SHEET_NAMES.log);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const range = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const candidates = Array.isArray(送信タイミング候補) ? 送信タイミング候補 : [送信タイミング候補];
  // columns: 0:被験者ID, 1:登録医療機関, 2:送信タイミング, 3:送信先, 4:送信日時, 5:送信結果
  return range.some((r) => {
    if (String(r[0]) !== 被験者ID) return false;
    if (!String(r[4]).startsWith(today)) return false;
    const logged = String(r[2]);
    return candidates.includes(logged);
  });
}

function sendMail(toList: string[], subject: string, htmlBody: string, cc?: string[], bcc?: string[]) {
  if (!toList.length) throw new Error('No recipients');
  MailApp.sendEmail({
    to: toList.join(','),
    cc: cc && cc.length ? cc.join(',') : undefined,
    bcc: bcc && bcc.length ? bcc.join(',') : undefined,
    subject,
    htmlBody,
  });
}

function buildContext(row: RecordRow, timingDisplay: string, settings: Settings, extra?: Partial<Record<string, string>>): Record<string, string> {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const subjectId = String(row.被験者ID || '');
  const prefix = (subjectId || '').split('-')[0].trim().toUpperCase();
  const nameFromMap = prefix && (settings as any).namesByCode ? (settings as any).namesByCode[prefix] : '';
  const instName = String(nameFromMap || settings.登録機関名 || prefix || '');
  const instCode = String(prefix || settings.登録機関コード || '');
  return {
    被験者ID: String(row.被験者ID || ''),
    登録機関名: instName,
    登録機関コード: instCode,
    送信タイミング: timingDisplay,
    登録日: row.登録日 ? Utilities.formatDate(new Date(row.登録日 as any), Session.getScriptTimeZone(), 'yyyy/MM/dd') : '',
    今日: today,
    ...(extra || {}),
  };
}

function processDueRows() {
  const settings = getSettings();
  const guideFallback = getGuideFallbackTemplates();
  const defaultSubject = '【{{登録機関コード}}】{{被験者ID}} {{送信タイミング}} リマインド';
  const defaultBody = [
    '<p>{{登録機関名}} ご担当者様</p>',
    '<p>以下の被験者について、<strong>{{送信タイミング}}</strong> のリマインダーをお送りします。</p>',
    '<p>',
    '被験者ID：<strong>{{被験者ID}}</strong><br>',
    '登録日：{{登録日}}<br>',
    '今回評価：<strong>{{送信タイミング}}</strong>',
    '</p>',
    '<p>・{{送信タイミング}} に該当する評価の実施／日程調整をお願いいたします。<br>・すでに評価が実施済みの場合は、恐れ入りますが本メールはご放念ください。</p>',
    '<p>ご不明点があれば、Kizuna事務局（担当：大森）までご連絡ください。</p>',
    '<hr>',
    '<p style="color:#666;font-size:12px">本メールはシステムより自動送信されています。</p>'
  ].join('');
  const subjectTpl = (settings.メール件名 && String(settings.メール件名).trim()) || guideFallback.subject || defaultSubject;
  const bodyTpl = (((settings as any)["メール本文（HTML）"] && String((settings as any)["メール本文（HTML）"]).trim()) || guideFallback.body || defaultBody);
  const globalTo = splitEmails(settings.送信先アドレス);
  const globalCc = splitEmails((settings as any).CC);
  const globalBcc = splitEmails((settings as any).BCC);

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');

  // Iterate all sheets starting with "管理表"
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().startsWith('管理表'));
  sheets.forEach(sh => {
    const { rows } = readRowsFromSheet(sh);
    rows.forEach((row) => {
      if (!isAutoOn(row)) return;
      if (!row.被験者ID) return;
      const 登録医療機関 = deriveInstitutionFromId(settings, String(row.被験者ID));

      // Determine evaluation completion per group
      const evalDone: Record<string, boolean> = {};
      EVAL_COLUMNS.forEach(({ evalKey, group }) => {
        const v = (row as any)[evalKey];
        evalDone[group] = !!v; // if evaluation date set, skip reminders for this group
      });

      TIMING_COLUMNS.forEach((col) => {
        const group = (col.match(/^(\d+M)/)?.[1] || col) as '1M' | '3M' | '6M' | '12M' | '18M' | '24M'; // 1M-1W -> 1M
        if (evalDone[group]) return; // already evaluated
        const groupLabel = groupLabelFrom(group);

        // Decide the planned date:
        // 1) If the sheet has a date in the timing column, use it.
        // 2) Otherwise, compute it from 登録日 + offset.
        let planned: Date | null = null;
        const raw = (row as any)[col];
        if (raw) {
          planned = toLocalDate(raw as any);
        } else if (row.登録日) {
          planned = computePlannedDateFromBase(row.登録日 as any, col);
        }
        if (!planned || isNaN(planned.getTime())) return;

        const dateStr = Utilities.formatDate(planned, tz, 'yyyy/MM/dd');
        if (dateStr !== today) return;
        // Backward-compat: treat existing logs that used raw timing (e.g., 1M-1W) as duplicates too
        if (alreadyLoggedToday(row.被験者ID, [groupLabel, col])) return;

        // 評価予定日（グループの基準タイミング、例: 1M）の日付を算出
        const basePlanned = computeGroupBasePlannedDate(row, group);
        const evalDateStr = basePlanned ? Utilities.formatDate(basePlanned, tz, 'yyyy/MM/dd') : '';
        const ctx = buildContext(row, groupLabel, settings, { 評価予定日: evalDateStr });
        const subject = renderTemplate(subjectTpl, ctx);
        const html = renderTemplate(bodyTpl, ctx);
        const { toSend, ccSend, bccSend } = resolveRecipientsForRow(settings, row, globalTo, globalCc, globalBcc);
        try {
          sendMail(toSend, subject, html, ccSend, bccSend);
          // ログは「原タイミング」（例：1M-1W）で記録する
          appendLog(row.被験者ID, 登録医療機関, col, [toSend.join(','), ccSend.join(','), bccSend.join(',')].filter(Boolean).join(' | '), true);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          appendLog(row.被験者ID, 登録医療機関, col, [toSend.join(','), ccSend.join(','), bccSend.join(',')].filter(Boolean).join(' | '), false, msg);
        }
      });
    });
  });

  // After processing sends, update overdue alert coloring
  updateOverdueAlerts();
}

function setupDailyTrigger() {
  // Creates a time-driven trigger to run processDueRows daily around 9am
  const projectTriggers = ScriptApp.getProjectTriggers();
  projectTriggers.forEach((t) => {
    if (t.getHandlerFunction() === 'processDueRows') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processDueRows').timeBased().atHour(9).everyDays(1).create();
}

function removeAllProjectTriggersForCurrentUser() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    try { ScriptApp.deleteTrigger(t); } catch (e) { /* ignore triggers not owned by this user */ }
  });
}

function removeTriggersWithToast() {
  removeAllProjectTriggersForCurrentUser();
  toast('このユーザーが作成したトリガーを削除しました');
}

// Exposed functions
function main() {
  processDueRows();
}

function init() {
  setupDailyTrigger();
}

// --- Spreadsheet UI helpers ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('リマインダー機能')
    .addItem('自動送信を起動する', 'initWithToast')
    .addItem('自動送信を停止する', 'removeTriggersWithToast')
    .addSeparator()
    .addItem('アラート状況更新', 'updateOverdueAlertsWithToast')
    .addSeparator()
    .addItem('期限超過の未送信分を送信', 'sendOverdueRemindersWithDialog')
    .addItem('評価日アラート対象へ送信', 'sendEvalAlertRemindersWithDialog')
    .addSeparator()
    .addItem('当日リマインド送信テスト', 'mainWithToast')
    .addItem('スキーマ診断', 'diagnoseSchemaWithSheet')
    .addItem('本日リマインドのドライラン', 'previewRemindersWithSheet')
    .addItem('任意日付でドライラン', 'previewRemindersAsOfWithSheet')
    .addItem('任意日付でアラート判定プレビュー', 'previewOverdueAlertsAsOfWithSheet')
    .addItem('単体テスト実行', 'runUnitTests')
    .addToUi();
  // Ensure guide sheet exists for first-time users
  ensureGuideSheet();
}

function initWithToast() {
  init();
  toast('毎日9:00に実行されるトリガーを作成しました');
}

function mainWithToast() {
  processDueRows();
  toast('本日のリマインド処理を実行しました');
}

function toast(message: string, title = 'Visit Management', seconds = 5) {
  SpreadsheetApp.getActive().toast(message, title, seconds);
}

function deriveInstitutionFromId(settings: Settings, subjectId: string): string {
  const prefix = (subjectId || '').split('-')[0].trim().toUpperCase();
  const code = (settings.登録機関コード || '').trim().toUpperCase();
  const name = (settings.登録機関名 || '').trim();
  if (prefix && code && prefix === code && name) return name;
  // Prefer name by mapping if available
  if (prefix && (settings as any).namesByCode && (settings as any).namesByCode[prefix]) {
    return (settings as any).namesByCode[prefix];
  }
  return name || prefix || '';
}

// Choose recipients per subject by prefix mapping; fallback to global 設定
function getRecipientsForSubject(settings: Settings, subjectId: string): { to: string[]; cc: string[]; bcc: string[] } {
  const prefix = (subjectId || '').split('-')[0].trim().toUpperCase();
  const mappedTo = (settings as any).recipientsByCode && prefix ? (settings as any).recipientsByCode[prefix] : '';
  const mappedCc = (settings as any).ccByCode && prefix ? (settings as any).ccByCode[prefix] : '';
  const mappedBcc = (settings as any).bccByCode && prefix ? (settings as any).bccByCode[prefix] : '';
  const to = splitEmails(mappedTo || (settings.送信先アドレス || ''));
  const cc = splitEmails(mappedCc || (settings as any).CC);
  const bcc = splitEmails(mappedBcc || (settings as any).BCC);
  return { to, cc, bcc };
}

function resolveRecipientsForRow(
  settings: Settings,
  row: RecordRow,
  globalTo: string[],
  globalCc: string[],
  globalBcc: string[]
): { toSend: string[]; ccSend: string[]; bccSend: string[] } {
  const { to, cc, bcc } = getRecipientsForSubject(settings, String(row.被験者ID || ''));
  const customTo = splitEmails(String((row as any).カスタム宛先 || ''));
  const toSend = customTo.length ? customTo : (to.length ? to : globalTo);
  const ccSend = cc.length ? cc : globalCc;
  const bccSend = bcc.length ? bcc : globalBcc;
  return { toSend, ccSend, bccSend };
}

// (removed duplicate simpler implementation)

// --- Diagnostics & Preview ---
function diagnoseSchema(): string[] {
  const issues: string[] = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Sheets
  const settingsSh = ss.getSheetByName(SHEET_NAMES.settings);
  if (!settingsSh) issues.push('設定 シートが見つかりません');
  const logSh = ss.getSheetByName(SHEET_NAMES.log);
  if (!logSh) issues.push('送信ログ シートが見つかりません');

  // 設定 values via flexible parser
  if (settingsSh) {
    const s = getSettings();
    if (!s.登録機関名 || !String(s.登録機関名).trim()) issues.push('設定: 「登録機関名（または 登録機関）」が未設定です');
    if (!s.登録機関コード || !String(s.登録機関コード).trim()) issues.push('設定: 「登録機関コード」が未設定です');
    if (!s.送信先アドレス || !String(s.送信先アドレス).trim()) issues.push('設定: 「送信先アドレス」が未設定です');
  }

  // 送信ログ header
  if (logSh) {
    const header = logSh.getRange(1, 1, 1, Math.max(6, logSh.getLastColumn())).getValues()[0];
    const ok = String(header[0]) === '被験者ID' && String(header[1]) === '登録医療機関' && String(header[2]) === '送信タイミング' && String(header[3]) === '送信先' && String(header[4]) === '送信日時' && String(header[5]) === '送信結果';
    if (!ok) issues.push('送信ログ: ヘッダーが想定と異なります（被験者ID, 登録医療機関, 送信タイミング, 送信先, 送信日時, 送信結果）');
  }

  // 管理表* sheets
  const managed = ss.getSheets().filter(s => s.getName().startsWith('管理表'));
  if (!managed.length) issues.push('「管理表」で始まるシートが見つかりません');
  managed.forEach(sh => {
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const required = ['自動送信On/Off', '被験者ID', '登録日'];
    required.forEach(k => { if (!header.includes(k)) issues.push(`${sh.getName()}: 列「${k}」がありません`); });
  });
  return issues;
}

// --- Alerts: highlight overdue timing cells with no dispatch log ---
function anyLogged(被験者ID: string, 送信タイミング候補: string | string[]): boolean {
  const sh = getSheetByName(SHEET_NAMES.log);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const range = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  const candidates = Array.isArray(送信タイミング候補) ? 送信タイミング候補 : [送信タイミング候補];
  return range.some((r) => String(r[0]) === 被験者ID && candidates.includes(String(r[2])));
}

// Check if there's at least one successful send log for the given timing(s)
function hasSuccessfulLog(被験者ID: string, 送信タイミング候補: string | string[]): boolean {
  const sh = getSheetByName(SHEET_NAMES.log);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const range = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  const candidates = Array.isArray(送信タイミング候補) ? 送信タイミング候補 : [送信タイミング候補];
  return range.some((r) => String(r[0]) === 被験者ID && candidates.includes(String(r[2])) && String(r[5]).startsWith('成功'));
}

function updateOverdueAlertsWithToast() {
  const count = updateOverdueAlerts();
  toast(`期限超過の未送信セルを ${count} 件ハイライトしました`);
}

function updateOverdueAlerts(): number {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().startsWith('管理表'));
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');
  let highlighted = 0;

  sheets.forEach(sh => {
    const { rows, headerMap } = readRowsFromSheet(sh);
    rows.forEach((row, rowIdx) => {
      if (!isAutoOn(row) || !row.被験者ID) return;
      // eval done check per group
      const evalDone: Record<string, boolean> = {};
      EVAL_COLUMNS.forEach(({ evalKey, group }) => { evalDone[group] = !!(row as any)[evalKey]; });

      TIMING_COLUMNS.forEach((col) => {
        const group = (col.match(/^(\d+M)/)?.[1] || col) as '1M' | '3M' | '6M' | '12M' | '18M' | '24M';
        const idx = headerMap.get(col);
        if (idx == null) return;
        const cell = sh.getRange(rowIdx + 2, idx + 1);
        // If evaluation is done for this group, grey out the group's timing cells and skip further checks
        if (evalDone[group]) {
          if (cell.getBackground() !== GREY_COLOR) {
            cell.setBackground(GREY_COLOR);
          }
          return;
        }
        const raw = (row as any)[col];
        let planned: Date | null = null;
        if (raw) {
          planned = new Date(raw as any);
        } else if (row.登録日) {
          planned = computePlannedDateFromBase(row.登録日 as any, col);
        }
        if (!planned) return;
        if (isNaN(planned.getTime())) return;
        const plannedStr = Utilities.formatDate(planned, tz, 'yyyy/MM/dd');
        const isPast = plannedStr < todayStr;
        const groupLabel = groupLabelFrom(group);
        const candidates = [groupLabel, col];
        const success = hasSuccessfulLog(String(row.被験者ID), candidates);
        const hasAnyLog = success || anyLogged(String(row.被験者ID), candidates);
        // cell defined above
        if (success) {
          if (cell.getBackground() !== GREY_COLOR) {
            cell.setBackground(GREY_COLOR);
          }
        } else if (isPast && !hasAnyLog) {
          if (cell.getBackground() !== ALERT_COLOR) {
            cell.setBackground(ALERT_COLOR);
            highlighted++;
          }
        } else {
          const bg = cell.getBackground();
          if (bg === ALERT_COLOR) {
            cell.setBackground(CLEAR_COLOR);
          }
        }
      });

      // After per-timing processing: color evaluation cells red if all 3 reminders were sent successfully and eval is still empty
      EVAL_COLUMNS.forEach(({ evalKey, group }) => {
        const evalIdx = headerMap.get(evalKey);
        if (evalIdx == null) return;
        const evalCell = sh.getRange(rowIdx + 2, evalIdx + 1);
        const evalVal = (row as any)[evalKey];
        if (evalVal) {
          // Clear red if evaluation is entered
          if (evalCell.getBackground() === ALERT_COLOR) evalCell.setBackground(CLEAR_COLOR);
          return;
        }
        const timings = groupTimingsFor(group);
        const allSent = timings.every(t => hasSuccessfulLog(String(row.被験者ID), t));
        if (allSent) {
          if (evalCell.getBackground() !== ALERT_COLOR) {
            evalCell.setBackground(ALERT_COLOR);
            highlighted++;
          }
        } else if (evalCell.getBackground() === ALERT_COLOR) {
          evalCell.setBackground(CLEAR_COLOR);
        }
      });
    });
  });
  return highlighted;
}

function diagnoseSchemaWithSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '診断レポート';
  const sh = getOrCreateSheet(name);
  sh.clear();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  const issues = diagnoseSchema();
  sh.getRange(1, 1, 1, 2).setValues([['診断日時', ts]]);
  if (!issues.length) {
    sh.getRange(3, 1, 1, 1).setValues([['問題は見つかりませんでした']]);
  } else {
    sh.getRange(3, 1, issues.length, 1).setValues(issues.map(x => [x]));
  }
  toast('診断レポートを更新しました');
}

function previewRemindersWithSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '診断レポート';
  const sh = getOrCreateSheet(name);
  const settings = getSettings();
  const items = findDueReminders(settings, new Date());
  let row = 10;
  sh.getRange(5, 1, 1, 1).setValues([['本日送信予定（ドライラン）']]);
  if (!items.length) {
    sh.getRange(row, 1, 1, 1).setValues([['本日送信予定はありません']]);
  } else {
    sh.getRange(row, 1, 1, 8).setValues([['シート', '被験者ID', '登録医療機関', 'グループ', '送信タイミング(表示)', '送信タイミング(元列)', '予定日', '本日同一ログあり?']]);
    row++;
    const tz = Session.getScriptTimeZone();
    const data = items.map(it => [it.sheetName, it.被験者ID, it.登録医療機関, it.group, it.groupLabel, it.rawTiming, Utilities.formatDate(it.plannedDate, tz, 'yyyy/MM/dd'), it.alreadyLogged ? 'あり' : 'なし']);
    sh.getRange(row, 1, data.length, 8).setValues(data);
  }
  toast('ドライラン結果を診断レポートに書き出しました');
}

// Prompt for a date string and return a Date, or null if canceled/invalid
function promptDateOrNull(title: string): Date | null {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(title, 'yyyy/MM/dd 形式で入力してください（例：2025/09/12）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  const s = (res.getResponseText() || '').trim();
  if (!s) return null;
  const d = toLocalDate(s);
  if (isNaN(d.getTime())) {
    toast('日付が不正です。yyyy/MM/dd 形式で入力してください');
    return null;
  }
  return d;
}

function previewRemindersAsOfWithSheet() {
  const when = promptDateOrNull('任意日付でドライラン');
  if (!when) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '診断レポート';
  const sh = getOrCreateSheet(name);
  const settings = getSettings();
  const items = findDueReminders(settings, when);
  const tz = Session.getScriptTimeZone();
  const whenStr = Utilities.formatDate(when, tz, 'yyyy/MM/dd');
  const start = Math.max(10, sh.getLastRow() + 2);
  sh.getRange(start, 1, 1, 1).setValues([[`任意日付ドライラン（${whenStr} 時点）`]]);
  if (!items.length) {
    sh.getRange(start + 1, 1, 1, 1).setValues([['送信予定はありません']]);
  } else {
    sh.getRange(start + 1, 1, 1, 8).setValues([['シート', '被験者ID', '登録医療機関', 'グループ', '送信タイミング(表示)', '送信タイミング(元列)', '予定日', '同一日ログあり?']]);
    const data = items.map(it => [it.sheetName, it.被験者ID, it.登録医療機関, it.group, it.groupLabel, it.rawTiming, Utilities.formatDate(it.plannedDate, tz, 'yyyy/MM/dd'), it.alreadyLogged ? 'あり' : 'なし']);
    sh.getRange(start + 2, 1, data.length, 8).setValues(data);
  }
  toast('任意日付のドライラン結果を診断レポートに書き出しました');
}

function findOverdueCandidates(settings: Settings, when: Date): Array<{ sheetName: string; row: number; col: string; 被験者ID: string; 登録医療機関: string; group: '1M' | '3M' | '6M' | '12M' | '18M' | '24M'; groupLabel: string; plannedDate: Date; alreadyLogged: boolean; }> {
  const results: Array<{ sheetName: string; row: number; col: string; 被験者ID: string; 登録医療機関: string; group: '1M' | '3M' | '6M' | '12M' | '18M' | '24M'; groupLabel: string; plannedDate: Date; alreadyLogged: boolean; }> = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().startsWith('管理表'));
  const tz = Session.getScriptTimeZone();
  const whenStr = Utilities.formatDate(when, tz, 'yyyy/MM/dd');
  sheets.forEach(sh => {
    const { rows, headerMap } = readRowsFromSheet(sh);
    rows.forEach((row, rowIdx) => {
      if (!isAutoOn(row) || !row.被験者ID) return;
      const 登録医療機関 = deriveInstitutionFromId(settings, String(row.被験者ID));
      const evalDone: Record<string, boolean> = {};
      EVAL_COLUMNS.forEach(({ evalKey, group }) => { evalDone[group] = !!(row as any)[evalKey]; });
      TIMING_COLUMNS.forEach(col => {
        const group = (col.match(/^(\d+M)/)?.[1] || col) as '1M' | '3M' | '6M' | '12M' | '18M' | '24M';
        if (evalDone[group]) return;
        const idx = headerMap.get(col);
        if (idx == null) return;
        const raw = (row as any)[col];
        let planned: Date | null = null;
        if (raw) planned = toLocalDate(raw as any);
        else if (row.登録日) planned = computePlannedDateFromBase(row.登録日 as any, col);
        if (!planned || isNaN(planned.getTime())) return;
        const plannedStr = Utilities.formatDate(planned, tz, 'yyyy/MM/dd');
        const isPast = plannedStr < whenStr; // strictly past
        const groupLabel = groupLabelFrom(group);
        const logged = anyLogged(String(row.被験者ID), [groupLabel, col]);
        if (isPast && !logged) {
          results.push({ sheetName: sh.getName(), row: rowIdx + 2, col, 被験者ID: String(row.被験者ID), 登録医療機関, group, groupLabel, plannedDate: planned, alreadyLogged: logged });
        }
      });
    });
  });
  return results;
}

function previewOverdueAlertsAsOfWithSheet() {
  const when = promptDateOrNull('任意日付でアラート判定プレビュー');
  if (!when) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '診断レポート';
  const sh = getOrCreateSheet(name);
  const settings = getSettings();
  const items = findOverdueCandidates(settings, when);
  const tz = Session.getScriptTimeZone();
  const whenStr = Utilities.formatDate(when, tz, 'yyyy/MM/dd');
  const start = Math.max(10, sh.getLastRow() + 2);
  sh.getRange(start, 1, 1, 1).setValues([[`任意日付アラート判定プレビュー（${whenStr} 時点）`]]);
  if (!items.length) {
    sh.getRange(start + 1, 1, 1, 1).setValues([['ハイライト対象はありません']]);
  } else {
    sh.getRange(start + 1, 1, 1, 9).setValues([['シート', '行', '列(元タイミング)', '被験者ID', '登録医療機関', 'グループ', '送信タイミング(表示)', '予定日', 'ログあり?']]);
    const data = items.map(it => [it.sheetName, it.row, it.col, it.被験者ID, it.登録医療機関, it.group, it.groupLabel, Utilities.formatDate(it.plannedDate, tz, 'yyyy/MM/dd'), it.alreadyLogged ? 'あり' : 'なし']);
    sh.getRange(start + 2, 1, data.length, 9).setValues(data);
  }
  toast('任意日付のアラート判定プレビューを診断レポートに書き出しました');
}

// --- Overdue sending ---
function sendOverdueRemindersWithDialog() {
  const settings = getSettings();
  const items = findOverdueCandidates(settings, new Date());
  if (!items.length) {
    toast('期限超過の未送信はありません');
    return;
  }
  items.sort((a, b) => a.plannedDate.getTime() - b.plannedDate.getTime());
  const tz = Session.getScriptTimeZone();
  const first = Utilities.formatDate(items[0].plannedDate, tz, 'yyyy/MM/dd');
  const last = Utilities.formatDate(items[items.length - 1].plannedDate, tz, 'yyyy/MM/dd');
  const ui = SpreadsheetApp.getUi();
  // Build detailed list like:
  // 対象: N 件
  // ・機関名（yyyy/MM/dd）
  const lines = items.map(it => `・${it.登録医療機関}（${Utilities.formatDate(it.plannedDate, tz, 'yyyy/MM/dd')}）`);
  const body = `対象: ${items.length} 件\n${lines.join('\n')}\nを送信します`;
  const res = ui.alert(
    '期限超過の未送信を送信',
    body,
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) {
    toast('送信をキャンセルしました');
    return;
  }
  const result = sendOverdueRemindersNow(new Date());
  toast(`期限超過の未送信を ${result.success}/${result.total} 件送信しました`);
}

function sendOverdueRemindersNow(when: Date): { total: number; success: number; failure: number } {
  const settings = getSettings();
  const guideFallback = getGuideFallbackTemplates();
  const defaultSubject = '【{{登録機関コード}}】{{被験者ID}} {{送信タイミング}} リマインド';
  const defaultBody = [
    '<p>{{登録機関名}} ご担当者様</p>',
    '<p>以下の被験者について、<strong>{{送信タイミング}}</strong> のリマインダーをお送りします。</p>',
    '<p>',
    '被験者ID：<strong>{{被験者ID}}</strong><br>',
    '登録日：{{登録日}}<br>',
    '今回評価：<strong>{{送信タイミング}}</strong>',
    '</p>',
    '<p>・{{送信タイミング}} に該当する評価の実施／日程調整をお願いいたします。<br>・すでに評価が実施済みの場合は、恐れ入りますが本メールはご放念ください。</p>',
    '<p>ご不明点があれば、Kizuna事務局（担当：大森）までご連絡ください。</p>',
    '<hr>',
    '<p style="color:#666;font-size:12px">本メールはシステムより自動送信されています。</p>'
  ].join('');
  const subjectTpl = (settings.メール件名 && String(settings.メール件名).trim()) || guideFallback.subject || defaultSubject;
  const bodyTpl = (((settings as any)['メール本文（HTML）'] && String((settings as any)['メール本文（HTML）']).trim()) || guideFallback.body || defaultBody);
  const globalTo = splitEmails(settings.送信先アドレス);
  const globalCc = splitEmails((settings as any).CC);
  const globalBcc = splitEmails((settings as any).BCC);

  const items = findOverdueCandidates(settings, when);
  if (!items.length) {
    return { total: 0, success: 0, failure: 0 };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const perSheet: Record<string, { rows: RecordRow[]; sheet: GoogleAppsScript.Spreadsheet.Sheet }> = {};
  const getSheetRows = (name: string) => {
    if (!perSheet[name]) {
      const sh = ss.getSheetByName(name);
      if (!sh) throw new Error(`Sheet not found: ${name}`);
      const { rows } = readRowsFromSheet(sh);
      perSheet[name] = { rows, sheet: sh };
    }
    return perSheet[name].rows;
  };

  let success = 0;
  let failure = 0;
  const tz = Session.getScriptTimeZone();

  items.forEach((it) => {
    try {
      const rows = getSheetRows(it.sheetName);
      const rowObj = rows[it.row - 2]; // data rows start at row 2
      if (!rowObj || !isAutoOn(rowObj)) return; // skip if toggled off now
      // Skip if evaluation done now
      const evalKey = EVAL_COLUMNS.find(e => e.group === it.group)?.evalKey;
      if (evalKey && (rowObj as any)[evalKey]) return;
      // Skip if logged (race safety)
      if (anyLogged(String(rowObj.被験者ID), [it.groupLabel, it.col])) return;

      // Build context and send
      const basePlanned = computeGroupBasePlannedDate(rowObj, it.group);
      const evalDateStr = basePlanned ? Utilities.formatDate(basePlanned, tz, 'yyyy/MM/dd') : '';
      const ctx = buildContext(rowObj, it.groupLabel, settings, { 評価予定日: evalDateStr });
      const subject = renderTemplate(subjectTpl, ctx);
      const html = renderTemplate(bodyTpl, ctx);
      const { toSend, ccSend, bccSend } = resolveRecipientsForRow(settings, rowObj, globalTo, globalCc, globalBcc);
      sendMail(toSend, subject, html, ccSend, bccSend);
      appendLog(String(rowObj.被験者ID), it.登録医療機関, it.col, [toSend.join(','), ccSend.join(','), bccSend.join(',')].filter(Boolean).join(' | '), true);
      success++;
    } catch (e) {
      try {
        const rows = getSheetRows(it.sheetName);
        const rowObj = rows[it.row - 2];
        const fallbackRow = ({ 被験者ID: String(it.被験者ID) } as RecordRow);
        const { toSend, ccSend, bccSend } = resolveRecipientsForRow(settings, rowObj || fallbackRow, globalTo, globalCc, globalBcc);
        appendLog(String(rowObj?.被験者ID || it.被験者ID), it.登録医療機関, it.col, [toSend.join(','), ccSend.join(','), bccSend.join(',')].filter(Boolean).join(' | '), false, e instanceof Error ? e.message : String(e));
      } catch (_) {
        // ignore logging error
      }
      failure++;
    }
  });

  // refresh alerts after sending
  updateOverdueAlerts();
  return { total: items.length, success, failure };
}

// --- Evaluation alert sending (all 3 reminders sent, eval date empty) ---
function findEvalAlertCandidates(settings: Settings): Array<{ sheetName: string; row: number; 被験者ID: string; 登録医療機関: string; group: '1M' | '3M' | '6M' | '12M' | '18M' | '24M'; groupLabel: string; evalKey: string; }> {
  const results: Array<{ sheetName: string; row: number; 被験者ID: string; 登録医療機関: string; group: '1M' | '3M' | '6M' | '12M' | '18M' | '24M'; groupLabel: string; evalKey: string; }> = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().startsWith('管理表'));
  sheets.forEach(sh => {
    const { rows } = readRowsFromSheet(sh);
    rows.forEach((row, rowIdx) => {
      if (!isAutoOn(row) || !row.被験者ID) return;
      const 登録医療機関 = deriveInstitutionFromId(settings, String(row.被験者ID));
      EVAL_COLUMNS.forEach(({ evalKey, group }) => {
        const evalVal = (row as any)[evalKey];
        if (evalVal) return; // already filled
        const timings = groupTimingsFor(group);
        const allSent = timings.every(t => hasSuccessfulLog(String(row.被験者ID), t));
        if (allSent) {
          results.push({ sheetName: sh.getName(), row: rowIdx + 2, 被験者ID: String(row.被験者ID), 登録医療機関, group, groupLabel: groupLabelFrom(group), evalKey });
        }
      });
    });
  });
  return results;
}

function sendEvalAlertRemindersWithDialog() {
  const settings = getSettings();
  const items = findEvalAlertCandidates(settings);
  if (!items.length) {
    toast('評価日アラート対象はありません');
    return;
  }
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('評価日アラート対象に送信', `対象: ${items.length} 件にリマインドを送信しますか？`, ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  const result = sendEvalAlertRemindersNow();
  toast(`評価日アラート対象へ ${result.success}/${result.total} 件送信しました`);
}

function sendEvalAlertRemindersNow(): { total: number; success: number; failure: number } {
  const settings = getSettings();
  const guideFallback = getGuideFallbackTemplates();
  const defaultSubject = '【{{登録機関コード}}】{{被験者ID}} {{送信タイミング}} リマインド（評価日未入力）';
  const defaultBody = [
    '<p>{{登録機関名}} ご担当者様</p>',
    '<p>以下の被験者について、<strong>{{送信タイミング}}</strong> のリマインダーは全て送信済みですが、評価日の入力が確認できません。</p>',
    '<p>お手数ですが、評価日のご入力または状況のご確認をお願いいたします。</p>',
    '<p>被験者ID：<strong>{{被験者ID}}</strong><br>登録日：{{登録日}}<br>今回評価：<strong>{{送信タイミング}}</strong><br>評価予定日：{{評価予定日}}</p>',
    '<hr><p style="color:#666;font-size:12px">本メールはシステムより自動送信されています。</p>'
  ].join('');
  const subjectTpl = (settings.メール件名 && String(settings.メール件名).trim()) || guideFallback.subject || defaultSubject;
  const bodyTpl = (((settings as any)['メール本文（HTML）'] && String((settings as any)['メール本文（HTML）']).trim()) || guideFallback.body || defaultBody);
  const globalTo = splitEmails(settings.送信先アドレス);
  const globalCc = splitEmails((settings as any).CC);
  const globalBcc = splitEmails((settings as any).BCC);

  const items = findEvalAlertCandidates(settings);
  if (!items.length) return { total: 0, success: 0, failure: 0 };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const perSheet: Record<string, { rows: RecordRow[]; sheet: GoogleAppsScript.Spreadsheet.Sheet }> = {};
  const getSheetRows = (name: string) => {
    if (!perSheet[name]) {
      const sh = ss.getSheetByName(name);
      if (!sh) throw new Error(`Sheet not found: ${name}`);
      const { rows } = readRowsFromSheet(sh);
      perSheet[name] = { rows, sheet: sh };
    }
    return perSheet[name].rows;
  };

  let success = 0;
  let failure = 0;
  const tz = Session.getScriptTimeZone();

  items.forEach(it => {
    try {
      const rows = getSheetRows(it.sheetName);
      const rowObj = rows[it.row - 2];
      if (!rowObj || !isAutoOn(rowObj)) return;
      // skip if eval has been entered now
      if ((rowObj as any)[it.evalKey]) return;
      // Build context using group base planned date
      const basePlanned = computeGroupBasePlannedDate(rowObj, it.group);
      const evalDateStr = basePlanned ? Utilities.formatDate(basePlanned, tz, 'yyyy/MM/dd') : '';
      const ctx = buildContext(rowObj, groupLabelFrom(it.group), settings, { 評価予定日: evalDateStr });
      const subject = renderTemplate(subjectTpl, ctx);
      const html = renderTemplate(bodyTpl, ctx);
      const { toSend, ccSend, bccSend } = resolveRecipientsForRow(settings, rowObj, globalTo, globalCc, globalBcc);
      sendMail(toSend, subject, html, ccSend, bccSend);
      // Log with synthetic timing label to distinguish
      appendLog(String(rowObj.被験者ID), it.登録医療機関, `${it.group}:評価日アラート`, [toSend.join(','), ccSend.join(','), bccSend.join(',')].filter(Boolean).join(' | '), true);
      success++;
    } catch (e) {
      try {
        const rows = getSheetRows(it.sheetName);
        const rowObj = rows[it.row - 2];
        const fallbackRow = ({ 被験者ID: String(it.被験者ID) } as RecordRow);
        const { toSend, ccSend, bccSend } = resolveRecipientsForRow(settings, rowObj || fallbackRow, globalTo, globalCc, globalBcc);
        appendLog(String(rowObj?.被験者ID || it.被験者ID), it.登録医療機関, `${it.group}:評価日アラート`, [toSend.join(','), ccSend.join(','), bccSend.join(',')].filter(Boolean).join(' | '), false, e instanceof Error ? e.message : String(e));
      } catch (_) { /* ignore */ }
      failure++;
    }
  });

  // refresh alerts to reflect any changes
  updateOverdueAlerts();
  return { total: items.length, success, failure };
}

function findDueReminders(settings: Settings, when: Date): Array<{ sheetName: string; 被験者ID: string; 登録医療機関: string; group: '1M' | '3M' | '6M' | '12M' | '18M' | '24M'; groupLabel: string; rawTiming: string; plannedDate: Date; alreadyLogged: boolean; }> {
  const results: Array<{ sheetName: string; 被験者ID: string; 登録医療機関: string; group: '1M' | '3M' | '6M' | '12M' | '18M' | '24M'; groupLabel: string; rawTiming: string; plannedDate: Date; alreadyLogged: boolean; }> = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().startsWith('管理表'));
  const tz = Session.getScriptTimeZone();
  const dayStr = Utilities.formatDate(when, tz, 'yyyy/MM/dd');

  sheets.forEach(sh => {
    const { rows } = readRowsFromSheet(sh);
    rows.forEach(row => {
      if (!isAutoOn(row) || !row.被験者ID) return;
      const 登録医療機関 = deriveInstitutionFromId(settings, String(row.被験者ID));
      const evalDone: Record<string, boolean> = {};
      EVAL_COLUMNS.forEach(({ evalKey, group }) => { evalDone[group] = !!(row as any)[evalKey]; });
      TIMING_COLUMNS.forEach(col => {
        const group = (col.match(/^(\d+M)/)?.[1] || col) as '1M' | '3M' | '6M' | '12M' | '18M' | '24M';
        if (evalDone[group]) return;
        let planned: Date | null = null;
        const raw = (row as any)[col];
        if (raw) planned = toLocalDate(raw as any);
        else if (row.登録日) planned = computePlannedDateFromBase(row.登録日 as any, col);
        if (!planned || isNaN(planned.getTime())) return;
        const ds = Utilities.formatDate(planned, tz, 'yyyy/MM/dd');
        if (ds !== dayStr) return;
        const groupLabel = groupLabelFrom(group);
        results.push({ sheetName: sh.getName(), 被験者ID: row.被験者ID, 登録医療機関, group, groupLabel, rawTiming: col, plannedDate: planned, alreadyLogged: alreadyLoggedToday(row.被験者ID, [groupLabel, col]) });
      });
    });
  });
  return results;
}

// Support both "自動送信On/Off" and legacy variations; accept boolean or common truthy strings
function isAutoOn(row: RecordRow): boolean {
  const v = (row as any)['自動送信On/Off'] ?? (row as any)['自動送信OnOff'] ?? (row as any)['自動送信'];
  if (typeof v === 'boolean') return v;
  const s = String(v || '').trim().toLowerCase();
  return s === 'true' || s === 'on' || s === '1' || s === 'y' || s === 'yes';
}

function getOrCreateSheet(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// --- Guide sheet and default templates ---
const GUIDE_KEYS = {
  subject: 'デフォルト件名',
  body: 'デフォルト本文（HTML）',
} as const;

function ensureGuideSheet() {
  const sh = getOrCreateSheet(SHEET_NAMES.guide);
  const lastRow = Math.max(sh.getLastRow(), 1);
  // Try to locate keys in column A
  const colA = sh.getRange(1, 1, lastRow, 1).getValues().map(r => String(r[0] || '').trim());
  const needHeader = colA.every(v => !v);
  let rowSubject = colA.findIndex(v => v === GUIDE_KEYS.subject) + 1;
  let rowBody = colA.findIndex(v => v === GUIDE_KEYS.body) + 1;
  if (needHeader) {
    const lines = [
      ['使い方とワークフロー', ''],
      ['（概要）', 'このシートには運用手順とデフォルトの件名/本文を記載します。'],
      ['（手順）', '1) 設定シートを整備 2) リマインダー→初期設定 3) ドライラン確認 4) 本日リマインド送信'],
      ['', ''],
      [GUIDE_KEYS.subject, '【{{登録機関コード}}】{{被験者ID}} {{送信タイミング}} リマインド'],
      [GUIDE_KEYS.body, '<p>{{登録機関名}} ご担当者様</p><p>以下の被験者について、<strong>{{送信タイミング}}</strong> のリマインダーをお送りします。</p><p>被験者ID：<strong>{{被験者ID}}</strong><br>登録日：{{登録日}}<br>今回評価：<strong>{{送信タイミング}}</strong></p><p>・{{送信タイミング}} に該当する評価の実施／日程調整をお願いいたします。<br>・すでに評価が実施済みの場合は、恐れ入りますが本メールはご放念ください。</p><hr><p style="color:#666;font-size:12px">本メールはシステムより自動送信されています。</p>'],
    ];
    sh.getRange(1, 1, lines.length, 2).setValues(lines);
    return;
  }
  if (!rowSubject) {
    rowSubject = lastRow + 1;
    sh.getRange(rowSubject, 1, 1, 2).setValues([[GUIDE_KEYS.subject, '【{{登録機関コード}}】{{被験者ID}} {{送信タイミング}} リマインド']]);
  }
  if (!rowBody) {
    rowBody = (rowSubject || lastRow) + 1;
    sh.getRange(rowBody, 1, 1, 2).setValues([[GUIDE_KEYS.body, '<p>{{登録機関名}} ご担当者様</p><p>以下の被験者について、<strong>{{送信タイミング}}</strong> のリマインダーをお送りします。</p><p>被験者ID：<strong>{{被験者ID}}</strong><br>登録日：{{登録日}}<br>今回評価：<strong>{{送信タイミング}}</strong></p><p>・{{送信タイミング}} に該当する評価の実施／日程調整をお願いいたします。<br>・すでに評価が実施済みの場合は、恐れ入りますが本メールはご放念ください。</p><hr><p style="color:#666;font-size:12px">本メールはシステムより自動送信されています。</p>']]);
  }
}

function openGuideSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureGuideSheet();
  const sh = ss.getSheetByName(SHEET_NAMES.guide)!;
  ss.setActiveSheet(sh);
}

function getGuideFallbackTemplates(): { subject?: string; body?: string } {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.guide);
  if (!sh) return {};
  const lastRow = sh.getLastRow();
  if (!lastRow) return {};
  const rows = sh.getRange(1, 1, lastRow, 2).getValues();
  const map: Record<string, string> = {};
  rows.forEach(([k, v]) => {
    const key = String(k || '').trim();
    if (!key) return;
    map[key] = String(v || '').trim();
  });
  // Prefer new keys; fall back to legacy key names for backward compatibility
  const subject = map[GUIDE_KEYS.subject] || map['フォールバック件名'];
  const body = map[GUIDE_KEYS.body] || map['フォールバック本文（HTML）'];
  return { subject, body };
}

// --- date calculation helpers ---
function computePlannedDateFromBase(base: Date | string, timing: string): Date {
  const d = toLocalDate(base as any);
  // Support patterns like: 12M, 12M-1W, 12M+2W, 12M-1M, 12M+1M
  const m = timing.match(/^(\d+)M(?:(-\d+[MW]|\+\d+[MW]))?$/);
  if (!m) {
    const mOnly = timing.match(/^(\d+)M$/);
    if (mOnly) return addMonths(d, parseInt(mOnly[1], 10));
    return d;
  }
  const months = parseInt(m[1], 10);
  let date = addMonths(d, months);
  const offset = m[2];
  if (offset) {
    const sign = offset.startsWith('-') ? -1 : 1;
    const num = parseInt(offset.slice(1, -1), 10);
    const unit = offset.slice(-1); // 'M' or 'W'
    if (unit === 'W') {
      date = addDays(date, sign * num * 7);
    } else if (unit === 'M') {
      date = addMonths(date, sign * num);
    }
  }
  return date;
}

function addMonths(date: Date, months: number): Date {
  const d = new Date(date);
  const day = d.getDate();
  d.setMonth(d.getMonth() + months);
  // handle month overflow (e.g., Jan 31 + 1 month)
  if (d.getDate() < day) d.setDate(0);
  return d;
}

function addDays(date: Date, days: number): Date {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

// Convert a cell value to a Date in script's timezone without UTC drift
function toLocalDate(val: any): Date {
  if (val instanceof Date) return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  const s = String(val || '').trim();
  // Expect formats like yyyy/MM/dd or yyyy-MM-dd
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) {
    const y = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    const d = parseInt(m[3], 10);
    return new Date(y, mo, d);
  }
  // Fallback
  return new Date(s as any);
}

// Given a row and group (e.g., '1M'), return the base planned date for that group
function computeGroupBasePlannedDate(row: RecordRow, group: '1M' | '3M' | '6M' | '12M' | '18M' | '24M'): Date | null {
  const raw = (row as any)[group];
  if (raw) {
    const d = toLocalDate(raw as any);
    if (!isNaN(d.getTime())) return d;
  }
  if (row.登録日) {
    const d = computePlannedDateFromBase(row.登録日 as any, group);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}
