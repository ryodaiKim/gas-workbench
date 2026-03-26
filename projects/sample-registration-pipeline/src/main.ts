type SettingsMap = {
  読み取りPDF格納先: string;
  更新頻度: string;
};

type ParsedRecord = {
  試験名: string;
  施設コード: string;
  施設名: string;
  被験者番号: string;
  性別: string;
  採取日: string;
  ポイント名: string;
  検査項目: string;
};

type DashboardStats = {
  totalRows: number;
  uniqueSubjects: number;
  uniqueFacilities: number;
  byPoint: Record<string, number>;
  byItem: Record<string, number>;
};

const SHEET_NAMES = {
  intake: '受付情報一覧',
  settings: '設定',
  log: 'ログ',
} as const;

const INTAKE_HEADERS: Array<keyof ParsedRecord> = [
  '試験名',
  '施設コード',
  '施設名',
  '被験者番号',
  '性別',
  '採取日',
  'ポイント名',
  '検査項目',
];

const LOG_HEADERS = [
  '処理日時',
  'fileId',
  'fileName',
  'result',
  'message',
  'recordsInserted',
] as const;

const FREQ_ALIASES: Record<string, 'hour' | 'day' | 'week' | 'month'> = {
  hour: 'hour',
  hourly: 'hour',
  毎時: 'hour',
  day: 'day',
  daily: 'day',
  日: 'day',
  毎日: 'day',
  week: 'week',
  weekly: 'week',
  週: 'week',
  毎週: 'week',
  month: 'month',
  monthly: 'month',
  月: 'month',
  毎月: 'month',
};

const ITEM_NORMALIZATION: Array<[RegExp, string]> = [
  [/dna\s*抽出/i, 'ＤＮＡ抽出（Ｎ）'],
  [/リンパ球.*11|リンパ球.*１１/i, 'リンパ球株化１１'],
  [/血清.*分離/i, '血清分離（用手法）'],
  [/血漿.*分離/i, '血漿分離（用手法）'],
];

const ITEM_ANCHOR_PATTERNS = [
  /ＤＮＡ抽出（?Ｎ）?/,
  /リンパ球株化[1１][1１]/,
  /血清分離（?用手法）?/,
  /血漿分離（?用手法）?/,
];

function getSheetOrThrow(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function ensureHeaders(sheet: GoogleAppsScript.Spreadsheet.Sheet, headers: readonly string[]): void {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([...headers]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0].map((v) => String(v || '').trim());
  const mismatch = headers.some((h, i) => existing[i] !== h);
  if (mismatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([Array.from(headers)]);
  }
}

function nowStr(): string {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

function appendLog(fileId: string, fileName: string, result: 'SUCCESS' | 'FAIL' | 'SKIP', message: string, recordsInserted: number): void {
  const sh = getSheetOrThrow(SHEET_NAMES.log);
  ensureHeaders(sh, LOG_HEADERS);
  sh.appendRow([nowStr(), fileId, fileName, result, message, recordsInserted]);
}

function getSettings(): SettingsMap {
  const sh = getSheetOrThrow(SHEET_NAMES.settings);
  const values = sh.getDataRange().getValues();
  const map: Record<string, string> = {};
  values.forEach((row) => {
    const key = String(row[0] || '').trim();
    const val = String(row[1] || '').trim();
    if (key) map[key] = val;
  });
  const source = map['読み取りPDF格納先'];
  const freq = map['更新頻度'];
  if (!source) throw new Error('設定シートに「読み取りPDF格納先」がありません');
  if (!freq) throw new Error('設定シートに「更新頻度」がありません');
  return { 読み取りPDF格納先: source, 更新頻度: freq };
}

function normalizeFrequency(value: string): 'hour' | 'day' | 'week' | 'month' {
  const normalized = String(value || '').trim().toLowerCase();
  const mapped = FREQ_ALIASES[normalized];
  if (!mapped) throw new Error(`更新頻度が不正です: ${value}`);
  return mapped;
}

function extractIdFromDriveUrl(url: string): string {
  const s = String(url || '').trim();
  const byFolderPath = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (byFolderPath && byFolderPath[1]) return byFolderPath[1];
  const byFilePath = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (byFilePath && byFilePath[1]) return byFilePath[1];
  const byIdQuery = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (byIdQuery && byIdQuery[1]) return byIdQuery[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  throw new Error(`Drive URL/ID 解析失敗: ${url}`);
}

function isProcessableSourceFile(file: GoogleAppsScript.Drive.File): boolean {
  const mime = String(file.getMimeType() || '');
  const lower = file.getName().toLowerCase();
  return mime === MimeType.PDF || mime === MimeType.GOOGLE_DOCS || mime === 'application/vnd.google-apps.shortcut' || lower.endsWith('.pdf');
}

function listTargetPdfs(sourceUrlOrId: string): GoogleAppsScript.Drive.File[] {
  const id = extractIdFromDriveUrl(sourceUrlOrId);
  try {
    const file = DriveApp.getFileById(id);
    return isProcessableSourceFile(file) ? [file] : [];
  } catch (_e) {
    // Try as folder below.
  }

  const files: GoogleAppsScript.Drive.File[] = [];
  const folder = DriveApp.getFolderById(id);
  const it = folder.getFiles();
  while (it.hasNext()) {
    const file = it.next();
    if (isProcessableSourceFile(file)) {
      files.push(file);
    }
  }
  return files;
}

function getProcessedFileIds(): Set<string> {
  const sh = getSheetOrThrow(SHEET_NAMES.log);
  ensureHeaders(sh, LOG_HEADERS);
  const last = sh.getLastRow();
  const processed = new Set<string>();
  if (last < 2) return processed;
  const values = sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues();
  values.forEach((row) => {
    const fileId = String(row[1] || '').trim();
    const result = String(row[3] || '').trim();
    if (fileId && result === 'SUCCESS') processed.add(fileId);
  });
  return processed;
}

function normalizeText(raw: string): string {
  return raw
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function extractTextFromPdfViaOcr(fileId: string): string {
  const file = DriveApp.getFileById(fileId);
  const mime = String(file.getMimeType() || '');
  if (mime === MimeType.GOOGLE_DOCS) {
    return normalizeText(DocumentApp.openById(fileId).getBody().getText());
  }
  if (mime !== MimeType.PDF && !file.getName().toLowerCase().endsWith('.pdf')) {
    throw new Error(`OCR対象がPDFではありません: ${mime} (${file.getName()})`);
  }

  const blob = file.getBlob();
  const blobMime = String(blob.getContentType() || '');
  if (blobMime && blobMime !== MimeType.PDF && blobMime !== 'application/pdf') {
    throw new Error(`OCR対象BlobがPDFではありません: ${blobMime} (${file.getName()})`);
  }

  const tempDocFile = Drive.Files.insert(
    {
      title: `ocr_${fileId}_${new Date().getTime()}`,
      mimeType: MimeType.GOOGLE_DOCS,
    },
    blob,
    { ocr: true, ocrLanguage: 'ja' }
  );
  const docId = tempDocFile.id;
  if (!docId) throw new Error('OCR失敗: Google Doc作成に失敗しました');
  try {
    const text = DocumentApp.openById(docId).getBody().getText();
    return normalizeText(text);
  } finally {
    DriveApp.getFileById(docId).setTrashed(true);
  }
}

function resolveShortcutTarget(file: GoogleAppsScript.Drive.File): GoogleAppsScript.Drive.File {
  const mime = String(file.getMimeType() || '');
  if (mime !== 'application/vnd.google-apps.shortcut') return file;
  const targetId = typeof (file as any).getTargetId === 'function' ? String((file as any).getTargetId() || '') : '';
  if (!targetId) {
    throw new Error(`ショートカットのリンク先IDを取得できません: ${file.getName()}`);
  }
  return DriveApp.getFileById(targetId);
}

function extractTextForSupportedFile(file: GoogleAppsScript.Drive.File): string {
  const resolved = resolveShortcutTarget(file);
  const mime = String(resolved.getMimeType() || '');
  const lowerName = resolved.getName().toLowerCase();

  if (mime === MimeType.GOOGLE_DOCS) {
    return normalizeText(DocumentApp.openById(resolved.getId()).getBody().getText());
  }

  if (mime === MimeType.PDF || lowerName.endsWith('.pdf')) {
    return extractTextFromPdfViaOcr(resolved.getId());
  }

  throw new Error(`未対応ファイル形式: ${mime} (source=${file.getName()}, resolved=${resolved.getName()})`);
}

function extractOne(text: string, labels: string[]): string {
  const lines = text.split('\n').map((l) => l.trim()).filter(Boolean);
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    for (const label of labels) {
      const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const re = new RegExp(`^${escaped}\\s*[:：]?\\s*(.+)$`);
      const m = line.match(re);
      if (m && m[1]) return m[1].trim();
      if (line === label && lines[i + 1]) return lines[i + 1].trim();
    }
  }
  return '';
}

function normalizeGender(value: string): string {
  const s = String(value || '').trim();
  if (!s) return '';
  if (/^(男|male|m)$/i.test(s)) return '男';
  if (/^(女|female|f)$/i.test(s)) return '女';
  return s;
}

function normalizeDateYmd(value: string): string {
  const s = String(value || '').trim();
  if (!s) return '';
  const compact = s.replace(/[年月\/\-.日\s]/g, '');
  if (/^\d{8}$/.test(compact)) return compact;
  const m = s.match(/(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
  if (m) {
    const y = m[1];
    const mo = m[2].padStart(2, '0');
    const d = m[3].padStart(2, '0');
    return `${y}${mo}${d}`;
  }
  return compact;
}

function normalizeTestItem(value: string): string {
  const raw = String(value || '').trim();
  if (!raw) return '';
  for (const [re, normalized] of ITEM_NORMALIZATION) {
    if (re.test(raw)) return normalized;
  }
  return raw.replace(/\s+/g, '');
}

function extractTestItems(text: string): string[] {
  const startMatchers = ['検査項目', '検体項目'];
  const lines = text.split('\n').map((l) => l.trim()).filter(Boolean);
  const items: string[] = [];
  let inItems = false;

  for (const line of lines) {
    if (!inItems && startMatchers.some((k) => line.includes(k))) {
      inItems = true;
      const suffix = line.split(/[：:]/).slice(1).join(':').trim();
      if (suffix) items.push(suffix);
      continue;
    }
    if (!inItems) continue;
    if (/^(備考|連絡|施設|試験名|被験者|採取日|ポイント名)/.test(line)) break;
    if (/^[\-\*・●]/.test(line) || /（.*）/.test(line) || /分離|抽出|株化/.test(line)) {
      items.push(line.replace(/^[\-\*・●]\s*/, '').trim());
    }
  }

  const normalized = items.map(normalizeTestItem).filter(Boolean);
  const dedup = Array.from(new Set(normalized));
  return dedup;
}

function extractKnownItemsFromText(text: string): string[] {
  const compact = text.replace(/\s+/g, '');
  const results: string[] = [];
  if (/Ｄ?Ｎ?Ａ?抽出/.test(compact)) results.push('ＤＮＡ抽出（Ｎ）');
  if (/リンパ球.*株化.*[1１][1１]/.test(compact)) results.push('リンパ球株化１１');
  if (/血清.*分離/.test(compact)) results.push('血清分離（用手法）');
  if (/血漿.*分離/.test(compact)) results.push('血漿分離（用手法）');
  return Array.from(new Set(results));
}

function parseRecordsFromOcrText(text: string): ParsedRecord[] {
  const trialName = extractOne(text, ['試験名', '研究名']) || 'レジストリ研究';
  const facilityCode = extractOne(text, ['施設コード', '医療機関コード']);
  const facilityName = extractOne(text, ['施設名', '医療機関名']);
  const subjectId = extractOne(text, ['被験者番号', '被験者ID', '症例番号']);
  const gender = normalizeGender(extractOne(text, ['性別']));
  const collectionDate = normalizeDateYmd(extractOne(text, ['採取日', '採血日']));
  const pointName = extractOne(text, ['ポイント名', '来院ポイント', 'Visit']) || '初回登録時';
  let items = extractTestItems(text);
  if (items.length === 0) {
    items = extractKnownItemsFromText(text);
  }
  if (items.length === 0 && ITEM_ANCHOR_PATTERNS.some((re) => re.test(text))) {
    items = ['ＤＮＡ抽出（Ｎ）', 'リンパ球株化１１', '血清分離（用手法）', '血漿分離（用手法）'].filter((x) =>
      x.includes('ＤＮＡ')
        ? /Ｄ?Ｎ?Ａ?抽出/.test(text)
        : x.includes('リンパ球')
        ? /リンパ球/.test(text)
        : x.includes('血清')
        ? /血清/.test(text)
        : /血漿/.test(text)
    );
  }

  if (!facilityCode || !facilityName || !subjectId || !collectionDate || items.length === 0) {
    throw new Error('OCR解析に必要な項目を抽出できませんでした');
  }

  return items.map((item) => ({
    試験名: trialName,
    施設コード: facilityCode,
    施設名: facilityName,
    被験者番号: subjectId,
    性別: gender,
    採取日: collectionDate,
    ポイント名: pointName,
    検査項目: item,
  }));
}

function appendRecords(records: ParsedRecord[]): void {
  const sh = getSheetOrThrow(SHEET_NAMES.intake);
  ensureHeaders(sh, INTAKE_HEADERS);
  if (!records.length) return;
  const values = records.map((r) => INTAKE_HEADERS.map((h) => r[h]));
  sh.getRange(sh.getLastRow() + 1, 1, values.length, INTAKE_HEADERS.length).setValues(values);
}

function processUnreadPdfs(): { processed: number; inserted: number; skipped: number; failed: number } {
  const settings = getSettings();
  const files = listTargetPdfs(settings.読み取りPDF格納先);
  const processedSet = getProcessedFileIds();
  let processed = 0;
  let inserted = 0;
  let skipped = 0;
  let failed = 0;

  files.forEach((file) => {
    const fileId = file.getId();
    const fileName = file.getName();
    if (processedSet.has(fileId)) {
      skipped += 1;
      appendLog(fileId, fileName, 'SKIP', 'already processed', 0);
      return;
    }
    try {
      const text = extractTextForSupportedFile(file);
      const records = parseRecordsFromOcrText(text);
      appendRecords(records);
      appendLog(fileId, fileName, 'SUCCESS', 'parsed', records.length);
      processed += 1;
      inserted += records.length;
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      appendLog(fileId, fileName, 'FAIL', message, 0);
      failed += 1;
    }
  });

  return { processed, inserted, skipped, failed };
}

function ensureScheduledTrigger(): void {
  const settings = getSettings();
  const freq = normalizeFrequency(settings.更新頻度);
  const handler = 'runScheduledPipeline';
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
  });

  const builder = ScriptApp.newTrigger(handler).timeBased().atHour(9);
  if (freq === 'hour') builder.everyHours(1).create();
  if (freq === 'day') builder.everyDays(1).create();
  if (freq === 'week') builder.everyWeeks(1).create();
  if (freq === 'month') builder.everyDays(30).create();
}

function createDashboardStats(): DashboardStats {
  const sh = getSheetOrThrow(SHEET_NAMES.intake);
  ensureHeaders(sh, INTAKE_HEADERS);
  const last = sh.getLastRow();
  const stats: DashboardStats = {
    totalRows: 0,
    uniqueSubjects: 0,
    uniqueFacilities: 0,
    byPoint: {},
    byItem: {},
  };
  if (last < 2) return stats;

  const values = sh.getRange(2, 1, last - 1, INTAKE_HEADERS.length).getValues();
  const subjects = new Set<string>();
  const facilities = new Set<string>();

  values.forEach((row) => {
    const facility = String(row[2] || '').trim();
    const subject = String(row[3] || '').trim();
    const point = String(row[6] || '').trim();
    const item = String(row[7] || '').trim();
    stats.totalRows += 1;
    if (subject) subjects.add(subject);
    if (facility) facilities.add(facility);
    if (point) stats.byPoint[point] = (stats.byPoint[point] || 0) + 1;
    if (item) stats.byItem[item] = (stats.byItem[item] || 0) + 1;
  });

  stats.uniqueSubjects = subjects.size;
  stats.uniqueFacilities = facilities.size;
  return stats;
}

function renderDashboard(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashboard = ss.getSheetByName('ダッシュボード');
  if (!dashboard) dashboard = ss.insertSheet('ダッシュボード');
  dashboard.clear();

  const stats = createDashboardStats();
  dashboard.getRange('A1:B1').setValues([['指標', '値']]);
  dashboard.getRange('A2:B5').setValues([
    ['最終更新', nowStr()],
    ['総レコード数', stats.totalRows],
    ['被験者数', stats.uniqueSubjects],
    ['施設数', stats.uniqueFacilities],
  ]);

  const pointRows = Object.keys(stats.byPoint).sort().map((k) => [k, stats.byPoint[k]]);
  dashboard.getRange('D1:E1').setValues([['ポイント名', '件数']]);
  if (pointRows.length) dashboard.getRange(2, 4, pointRows.length, 2).setValues(pointRows);

  const itemRows = Object.keys(stats.byItem).sort((a, b) => stats.byItem[b] - stats.byItem[a]).slice(0, 10).map((k) => [k, stats.byItem[k]]);
  dashboard.getRange('G1:H1').setValues([['検査項目(TOP10)', '件数']]);
  if (itemRows.length) dashboard.getRange(2, 7, itemRows.length, 2).setValues(itemRows);

  dashboard.getCharts().forEach((chart) => dashboard.removeChart(chart));

  if (pointRows.length) {
    const pointChart = dashboard
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dashboard.getRange(1, 4, pointRows.length + 1, 2))
      .setPosition(7, 1, 0, 0)
      .setOption('title', 'ポイント別件数')
      .build();
    dashboard.insertChart(pointChart);
  }

  if (itemRows.length) {
    const itemChart = dashboard
      .newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dashboard.getRange(1, 7, itemRows.length + 1, 2))
      .setPosition(7, 8, 0, 0)
      .setOption('title', '検査項目 TOP10')
      .build();
    dashboard.insertChart(itemChart);
  }
}

function runPipelineCore(showToast: boolean, syncTrigger = true): void {
  const result = processUnreadPdfs();
  renderDashboard();
  if (syncTrigger) {
    ensureScheduledTrigger();
  }
  if (showToast) {
    SpreadsheetApp.getActive().toast(
      `処理完了: 成功${result.processed}件 / 追加${result.inserted}行 / スキップ${result.skipped}件 / 失敗${result.failed}件`,
      'Sample Registration Pipeline',
      8
    );
  }
}

function runScheduledPipeline(): void {
  runPipelineCore(false);
}

function runNow(): void {
  runPipelineCore(true);
}

function manualPdfReadAndUpdateTable(): void {
  runPipelineCore(true, false);
}

function setup(): void {
  ensureScheduledTrigger();
  SpreadsheetApp.getActive().toast('更新頻度に応じたトリガーを設定しました', 'Sample Registration Pipeline', 5);
}

function rebuildDashboard(): void {
  renderDashboard();
  SpreadsheetApp.getActive().toast('ダッシュボードを更新しました', 'Sample Registration Pipeline', 5);
}

function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('受付自動化')
    .addItem('PDF手動読み取り＆表更新', 'manualPdfReadAndUpdateTable')
    .addItem('今すぐ実行（トリガー同期あり）', 'runNow')
    .addItem('スケジュール設定を再作成', 'setup')
    .addItem('ダッシュボード再生成', 'rebuildDashboard')
    .addToUi();
}
