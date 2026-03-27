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
  return (
    mime === MimeType.PDF ||
    mime === 'application/pdf' ||
    mime === MimeType.GOOGLE_DOCS ||
    mime === 'application/vnd.google-apps.document' ||
    mime === 'application/vnd.google-apps.shortcut'
  );
}

function listTargetPdfs(sourceUrlOrId: string): GoogleAppsScript.Drive.File[] {
  const id = extractIdFromDriveUrl(sourceUrlOrId);

  // Try as a single processable file first, but skip folders
  try {
    const file = DriveApp.getFileById(id);
    const mime = String(file.getMimeType() || '');
    if (mime !== 'application/vnd.google-apps.folder' && isProcessableSourceFile(file)) {
      return [file];
    }
    // If it's a folder or non-processable file, fall through to folder logic
  } catch (_e) {
    // Not a file — try as folder below
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

function extractTextFromGoogleDoc(fileId: string): string {
  return normalizeText(DocumentApp.openById(fileId).getBody().getText());
}

function ocrPdfViaFetch(fileId: string, fileName: string): string {
  const token = ScriptApp.getOAuthToken();
  const url = `https://www.googleapis.com/drive/v2/files/${fileId}?alt=media`;
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) {
    throw new Error(`PDF取得失敗: HTTP ${response.getResponseCode()} (${fileName})`);
  }
  const blob = response.getBlob().setName(`ocr_${fileId}.pdf`).setContentType('application/pdf');
  const tempDoc = Drive.Files.insert(
    { title: `ocr_${fileId}_${new Date().getTime()}`, mimeType: MimeType.GOOGLE_DOCS },
    blob,
    { ocr: true, ocrLanguage: 'ja' }
  );
  const docId = tempDoc.id;
  if (!docId) throw new Error('OCR失敗: Doc作成失敗');
  try {
    return normalizeText(DocumentApp.openById(docId).getBody().getText());
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
  const resolvedId = resolved.getId();
  const resolvedMime = String(resolved.getMimeType() || '');
  const resolvedName = resolved.getName();

  Logger.log(`[diag] resolved: id=${resolvedId}, name=${resolvedName}, mime=${resolvedMime}`);

  // Strategy 1: Try DocumentApp directly (works for Google Docs AND secretly-converted PDFs)
  try {
    const text = normalizeText(DocumentApp.openById(resolvedId).getBody().getText());
    if (text.length > 50) {
      Logger.log(`[diag] DocumentApp succeeded, length=${text.length}`);
      return text;
    }
  } catch (e) {
    Logger.log(`[diag] DocumentApp failed: ${e instanceof Error ? e.message : String(e)}`);
  }

  // Strategy 2: For PDFs, download raw bytes via UrlFetchApp and OCR
  if (resolvedMime === MimeType.PDF || resolvedMime === 'application/pdf') {
    Logger.log(`[diag] Trying OCR via UrlFetchApp for ${resolvedName}`);
    return ocrPdfViaFetch(resolvedId, resolvedName);
  }

  throw new Error(`テキスト抽出不可: mime=${resolvedMime} (${resolvedName})`);
}

// ---------------------------------------------------------------------------
// Parser: table-format clinical specimen receipt documents (治験受付検体一覧)
// ---------------------------------------------------------------------------

function extractTrialName(text: string): string {
  const m = text.match(/試験名[：:]\s*(.+)/);
  return m ? m[1].trim() : 'レジストリ研究';
}

function extractDocumentYear(text: string): string {
  // Try 治験受付日 or 発信日 for the year
  const patterns = [
    /(?:治験)?受付日[：:\s]*(\d{4})/,
    /発信日[：:\s]*(\d{4})/,
  ];
  for (const p of patterns) {
    const m = text.match(p);
    if (m) return m[1];
  }
  return String(new Date().getFullYear());
}

type FacilityBlock = { name: string; code: string; index: number };

function extractFacilityBlocks(text: string): FacilityBlock[] {
  const blocks: FacilityBlock[] = [];
  const lines = text.split('\n');
  let offset = 0;

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i].trim();
    const lineIndex = offset;
    offset += lines[i].length + 1;

    if (!/(?:病院|大学|センター|医院|クリニック)/.test(line)) continue;

    // Check same line for 5-digit code
    const sameMatch = line.match(/(\d{5})/);
    if (sameMatch) {
      const name = line.replace(/\d{5}/, '').replace(/[*□\s\u3000]/g, '').trim();
      if (name) blocks.push({ name, code: sameMatch[1], index: lineIndex });
      continue;
    }
    // Check next line
    if (i + 1 < lines.length) {
      const nextLine = lines[i + 1].trim();
      const nextMatch = nextLine.match(/^(\d{5})$/);
      if (nextMatch) {
        const name = line.replace(/[*□\s\u3000]/g, '').trim();
        if (name) blocks.push({ name, code: nextMatch[1], index: lineIndex });
      }
    }
  }
  return blocks;
}

function findLastBefore<T extends { index: number }>(items: T[], position: number): T | null {
  let best: T | null = null;
  for (const item of items) {
    if (item.index <= position) best = item;
  }
  return best;
}

function expandItemGroup(rawText: string): string[] {
  const parts = rawText.split(/[・\·、,]/);
  const items: string[] = [];

  for (const part of parts) {
    const t = part.trim();
    if (!t) continue;
    if (/dna|ＤＮＡ|DNA/i.test(t)) {
      items.push('ＤＮＡ抽出（Ｎ）');
    } else if (/株化.*リンパ|リンパ.*株化|リンパ球/i.test(t)) {
      items.push('リンパ球株化１１');
    } else if (/血清/i.test(t)) {
      items.push('血清分離（用手法）');
    } else if (/血漿/i.test(t)) {
      items.push('血漿分離（用手法）');
    } else {
      items.push(t);
    }
  }
  return Array.from(new Set(items));
}

function normalizeGender(value: string): string {
  const s = String(value || '').trim();
  if (!s) return '';
  if (/^(男|male|m)$/i.test(s)) return '男';
  if (/^(女|female|f)$/i.test(s)) return '女';
  return s;
}

type PositionedSubject = { id: string; index: number };
type PositionedGender = { value: string; index: number };
type PositionedPoint = { value: string; index: number };

function collectSubjects(text: string): PositionedSubject[] {
  const results: PositionedSubject[] = [];
  const re = /CIDP-([A-Z]{3})-(\d{4})/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    results.push({ id: m[0], index: m.index });
  }
  // Deduplicate consecutive duplicates (OCR often repeats the same ID)
  return results.filter((item, i) => i === 0 || item.id !== results[i - 1].id || item.index - results[i - 1].index > 50);
}

function collectGenders(text: string): PositionedGender[] {
  const results: PositionedGender[] = [];
  // Match standalone 男/女 (not inside longer words)
  const re = /(?:^|[\s\t])([男女])(?:[\s\t]|$)/gm;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    results.push({ value: m[1], index: m.index });
  }
  return results;
}

function collectPoints(text: string): PositionedPoint[] {
  const results: PositionedPoint[] = [];
  const re = /(初回登録時|追跡時[（(][^）)]*[）)])/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    results.push({ value: m[1], index: m.index });
  }
  return results;
}

type ItemGroupMatch = { rawItems: string; index: number };

function collectItemGroups(text: string): ItemGroupMatch[] {
  const results: ItemGroupMatch[] = [];
  // Match NNN【content】 — bracket may be unclosed in OCR text
  const re = /\d{3}【([^】\n]+)】?/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m[1].trim()) results.push({ rawItems: m[1].trim(), index: m.index });
  }
  return results;
}

function findDateNear(text: string, position: number): { month: string; day: string } | null {
  // Search in a window around the item group (after is more common in table layout)
  const start = Math.max(0, position - 100);
  const end = Math.min(text.length, position + 300);
  const window = text.slice(start, end);

  const patterns = [
    /(\d{1,2})\s*月\s*(\d{1,2})\s*日?/,
    /(\d{1,2})\s+(\d{1,2})(?:\s|$)/,
  ];
  for (const p of patterns) {
    const m = window.match(p);
    if (m) {
      const month = parseInt(m[1], 10);
      const day = parseInt(m[2], 10);
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        return { month: String(month), day: String(day) };
      }
    }
  }
  return null;
}

function parseRecordsFromOcrText(text: string): ParsedRecord[] {
  const trialName = extractTrialName(text);
  const year = extractDocumentYear(text);
  const facilityBlocks = extractFacilityBlocks(text);
  const subjects = collectSubjects(text);
  const genders = collectGenders(text);
  const points = collectPoints(text);
  const itemGroups = collectItemGroups(text);

  Logger.log(`[diag] parse: year=${year}, facilities=${facilityBlocks.length}, subjects=${subjects.length}, genders=${genders.length}, points=${points.length}, itemGroups=${itemGroups.length}`);

  if (itemGroups.length === 0) {
    throw new Error(`テーブル行（【検査項目】）が見つかりません。text preview: ${text.slice(0, 500)}`);
  }

  // Deduplicate facility blocks — keep unique name+code, prefer later (table body) over header
  const seenFacility = new Map<string, FacilityBlock>();
  for (const fb of facilityBlocks) {
    seenFacility.set(fb.code, fb);
  }
  const dedupFacilities = Array.from(seenFacility.values());

  const records: ParsedRecord[] = [];

  for (const ig of itemGroups) {
    // Find nearest subject before this item group
    const subject = findLastBefore(subjects, ig.index);
    // Find nearest facility before this item group (prefer table-body occurrence)
    const facility = findLastBefore(dedupFacilities, ig.index) || dedupFacilities[0] || null;
    // Find nearest gender before this item group
    const gender = findLastBefore(genders, ig.index);
    // Find nearest point before this item group
    const point = findLastBefore(points, ig.index);
    // Find date near this item group
    const date = findDateNear(text, ig.index);

    const subjectId = subject ? subject.id : '';
    const facilityCode = facility ? facility.code : '';
    const facilityName = facility ? facility.name : '';
    const genderStr = gender ? normalizeGender(gender.value) : '';
    const pointStr = point ? point.value : '初回登録時';
    const dateStr = date
      ? `${year}${date.month.padStart(2, '0')}${date.day.padStart(2, '0')}`
      : '';

    const items = expandItemGroup(ig.rawItems);
    for (const item of items) {
      records.push({
        試験名: trialName,
        施設コード: facilityCode,
        施設名: facilityName,
        被験者番号: subjectId,
        性別: genderStr,
        採取日: dateStr,
        ポイント名: pointStr,
        検査項目: item,
      });
    }
  }

  if (records.length === 0) {
    throw new Error('解析結果が0件です');
  }

  // Validate critical fields
  const missing: string[] = [];
  if (records.some((r) => !r.被験者番号)) missing.push('被験者番号');
  if (records.some((r) => !r.施設コード)) missing.push('施設コード');
  if (records.some((r) => !r.採取日)) missing.push('採取日');
  if (missing.length > 0) {
    Logger.log(`[diag] 欠損フィールド: ${missing.join(', ')}`);
  }

  return records;
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
      const sourceMime = String(file.getMimeType() || '');
      let resolvedMime = sourceMime;
      try {
        const resolved = resolveShortcutTarget(file);
        resolvedMime = String(resolved.getMimeType() || '');
      } catch (_) {
        // ignore resolution failure in error path
      }
      const message = e instanceof Error ? e.message : String(e);
      appendLog(fileId, fileName, 'FAIL', `${message} [sourceMime=${sourceMime}, resolvedMime=${resolvedMime}]`, 0);
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

function diagnoseFolderContents(): void {
  const settings = getSettings();
  const sourceUrl = settings.読み取りPDF格納先;
  const id = extractIdFromDriveUrl(sourceUrl);

  const sh = getSheetOrThrow(SHEET_NAMES.log);
  ensureHeaders(sh, LOG_HEADERS);

  appendLog('--', '--', 'SKIP', `[DIAG] 設定値: ${sourceUrl}`, 0);
  appendLog('--', '--', 'SKIP', `[DIAG] 解析ID: ${id}`, 0);

  // Check if ID resolves as a single file or folder
  let isFolder = false;
  try {
    const f = DriveApp.getFileById(id);
    const fMime = String(f.getMimeType() || '');
    appendLog(id, f.getName(), 'SKIP', `[DIAG] IDはファイル: mime=${fMime}`, 0);
    if (fMime === 'application/vnd.google-apps.folder') {
      isFolder = true;
      appendLog('--', '--', 'SKIP', `[DIAG] フォルダとして内容を走査します`, 0);
    } else {
      appendLog('--', '--', 'SKIP', `[DIAG] 単一ファイル → フィルタ: ${isProcessableSourceFile(f) ? 'MATCH' : 'REJECT'}`, 0);
      return;
    }
  } catch (_) {
    appendLog('--', '--', 'SKIP', `[DIAG] IDはファイルではない（フォルダとして試行）`, 0);
    isFolder = true;
  }

  if (!isFolder) return;

  // List ALL files in folder (no filter)
  let folder: GoogleAppsScript.Drive.Folder;
  try {
    folder = DriveApp.getFolderById(id);
  } catch (e) {
    appendLog('--', '--', 'FAIL', `[DIAG] フォルダ取得失敗: ${e instanceof Error ? e.message : String(e)}`, 0);
    return;
  }

  appendLog('--', folder.getName(), 'SKIP', `[DIAG] フォルダ名: ${folder.getName()}`, 0);

  const it = folder.getFiles();
  let total = 0;
  let accepted = 0;
  while (it.hasNext()) {
    const file = it.next();
    const mime = String(file.getMimeType() || '');
    const passes = isProcessableSourceFile(file);
    const tag = passes ? 'MATCH' : 'REJECT';
    appendLog(file.getId(), file.getName(), 'SKIP', `[DIAG][${tag}] mime=${mime}`, 0);
    total += 1;
    if (passes) accepted += 1;
  }

  appendLog('--', '--', 'SKIP', `[DIAG] 合計: ${total}件, フィルタ通過: ${accepted}件, 除外: ${total - accepted}件`, 0);

  // Also check processed set
  const processedSet = getProcessedFileIds();
  appendLog('--', '--', 'SKIP', `[DIAG] 処理済み(SUCCESS)ファイル数: ${processedSet.size}件`, 0);

  SpreadsheetApp.getActive().toast(
    `診断完了: フォルダ内${total}件, フィルタ通過${accepted}件 → ログシートを確認してください`,
    'Diagnosis',
    10
  );
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
    .addSeparator()
    .addItem('【診断】フォルダ内容を検査', 'diagnoseFolderContents')
    .addToUi();
}
