/// <reference path="./types.ts" />
/// <reference path="./config.ts" />

function normalizeCcscMailBody(body: string): string {
  return String(body || '')
    .replace(/\r\n?/g, '\n')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t\u3000]+/g, ' ')
    .replace(/^[ \t]+|[ \t]+$/gm, '')
    .trim();
}

function extractCcscLine(
  body: string,
  labelPattern: string,
  required: boolean,
): string {
  const pattern = new RegExp(
    `^\\s*${labelPattern}\\s*[:：]?\\s*(.*?)\\s*$`,
    'm',
  );
  const match = body.match(pattern);
  const value = match ? String(match[1] || '').trim() : '';
  if (required && !value) {
    throw new Error(`必須項目を抽出できません: ${labelPattern}`);
  }
  return value;
}

function parseCcscJstDateTime(value: string): Date {
  const match = String(value)
    .trim()
    .match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})$/);
  if (!match) {
    throw new Error(`日時の形式が不正です: ${value}`);
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const daysInMonth =
    month >= 1 && month <= 12
      ? new Date(Date.UTC(year, month, 0)).getUTCDate()
      : 0;

  if (
    year < 2000 ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > daysInMonth ||
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59
  ) {
    throw new Error(`日時の値が不正です: ${value}`);
  }

  // Japan does not observe daylight saving time, so JST is always UTC+09:00.
  return new Date(Date.UTC(year, month - 1, day, hour - 9, minute, 0, 0));
}

function parseCcscReservation(rawBody: string): CcscReservation {
  const body = normalizeCcscMailBody(rawBody);
  const usageMatch = body.match(
    /^使用日時\s*[:：]?\s*(\d{4}\/\d{1,2}\/\d{1,2}\s+\d{1,2}:\d{2})\s*から\s*(\d{4}\/\d{1,2}\/\d{1,2}\s+\d{1,2}:\d{2})\s*まで\s*$/m,
  );
  if (!usageMatch) {
    throw new Error('必須項目を抽出できません: 使用日時');
  }

  const attendeeText = extractCcscLine(body, '使用人数', false);
  const attendeeMatch = attendeeText.match(
    /^(\d+)\s*人(?:\s*[（(]\s*(.*?)\s*[）)])?/,
  );
  const start = parseCcscJstDateTime(usageMatch[1]);
  const end = parseCcscJstDateTime(usageMatch[2]);
  if (end.getTime() <= start.getTime()) {
    throw new Error('使用日時の終了時刻が開始時刻以前です');
  }

  return {
    loanId: extractCcscLine(body, '貸出ID', true),
    representativeName: extractCcscLine(body, '代表者氏名', false),
    contact: extractCcscLine(body, '連絡先', false),
    affiliation: extractCcscLine(body, '代表者所属', false),
    location: extractCcscLine(body, '使用場所', true),
    purpose: extractCcscLine(body, '使用目的', false),
    title: extractCcscLine(body, '件\\s*名', false),
    attendeeCount: attendeeMatch ? Number(attendeeMatch[1]) : null,
    attendeeBreakdown:
      attendeeMatch && attendeeMatch[2]
        ? attendeeMatch[2].trim()
        : attendeeText,
    notes: extractCcscLine(body, '備\\s*考', false),
    start,
    end,
  };
}

