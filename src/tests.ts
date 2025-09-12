/// <reference path="./main.ts" />
// Lightweight unit tests for date math and labeling

function assertEqual(actual: any, expected: any, message: string) {
  if (actual !== expected) {
    throw new Error(`ASSERT FAILED: ${message} (actual=${actual}, expected=${expected})`);
  }
}

function assertDate(actual: Date, expected: Date, message: string) {
  const tz = Session.getScriptTimeZone();
  const a = Utilities.formatDate(actual, tz, 'yyyy/MM/dd');
  const e = Utilities.formatDate(expected, tz, 'yyyy/MM/dd');
  if (a !== e) throw new Error(`ASSERT FAILED: ${message} (actual=${a}, expected=${e})`);
}

function runUnitTests() {
  const base = new Date('2025-09-03T00:00:00Z');
  // 1M
  assertDate(computePlannedDateFromBase(base, '1M'), addMonths(base, 1), '1M from base');
  // 1M-1W
  assertDate(computePlannedDateFromBase(base, '1M-1W'), addDays(addMonths(base, 1), -7), '1M-1W from base');
  // 12M+1M = 13 months
  assertDate(computePlannedDateFromBase(base, '12M+1M'), addMonths(base, 13), '12M+1M from base');

  // group labels (global function)
  const l1 = groupLabelFrom('1M' as any);
  assertEqual(l1, '登録1ヶ月後評価', 'group label 1M');
  const l3 = groupLabelFrom('3M' as any);
  assertEqual(l3, '登録3ヶ月後評価', 'group label 3M');

  SpreadsheetApp.getActive().toast('単体テスト: すべて成功', 'Visit Management', 5);
}
