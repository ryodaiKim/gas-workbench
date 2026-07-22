/// <reference path="./types.ts" />
/// <reference path="./config.ts" />
/// <reference path="./parser.ts" />

function assertCcscTest(
  condition: boolean,
  message: string,
): void {
  if (!condition) throw new Error(`Test failed: ${message}`);
}

function assertCcscThrows(
  operation: () => void,
  message: string,
): void {
  let didThrow = false;
  try {
    operation();
  } catch (_error) {
    didThrow = true;
  }
  assertCcscTest(didThrow, message);
}

function runCcscSelfTests(): string {
  const sample = [
    '以下の内容で、予約を受け付けました。',
    '',
    '貸出ID      65083',
    '',
    '代表者氏名  テスト 太郎',
    '連絡先      08000000000',
    '代表者所属  大学医学部 学生（所属先：なし）',
    '',
    '使用場所    診察シミュレーション室（5）',
    '',
    '使用目的    個人トレーニング',
    '件　名      自主練',
    '使用人数    2人 （学生 2人）',
    '',
    '備　考      救急蘇生法',
    '',
    '使用日時    2026/08/04 10:00 から 2026/08/04 10:30 まで',
  ].join('\r\n');

  const parsed = parseCcscReservation(sample);
  assertCcscTest(parsed.loanId === '65083', '貸出ID');
  assertCcscTest(
    parsed.representativeName === 'テスト 太郎',
    '代表者氏名',
  );
  assertCcscTest(
    parsed.location === '診察シミュレーション室（5）',
    '使用場所',
  );
  assertCcscTest(parsed.title === '自主練', 'full-width space in 件名');
  assertCcscTest(parsed.attendeeCount === 2, '使用人数');
  assertCcscTest(parsed.attendeeBreakdown === '学生 2人', '人数内訳');
  assertCcscTest(parsed.notes === '救急蘇生法', '備考');
  assertCcscTest(
    parsed.start.toISOString() === '2026-08-04T01:00:00.000Z',
    'JST start conversion',
  );
  assertCcscTest(
    parsed.end.toISOString() === '2026-08-04T01:30:00.000Z',
    'JST end conversion',
  );

  assertCcscThrows(
    () =>
      parseCcscReservation(
        sample.replace(
          '使用日時    2026/08/04 10:00 から 2026/08/04 10:30 まで',
          '',
        ),
      ),
    'missing usage period must fail',
  );
  assertCcscThrows(
    () =>
      parseCcscReservation(
        sample.replace(
          '2026/08/04 10:00 から 2026/08/04 10:30',
          '2026/08/04 11:00 から 2026/08/04 10:30',
        ),
      ),
    'end before start must fail',
  );
  assertCcscThrows(
    () =>
      parseCcscReservation(
        sample.replace('2026/08/04 10:00', '2026/02/30 10:00'),
      ),
    'invalid calendar date must fail',
  );

  const result = 'All CCSC parser tests passed.';
  console.log(result);
  return result;
}
