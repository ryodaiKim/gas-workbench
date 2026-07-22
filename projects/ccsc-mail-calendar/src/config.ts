const CCSC_CONFIG = Object.freeze({
  expectedOwnerEmail: '23mb1095@student.gs.chiba-u.jp',
  guestEmail: '22m1084c@student.gs.chiba-u.jp',
  sourceEmail: 'byo-hp@system.ho.u-chiba.jp',
  subjectFragment: '【CCSC】予約申請 確認証および貸出証・利用記録',
  timeZone: 'Asia/Tokyo',

  // Pin writes to the university account's primary calendar. This avoids any
  // ambiguity from multi-login/default-calendar resolution.
  calendarId: '23mb1095@student.gs.chiba-u.jp',

  // Apps Script has no Gmail "message received" trigger. A one-minute clock
  // trigger searches this recent window instead.
  searchLookbackDays: 30,
  maxMessagesPerRun: 100,
  triggerEveryMinutes: 1,

  eventTitlePrefix: '【CCSC予約確定】',
  legacyProvisionalTitlePrefix: '【CCSC仮予約】',
  organizerPopupReminderMinutes: 30,
  sendCalendarInvites: true,

  processedPropertyPrefix: 'CCSC_PROCESSED_',
  processedPropertyRetentionDays: 60,
  eventLoanIdTag: 'ccscLoanId',
  eventMessageIdTag: 'ccscMessageId',
  eventReservationFingerprintTag: 'ccscReservationFingerprint',
  emailLinkMigrationProperty: 'CCSC_MIGRATION_EMAIL_LINK_V1',
});
