/// <reference path="./types.ts" />
/// <reference path="./config.ts" />
/// <reference path="./parser.ts" />

function requireCcscGmailService(): GoogleAppsScript.Gmail {
  if (typeof Gmail === 'undefined' || !Gmail) {
    throw new Error(
      'Advanced Gmail service is not enabled. Check appsscript.json dependencies.',
    );
  }
  return Gmail;
}

function getCcscHeader(
  payload: GoogleAppsScript.Gmail.Schema.MessagePart | undefined,
  name: string,
): string {
  const target = name.toLowerCase();
  const header = (payload?.headers || []).find(
    (item) => String(item.name || '').toLowerCase() === target,
  );
  return String(header?.value || '');
}

function getCcscMimeCharset(
  part: GoogleAppsScript.Gmail.Schema.MessagePart,
): string {
  const contentType = getCcscHeader(part, 'Content-Type');
  const match = contentType.match(
    /charset\s*=\s*(?:"([^"]+)"|'([^']+)'|([^;\s]+))/i,
  );
  return String(match?.[1] || match?.[2] || match?.[3] || 'UTF-8').trim();
}

function normalizeCcscBytes(values: unknown[]): number[] | null {
  const numbers = values.map((value) => Number(value));
  if (
    numbers.some(
      (value) =>
        !Number.isInteger(value) || value < -128 || value > 255,
    )
  ) {
    return null;
  }
  // Apps Script Byte[] values are signed. Normalize unsigned representations.
  return numbers.map((value) => (value > 127 ? value - 256 : value));
}

function getCcscByteArray(data: unknown): number[] | null {
  if (Array.isArray(data)) return normalizeCcscBytes(data);
  if (!data || typeof data !== 'object') return null;

  const arrayLike = data as {
    length?: unknown;
    [index: number]: unknown;
  };
  const length = Number(arrayLike.length);
  if (!Number.isInteger(length) || length <= 0 || length > 10_000_000) {
    return null;
  }
  const values: unknown[] = [];
  for (let index = 0; index < length; index += 1) {
    values.push(arrayLike[index]);
  }
  return normalizeCcscBytes(values);
}

function describeCcscBodyData(data: unknown): string {
  const raw = String(data ?? '');
  const invalidBase64Chars = (
    raw.match(/[^A-Za-z0-9_+\-\/=\s]/g) || []
  ).length;
  return [
    `type=${typeof data}`,
    `array=${Array.isArray(data)}`,
    `length=${raw.length}`,
    `nonAscii=${/[^\x00-\x7F]/.test(raw)}`,
    `commas=${(raw.match(/,/g) || []).length}`,
    `invalidBase64Chars=${invalidBase64Chars}`,
  ].join(',');
}

function decodeCcscBodyData(data: unknown, charset: string): string {
  if (data === null || data === undefined || data === '') return '';

  const directBytes = getCcscByteArray(data);
  if (directBytes) {
    return Utilities.newBlob(directBytes).getDataAsString(charset);
  }

  const raw = String(data);
  // Some Google API wrappers deserialize a bytes field before returning it.
  if (
    raw.includes('貸出ID') ||
    raw.includes('使用日時') ||
    raw.includes('予約申請')
  ) {
    return raw;
  }

  const numericText = raw.replace(/\s+/g, '');
  if (/^-?\d+(?:,-?\d+)+$/.test(numericText)) {
    const parsedBytes = normalizeCcscBytes(numericText.split(','));
    if (parsedBytes) {
      return Utilities.newBlob(parsedBytes).getDataAsString(charset);
    }
  }

  const compact = raw.replace(/\s+/g, '');
  const withoutPadding = compact.replace(/=+$/g, '');
  const remainder = withoutPadding.length % 4;
  if (remainder === 1) {
    throw new Error(
      `Invalid Gmail MIME body (${describeCcscBodyData(data)})`,
    );
  }
  const padded =
    withoutPadding + '='.repeat((4 - remainder) % 4);

  let bytes: number[];
  try {
    bytes = Utilities.base64DecodeWebSafe(padded);
  } catch (webSafeError) {
    // Some Apps Script runtimes are stricter about web-safe input. Convert to
    // standard base64 as a compatibility fallback.
    const standard = padded.replace(/-/g, '+').replace(/_/g, '/');
    try {
      bytes = Utilities.base64Decode(standard);
    } catch (standardError) {
      throw new Error(
        `Could not decode Gmail MIME body (${describeCcscBodyData(data)}): ` +
          `${String(webSafeError)}; fallback=${String(standardError)}`,
      );
    }
  }
  try {
    return Utilities.newBlob(bytes).getDataAsString(charset);
  } catch (error) {
    console.warn(
      `Could not decode MIME body as ${charset}; falling back to UTF-8: ${String(
        error,
      )}`,
    );
    return Utilities.newBlob(bytes).getDataAsString('UTF-8');
  }
}

function collectCcscMimeParts(
  part: GoogleAppsScript.Gmail.Schema.MessagePart | undefined,
  plainParts: string[],
  htmlParts: string[],
): void {
  if (!part) return;
  const mimeType = String(part.mimeType || '').toLowerCase();
  const data = part.body?.data as unknown;
  const isAttachment = Boolean(part.filename);
  const hasData = data !== null && data !== undefined && data !== '';

  if (!isAttachment && hasData && mimeType === 'text/plain') {
    plainParts.push(decodeCcscBodyData(data, getCcscMimeCharset(part)));
  } else if (!isAttachment && hasData && mimeType === 'text/html') {
    htmlParts.push(decodeCcscBodyData(data, getCcscMimeCharset(part)));
  }

  (part.parts || []).forEach((child) => {
    collectCcscMimeParts(child, plainParts, htmlParts);
  });
}

function ccscHtmlToPlainText(html: string): string {
  return String(html || '')
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\s*\/p\s*>/gi, '\n')
    .replace(/<\s*\/div\s*>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}

function getCcscPlainBody(
  payload: GoogleAppsScript.Gmail.Schema.MessagePart | undefined,
): string {
  const plainParts: string[] = [];
  const htmlParts: string[] = [];
  collectCcscMimeParts(payload, plainParts, htmlParts);
  if (plainParts.length) return plainParts.join('\n');
  if (htmlParts.length) return ccscHtmlToPlainText(htmlParts.join('\n'));
  return '';
}

function toCcscMailMessage(
  message: GoogleAppsScript.Gmail.Schema.Message,
): CcscMailMessage {
  const id = String(message.id || '');
  if (!id) throw new Error('Gmail message has no ID');
  const timestamp = Number(message.internalDate || 0);
  return {
    id,
    threadId: String(message.threadId || ''),
    from: getCcscHeader(message.payload, 'From'),
    to: getCcscHeader(message.payload, 'To'),
    subject: getCcscHeader(message.payload, 'Subject'),
    body: getCcscPlainBody(message.payload),
    receivedAt: timestamp > 0 ? new Date(timestamp) : new Date(),
  };
}

function isExpectedCcscMessage(message: CcscMailMessage): boolean {
  const from = message.from.toLowerCase();
  // The Gmail API search already applies the subject query. Avoid checking the
  // Subject header again because non-ASCII subjects can be RFC 2047 encoded.
  return from.includes(CCSC_CONFIG.sourceEmail.toLowerCase());
}

function listRecentCcscMessageReferences(): GoogleAppsScript.Gmail.Schema.Message[] {
  const gmail = requireCcscGmailService();
  const query = [
    `from:${CCSC_CONFIG.sourceEmail}`,
    `subject:"${CCSC_CONFIG.subjectFragment}"`,
    `newer_than:${CCSC_CONFIG.searchLookbackDays}d`,
  ].join(' ');
  const references: GoogleAppsScript.Gmail.Schema.Message[] = [];
  let pageToken: string | undefined;

  do {
    const remaining = CCSC_CONFIG.maxMessagesPerRun - references.length;
    if (remaining <= 0) break;
    const response = gmail.Users.Messages.list('me', {
      q: query,
      maxResults: Math.min(100, remaining),
      pageToken,
    });
    references.push(...(response.messages || []));
    pageToken = response.nextPageToken;
  } while (pageToken && references.length < CCSC_CONFIG.maxMessagesPerRun);

  return references;
}

function loadCcscMessages(
  references: GoogleAppsScript.Gmail.Schema.Message[],
): CcscMailMessage[] {
  const gmail = requireCcscGmailService();
  const messages = references
    .map((reference) => {
      const id = String(reference.id || '');
      if (!id) return null;
      const fullMessage = gmail.Users.Messages.get('me', id, {
        format: 'full',
      });
      return toCcscMailMessage(fullMessage);
    })
    .filter((message): message is CcscMailMessage => Boolean(message))
    .filter(isExpectedCcscMessage);

  messages.sort(
    (left, right) => left.receivedAt.getTime() - right.receivedAt.getTime(),
  );
  return messages;
}

function getCcscCalendar(): GoogleAppsScript.Calendar.Calendar {
  if (!CCSC_CONFIG.calendarId) return CalendarApp.getDefaultCalendar();
  const calendar = CalendarApp.getCalendarById(CCSC_CONFIG.calendarId);
  if (!calendar) {
    throw new Error(
      `Configured calendar is not accessible: ${CCSC_CONFIG.calendarId}`,
    );
  }
  return calendar;
}

function requireCcscCalendarApi(): GoogleAppsScript.Calendar {
  if (typeof Calendar === 'undefined' || !Calendar) {
    throw new Error(
      'Advanced Calendar service is not enabled. Check appsscript.json dependencies.',
    );
  }
  return Calendar;
}

function getCcscCalendarApiId(): string {
  // Use the exact CalendarApp target rather than assuming that the Advanced
  // Calendar API's `primary` alias resolves to the same default calendar.
  return getCcscCalendar().getId();
}

function ccscMessagePropertyKey(messageId: string): string {
  return `${CCSC_CONFIG.processedPropertyPrefix}MESSAGE_${messageId}`;
}

function ccscLoanPropertyKey(loanId: string): string {
  const safeLoanId = loanId.replace(/[^A-Za-z0-9_-]/g, '_');
  return `${CCSC_CONFIG.processedPropertyPrefix}LOAN_${safeLoanId}`;
}

function readCcscProcessedRecord(
  key: string,
): CcscProcessedRecord | null {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) return null;
  try {
    return JSON.parse(value) as CcscProcessedRecord;
  } catch (_error) {
    return null;
  }
}

function saveCcscProcessedRecord(
  messageId: string,
  loanId: string,
  eventId: string,
  duplicate: boolean,
  apiEventId = '',
): void {
  const previousLoanRecord = readCcscProcessedRecord(
    ccscLoanPropertyKey(loanId),
  );
  const record: CcscProcessedRecord = {
    processedAt: new Date().toISOString(),
    loanId,
    eventId,
    apiEventId: apiEventId || previousLoanRecord?.apiEventId || undefined,
    duplicate,
  };
  const value = JSON.stringify(record);
  PropertiesService.getScriptProperties().setProperties({
    [ccscMessagePropertyKey(messageId)]: value,
    [ccscLoanPropertyKey(loanId)]: value,
  });
}

function cleanOldCcscProcessedRecords(): void {
  const properties = PropertiesService.getScriptProperties();
  const all = properties.getProperties();
  const cutoff =
    Date.now() -
    CCSC_CONFIG.processedPropertyRetentionDays * 24 * 60 * 60 * 1000;

  Object.keys(all).forEach((key) => {
    if (!key.startsWith(CCSC_CONFIG.processedPropertyPrefix)) return;
    try {
      const record = JSON.parse(all[key]) as Partial<CcscProcessedRecord>;
      const processedAt = Date.parse(String(record.processedAt || ''));
      if (!Number.isFinite(processedAt) || processedAt < cutoff) {
        properties.deleteProperty(key);
      }
    } catch (_error) {
      properties.deleteProperty(key);
    }
  });
}

function findExistingCcscEvent(
  calendar: GoogleAppsScript.Calendar.Calendar,
  reservation: CcscReservation,
): GoogleAppsScript.Calendar.CalendarEvent | null {
  const stored = readCcscProcessedRecord(
    ccscLoanPropertyKey(reservation.loanId),
  );
  if (stored?.eventId) {
    const storedEvent = calendar.getEventById(stored.eventId);
    if (storedEvent) return storedEvent;
  }

  const oneMinute = 60 * 1000;
  const nearbyEvents = calendar.getEvents(
    new Date(reservation.start.getTime() - oneMinute),
    new Date(reservation.end.getTime() + oneMinute),
  );
  return (
    nearbyEvents.find(
      (event) =>
        event.getTag(CCSC_CONFIG.eventLoanIdTag) === reservation.loanId,
    ) || null
  );
}

function buildCcscEventTitle(reservation: CcscReservation): string {
  const base = reservation.title || reservation.purpose || '自主練';
  const detail =
    reservation.notes && reservation.notes !== base
      ? `｜${reservation.notes}`
      : '';
  return `${CCSC_CONFIG.eventTitlePrefix}${base}${detail}`;
}

function hashCcscString(value: string): string {
  let first = 0x811c9dc5;
  let second = 0x9e3779b9;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    first = Math.imul(first ^ code, 0x01000193);
    second = Math.imul(second ^ (code + index), 0x85ebca6b);
  }
  return [first, second]
    .map((part) => (part >>> 0).toString(16).padStart(8, '0'))
    .join('');
}

function buildCcscReservationFingerprint(
  reservation: CcscReservation,
): string {
  return hashCcscString(
    JSON.stringify([
      reservation.loanId,
      reservation.representativeName,
      reservation.contact,
      reservation.affiliation,
      reservation.location,
      reservation.purpose,
      reservation.title,
      reservation.attendeeCount,
      reservation.attendeeBreakdown,
      reservation.notes,
      reservation.start.toISOString(),
      reservation.end.toISOString(),
    ]),
  );
}

function buildCcscGmailThreadUrl(message: CcscMailMessage): string {
  const threadOrMessageId = message.threadId || message.id;
  return `https://mail.google.com/mail/u/0/#all/${encodeURIComponent(
    threadOrMessageId,
  )}`;
}

function buildCcscEventDescription(
  reservation: CcscReservation,
  message: CcscMailMessage,
): string {
  const attendeeText =
    reservation.attendeeCount === null
      ? reservation.attendeeBreakdown
      : `${reservation.attendeeCount}人${
          reservation.attendeeBreakdown
            ? `（${reservation.attendeeBreakdown}）`
            : ''
        }`;
  const receivedAt = Utilities.formatDate(
    message.receivedAt,
    CCSC_CONFIG.timeZone,
    'yyyy/MM/dd HH:mm:ss',
  );

  return [
    'CCSC予約の確認証・貸出証・利用記録メールに基づく確定予定です。',
    '当日は元メールをプリントアウトし、利用記録を記入・提出してください。',
    '',
    `貸出ID: ${reservation.loanId}`,
    `代表者氏名: ${reservation.representativeName}`,
    `連絡先: ${reservation.contact}`,
    `代表者所属: ${reservation.affiliation}`,
    `使用場所: ${reservation.location}`,
    `使用目的: ${reservation.purpose}`,
    `件名: ${reservation.title}`,
    `使用人数: ${attendeeText}`,
    `備考: ${reservation.notes}`,
    '',
    `元メール受信日時: ${receivedAt}`,
    `元メールID: ${message.id}`,
    `元メールを開く: ${buildCcscGmailThreadUrl(message)}`,
    'この予定はCCSCメール連携により自動作成されました。',
  ].join('\n');
}

function chooseCcscCalendarApiEvent(
  items: GoogleAppsScript.Calendar.Schema.Event[],
  calendarEvent: GoogleAppsScript.Calendar.CalendarEvent,
  reservation: CcscReservation,
): GoogleAppsScript.Calendar.Schema.Event | null {
  if (!items.length) return null;
  const iCalUid = calendarEvent.getId();
  const exactUid = items.find(
    (item) => String(item.iCalUID || '') === iCalUid,
  );
  if (exactUid) return exactUid;

  const tagged = items.find(
    (item) =>
      String(
        item.extendedProperties?.private?.[CCSC_CONFIG.eventLoanIdTag] || '',
      ) === reservation.loanId,
  );
  if (tagged) return tagged;

  const loanText = `貸出ID: ${reservation.loanId}`;
  const described = items.find((item) =>
    String(item.description || '').includes(loanText),
  );
  if (described) return described;

  // The time-window query is deliberately narrow. Accept its sole result as a
  // final mapping fallback, but never guess when multiple events are present.
  return items.length === 1 ? items[0] : null;
}

function findCcscCalendarApiEvent(
  calendarEvent: GoogleAppsScript.Calendar.CalendarEvent,
  reservation: CcscReservation,
): GoogleAppsScript.Calendar.Schema.Event | null {
  const calendarApi = requireCcscCalendarApi();
  const calendarId = getCcscCalendarApiId();
  const stored = readCcscProcessedRecord(
    ccscLoanPropertyKey(reservation.loanId),
  );
  if (stored?.apiEventId) {
    try {
      return calendarApi.Events.get(calendarId, stored.apiEventId);
    } catch (error) {
      console.warn(
        `Stored Calendar API event ID is stale for loan ID ${reservation.loanId}: ${String(
          error,
        )}`,
      );
    }
  }

  const directResponse = calendarApi.Events.list(calendarId, {
    iCalUID: calendarEvent.getId(),
    maxResults: 10,
    showDeleted: false,
    singleEvents: true,
  });
  const directMatch = chooseCcscCalendarApiEvent(
    directResponse.items || [],
    calendarEvent,
    reservation,
  );
  if (directMatch) return directMatch;

  // CalendarApp event tags are custom metadata. Looking up the loan-ID tag
  // avoids relying exclusively on iCalUID-to-API-ID translation.
  const taggedResponse = calendarApi.Events.list(calendarId, {
    privateExtendedProperty: `${CCSC_CONFIG.eventLoanIdTag}=${reservation.loanId}`,
    maxResults: 10,
    showDeleted: false,
    singleEvents: true,
  });
  const taggedMatch = chooseCcscCalendarApiEvent(
    taggedResponse.items || [],
    calendarEvent,
    reservation,
  );
  if (taggedMatch) return taggedMatch;

  const oneMinute = 60 * 1000;
  const existingStart = calendarEvent.getStartTime();
  const existingEnd = calendarEvent.getEndTime();
  const timeResponse = calendarApi.Events.list(calendarId, {
    timeMin: new Date(existingStart.getTime() - oneMinute).toISOString(),
    timeMax: new Date(existingEnd.getTime() + oneMinute).toISOString(),
    maxResults: 50,
    showDeleted: false,
    singleEvents: true,
  });
  return chooseCcscCalendarApiEvent(
    timeResponse.items || [],
    calendarEvent,
    reservation,
  );
}

function updateCcscEventWithCalendarApp(
  calendarEvent: GoogleAppsScript.Calendar.CalendarEvent,
  reservation: CcscReservation,
  message: CcscMailMessage,
): void {
  calendarEvent
    .setTitle(buildCcscEventTitle(reservation))
    .setDescription(buildCcscEventDescription(reservation, message))
    .setLocation(reservation.location)
    .setTime(reservation.start, reservation.end);

  const guestEmail = CCSC_CONFIG.guestEmail.toLowerCase();
  const hasGuest = calendarEvent
    .getGuestList()
    .some((guest) => guest.getEmail().toLowerCase() === guestEmail);
  if (!hasGuest) calendarEvent.addGuest(CCSC_CONFIG.guestEmail);
}

function synchronizeConfirmedCcscEvent(
  calendarEvent: GoogleAppsScript.Calendar.CalendarEvent,
  reservation: CcscReservation,
  message: CcscMailMessage,
): boolean {
  const fingerprint = buildCcscReservationFingerprint(reservation);
  const wasProvisional = calendarEvent
    .getTitle()
    .startsWith(CCSC_CONFIG.legacyProvisionalTitlePrefix);
  if (
    !wasProvisional &&
    calendarEvent.getTag(CCSC_CONFIG.eventReservationFingerprintTag) ===
    fingerprint
  ) {
    return false;
  }

  const apiEvent = findCcscCalendarApiEvent(calendarEvent, reservation);
  if (!apiEvent?.id) {
    console.warn(
      `Calendar API event ID could not be resolved for loan ID ${reservation.loanId}; ` +
        'using CalendarApp update fallback.',
    );
    updateCcscEventWithCalendarApp(calendarEvent, reservation, message);
  } else {
    requireCcscCalendarApi().Events.patch(
      {
        summary: buildCcscEventTitle(reservation),
        description: buildCcscEventDescription(reservation, message),
        location: reservation.location,
        start: {
          dateTime: reservation.start.toISOString(),
          timeZone: CCSC_CONFIG.timeZone,
        },
        end: {
          dateTime: reservation.end.toISOString(),
          timeZone: CCSC_CONFIG.timeZone,
        },
      },
      getCcscCalendarApiId(),
      apiEvent.id,
      { sendUpdates: 'all' },
    );
  }

  try {
    calendarEvent.setTag(CCSC_CONFIG.eventMessageIdTag, message.id);
    calendarEvent.setTag(
      CCSC_CONFIG.eventReservationFingerprintTag,
      fingerprint,
    );
  } catch (error) {
    console.warn(
      `Could not update confirmation tags for ${reservation.loanId}: ${String(
        error,
      )}`,
    );
  }
  console.log(
    `CCSC confirmed event ${reservation.loanId} was ${
      wasProvisional ? 'upgraded from provisional' : 'updated'
    }.`,
  );
  return true;
}

function backfillCcscEmailLinksInternal(): CcscEmailLinkBackfillSummary {
  const summary: CcscEmailLinkBackfillSummary = {
    matched: 0,
    updated: 0,
    alreadyLinked: 0,
    missingEvents: 0,
    failed: 0,
  };
  const calendar = getCcscCalendar();
  const calendarApi = requireCcscCalendarApi();
  const messages = loadCcscMessages(listRecentCcscMessageReferences());
  const seenLoanIds = new Set<string>();

  messages.forEach((message) => {
    try {
      const reservation = parseCcscReservation(message.body);
      if (seenLoanIds.has(reservation.loanId)) return;
      seenLoanIds.add(reservation.loanId);
      summary.matched += 1;

      const calendarEvent = findExistingCcscEvent(calendar, reservation);
      if (!calendarEvent) {
        summary.missingEvents += 1;
        return;
      }

      const apiEvent = findCcscCalendarApiEvent(
        calendarEvent,
        reservation,
      );
      if (!apiEvent?.id) {
        summary.missingEvents += 1;
        return;
      }

      const description = buildCcscEventDescription(reservation, message);
      if (String(apiEvent.description || '') === description) {
        summary.alreadyLinked += 1;
        return;
      }

      calendarApi.Events.patch(
        { description },
        getCcscCalendarApiId(),
        apiEvent.id,
        { sendUpdates: 'none' },
      );
      summary.updated += 1;
    } catch (error) {
      summary.failed += 1;
      console.error(
        `Could not backfill Gmail link for message ${message.id}: ${String(
          error,
        )}`,
      );
    }
  });

  console.log(`CCSC email-link backfill: ${JSON.stringify(summary)}`);
  return summary;
}

function ensureCcscEmailLinksBackfilled(): void {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(CCSC_CONFIG.emailLinkMigrationProperty)) return;

  const summary = backfillCcscEmailLinksInternal();
  if (summary.failed > 0) {
    throw new Error(
      `CCSC email-link backfill failed for ${summary.failed} event(s).`,
    );
  }
  properties.setProperty(
    CCSC_CONFIG.emailLinkMigrationProperty,
    new Date().toISOString(),
  );
}

/**
 * Idempotently adds Gmail links to recent CCSC events without sending guest
 * update notifications.
 */
function backfillCcscEmailLinks(): CcscEmailLinkBackfillSummary {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    throw new Error('Another CCSC processing run is active. Try again shortly.');
  }
  try {
    const summary = backfillCcscEmailLinksInternal();
    if (summary.failed === 0) {
      PropertiesService.getScriptProperties().setProperty(
        CCSC_CONFIG.emailLinkMigrationProperty,
        new Date().toISOString(),
      );
    }
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function createCcscCalendarEvent(
  reservation: CcscReservation,
  message: CcscMailMessage,
): CcscCreatedEvent {
  const created = requireCcscCalendarApi().Events.insert(
    {
      summary: buildCcscEventTitle(reservation),
      description: buildCcscEventDescription(reservation, message),
      location: reservation.location,
      start: {
        dateTime: reservation.start.toISOString(),
        timeZone: CCSC_CONFIG.timeZone,
      },
      end: {
        dateTime: reservation.end.toISOString(),
        timeZone: CCSC_CONFIG.timeZone,
      },
      attendees: [{ email: CCSC_CONFIG.guestEmail }],
      reminders: {
        useDefault: false,
        overrides: [
          {
            method: 'popup',
            minutes: CCSC_CONFIG.organizerPopupReminderMinutes,
          },
        ],
      },
      extendedProperties: {
        private: {
          [CCSC_CONFIG.eventLoanIdTag]: reservation.loanId,
          [CCSC_CONFIG.eventMessageIdTag]: message.id,
          [CCSC_CONFIG.eventReservationFingerprintTag]:
            buildCcscReservationFingerprint(reservation),
        },
      },
    },
    getCcscCalendarApiId(),
    {
      sendUpdates: CCSC_CONFIG.sendCalendarInvites ? 'all' : 'none',
    },
  );
  const apiEventId = String(created.id || '');
  const eventId = String(created.iCalUID || '');
  if (!apiEventId || !eventId) {
    throw new Error(
      `Calendar API created loan ID ${reservation.loanId} without complete identifiers.`,
    );
  }
  console.log(
    `Created CCSC event ${reservation.loanId} in ${getCcscCalendarApiId()} ` +
      `(apiEventId=${apiEventId}, iCalUID=${eventId}).`,
  );
  return { eventId, apiEventId };
}

/**
 * Main one-minute trigger handler.
 *
 * Searches matching messages, creates exactly one event per reservation,
 * adds the configured guest, and requests a Calendar invitation email.
 */
function processCcscReservationEmailsInternal(
  includePreviouslyProcessed: boolean,
): CcscProcessingSummary {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    console.log('Another CCSC processing run is still active; skipping.');
    return {
      matched: 0,
      created: 0,
      updated: 0,
      duplicates: 0,
      alreadyProcessed: 0,
      failed: 0,
    };
  }

  const summary: CcscProcessingSummary = {
    matched: 0,
    created: 0,
    updated: 0,
    duplicates: 0,
    alreadyProcessed: 0,
    failed: 0,
  };
  const failures: string[] = [];

  try {
    cleanOldCcscProcessedRecords();
    const calendar = getCcscCalendar();
    const references = listRecentCcscMessageReferences();
    summary.matched = references.length;
    const pendingReferences = references.filter((reference) => {
      const messageId = String(reference.id || '');
      if (!messageId) return false;
      if (
        !includePreviouslyProcessed &&
        readCcscProcessedRecord(ccscMessagePropertyKey(messageId))
      ) {
        summary.alreadyProcessed += 1;
        return false;
      }
      return true;
    });
    // Fetch full MIME bodies only for messages that have not been processed.
    // The one-minute poll still performs a cheap ID search, but does not
    // repeatedly download every historical email in the lookback window.
    const messages = loadCcscMessages(pendingReferences);

    messages.forEach((message) => {
      try {
        const reservation = parseCcscReservation(message.body);
        const existingEvent = findExistingCcscEvent(calendar, reservation);
        if (existingEvent) {
          const updated = synchronizeConfirmedCcscEvent(
            existingEvent,
            reservation,
            message,
          );
          saveCcscProcessedRecord(
            message.id,
            reservation.loanId,
            existingEvent.getId(),
            !updated,
          );
          if (updated) {
            summary.updated += 1;
          } else {
            summary.duplicates += 1;
          }
          return;
        }

        const event = createCcscCalendarEvent(reservation, message);
        saveCcscProcessedRecord(
          message.id,
          reservation.loanId,
          event.eventId,
          false,
          event.apiEventId,
        );
        summary.created += 1;
      } catch (error) {
        summary.failed += 1;
        failures.push(`${message.id}: ${String(error)}`);
        console.error(
          `CCSC message ${message.id} could not be processed: ${String(error)}`,
        );
      }
    });

    console.log(
      `CCSC ${
        includePreviouslyProcessed ? 'reconciliation' : 'processing'
      } result: ${JSON.stringify(summary)}`,
    );
    if (failures.length) {
      throw new Error(
        `Failed to process ${failures.length} CCSC message(s): ${failures.join(
          ' | ',
        )}`,
      );
    }
    ensureCcscEmailLinksBackfilled();
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function processCcscReservationEmails(): CcscProcessingSummary {
  return processCcscReservationEmailsInternal(false);
}

/**
 * Rechecks every matching confirmation in the lookback window, including
 * messages recorded as processed. Missing events are recreated; existing
 * events are left unchanged unless the confirmation contents changed.
 */
function reconcileCcscReservationEmails(): CcscProcessingSummary {
  return processCcscReservationEmailsInternal(true);
}

function requireExpectedCcscOwner(): string {
  const profile = requireCcscGmailService().Users.getProfile('me');
  const actualEmail = String(profile.emailAddress || '').toLowerCase();
  const expectedEmail = CCSC_CONFIG.expectedOwnerEmail.toLowerCase();
  if (actualEmail !== expectedEmail) {
    throw new Error(
      `Wrong Google account. Expected ${expectedEmail}, but the script is authorized as ${
        actualEmail || '(unknown)'
      }.`,
    );
  }
  return actualEmail;
}

function installCcscAutomationTrigger(): void {
  ScriptApp.newTrigger('processCcscReservationEmails')
    .timeBased()
    .everyMinutes(CCSC_CONFIG.triggerEveryMinutes)
    .create();
}

function getCcscTrackedEventIds(): string[] {
  const values = PropertiesService.getScriptProperties().getProperties();
  const eventIds = new Set<string>();
  Object.keys(values).forEach((key) => {
    if (!key.startsWith(CCSC_CONFIG.processedPropertyPrefix)) return;
    try {
      const record = JSON.parse(values[key]) as Partial<CcscProcessedRecord>;
      if (record.eventId) eventIds.add(String(record.eventId));
    } catch (_error) {
      // Malformed app records are removed later with the other CCSC state.
    }
  });
  return [...eventIds];
}

function collectCcscEventsForReset(
  messages: CcscMailMessage[],
): GoogleAppsScript.Calendar.CalendarEvent[] {
  const targetCalendar = getCcscCalendar();
  const defaultCalendar = CalendarApp.getDefaultCalendar();
  const calendars = [targetCalendar];
  if (defaultCalendar.getId() !== targetCalendar.getId()) {
    calendars.push(defaultCalendar);
  }

  const found = new Map<
    string,
    GoogleAppsScript.Calendar.CalendarEvent
  >();
  const addEvent = (
    _calendar: GoogleAppsScript.Calendar.Calendar,
    event: GoogleAppsScript.Calendar.CalendarEvent | null,
  ): void => {
    if (!event) return;
    // CalendarApp can surface multiple object references for the same iCalUID
    // (for example through default and explicitly addressed calendars).
    found.set(event.getId(), event);
  };

  const trackedEventIds = getCcscTrackedEventIds();
  calendars.forEach((calendar) => {
    trackedEventIds.forEach((eventId) => {
      addEvent(calendar, calendar.getEventById(eventId));
    });
  });

  messages.forEach((message) => {
    const reservation = parseCcscReservation(message.body);
    const oneMinute = 60 * 1000;
    calendars.forEach((calendar) => {
      addEvent(calendar, findExistingCcscEvent(calendar, reservation));
      calendar
        .getEvents(
          new Date(reservation.start.getTime() - oneMinute),
          new Date(reservation.end.getTime() + oneMinute),
        )
        .filter((event) => {
          const loanTag = event.getTag(CCSC_CONFIG.eventLoanIdTag);
          if (loanTag === reservation.loanId) return true;
          const managedTitle =
            event.getTitle().startsWith(CCSC_CONFIG.eventTitlePrefix) ||
            event
              .getTitle()
              .startsWith(CCSC_CONFIG.legacyProvisionalTitlePrefix);
          return (
            managedTitle &&
            event
              .getDescription()
              .includes(`貸出ID: ${reservation.loanId}`)
          );
        })
        .forEach((event) => addEvent(calendar, event));
    });
  });

  return [...found.values()];
}

function deleteCcscEventsForReset(
  events: GoogleAppsScript.Calendar.CalendarEvent[],
): number {
  let deleted = 0;
  events.forEach((event) => {
    try {
      event.deleteEvent();
      deleted += 1;
    } catch (error) {
      const message = String(error);
      if (/does not exist|already been deleted/i.test(message)) {
        console.warn(
          `CCSC reset skipped an event that was already absent (${event.getId()}).`,
        );
        return;
      }
      throw error;
    }
  });
  return deleted;
}

function clearCcscScriptState(): number {
  const properties = PropertiesService.getScriptProperties();
  const keys = Object.keys(properties.getProperties()).filter(
    (key) =>
      key.startsWith(CCSC_CONFIG.processedPropertyPrefix) ||
      key === CCSC_CONFIG.emailLinkMigrationProperty,
  );
  keys.forEach((key) => properties.deleteProperty(key));
  return keys.length;
}

function verifyCcscConfirmedEvents(): CcscEventVerification[] {
  const messages = loadCcscMessages(listRecentCcscMessageReferences());
  const reservationsByLoan = new Map<string, CcscReservation>();
  messages.forEach((message) => {
    const reservation = parseCcscReservation(message.body);
    reservationsByLoan.set(reservation.loanId, reservation);
  });

  return [...reservationsByLoan.values()].map((reservation) => {
    const record = readCcscProcessedRecord(
      ccscLoanPropertyKey(reservation.loanId),
    );
    let apiEvent: GoogleAppsScript.Calendar.Schema.Event | null = null;
    if (record?.apiEventId) {
      try {
        apiEvent = requireCcscCalendarApi().Events.get(
          getCcscCalendarApiId(),
          record.apiEventId,
        );
      } catch (_error) {
        apiEvent = null;
      }
    }
    return {
      loanId: reservation.loanId,
      found: Boolean(apiEvent?.id && apiEvent.status !== 'cancelled'),
      apiEventId: String(apiEvent?.id || ''),
      title: String(apiEvent?.summary || ''),
      start: String(apiEvent?.start?.dateTime || apiEvent?.start?.date || ''),
    };
  });
}

/**
 * Destructive only within this app's scope: removes tracked CCSC events and
 * CCSC script state, recreates confirmed reservations on the pinned university
 * calendar, verifies the API records, and reinstalls exactly one trigger.
 */
function resetCcscAutomation(): CcscResetSummary {
  const actualEmail = requireExpectedCcscOwner();
  removeCcscAutomation();

  let deletedEvents = 0;
  let deletedProperties = 0;
  let processing: CcscProcessingSummary;
  try {
    const resetLock = LockService.getScriptLock();
    if (!resetLock.tryLock(5000)) {
      throw new Error('Another CCSC processing run is active. Try again shortly.');
    }
    try {
      const messages = loadCcscMessages(listRecentCcscMessageReferences());
      const events = collectCcscEventsForReset(messages);
      deletedEvents = deleteCcscEventsForReset(events);
      deletedProperties = clearCcscScriptState();
    } finally {
      resetLock.releaseLock();
    }

    processing = processCcscReservationEmails();
  } finally {
    // Even if the immediate import fails, leave one retry trigger installed.
    removeCcscAutomation();
    installCcscAutomationTrigger();
  }

  const verifiedEvents = verifyCcscConfirmedEvents();
  const missing = verifiedEvents.filter((event) => !event.found);
  if (missing.length) {
    throw new Error(
      `Reset created unverifiable events for loan ID(s): ${missing
        .map((event) => event.loanId)
        .join(', ')}`,
    );
  }
  const triggerCount = ScriptApp.getProjectTriggers().filter(
    (trigger) =>
      trigger.getHandlerFunction() === 'processCcscReservationEmails',
  ).length;
  const summary: CcscResetSummary = {
    calendarId: getCcscCalendarApiId(),
    deletedEvents,
    deletedProperties,
    processing,
    verifiedEvents,
    triggerCount,
  };
  console.log(
    `CCSC hard reset completed for ${actualEmail}: ${JSON.stringify(summary)}`,
  );
  return summary;
}

/**
 * Run once, manually, while signed in as the configured owner.
 * It validates the account, replaces only this app's old triggers, installs
 * the one-minute trigger, and immediately imports matching recent messages.
 */
function setupCcscAutomation(): CcscProcessingSummary {
  const actualEmail = requireExpectedCcscOwner();

  removeCcscAutomation();
  let summary: CcscProcessingSummary;
  try {
    // Setup is also the recovery path: do not let a stale processed-message
    // property suppress an event that is no longer present in Calendar.
    summary = reconcileCcscReservationEmails();
  } finally {
    installCcscAutomationTrigger();
  }
  console.log(
    `CCSC automation installed for ${actualEmail}; guest=${CCSC_CONFIG.guestEmail}`,
  );
  return summary;
}

/** Removes this app's time trigger. Existing Calendar events are untouched. */
function removeCcscAutomation(): number {
  const triggers = ScriptApp.getProjectTriggers().filter(
    (trigger) =>
      trigger.getHandlerFunction() === 'processCcscReservationEmails',
  );
  triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  console.log(`Removed ${triggers.length} CCSC trigger(s).`);
  return triggers.length;
}

/**
 * Read-only preview for checking the current Gmail match and parser output.
 * It does not create Calendar events or mark messages as processed.
 */
function previewRecentCcscReservations(): object[] {
  const results = loadCcscMessages(
    listRecentCcscMessageReferences(),
  ).map((message) => {
    try {
      const reservation = parseCcscReservation(message.body);
      return {
        messageId: message.id,
        receivedAt: message.receivedAt.toISOString(),
        loanId: reservation.loanId,
        title: buildCcscEventTitle(reservation),
        location: reservation.location,
        start: reservation.start.toISOString(),
        end: reservation.end.toISOString(),
        parseStatus: 'OK',
      };
    } catch (error) {
      return {
        messageId: message.id,
        receivedAt: message.receivedAt.toISOString(),
        parseStatus: 'ERROR',
        error: String(error),
      };
    }
  });
  console.log(JSON.stringify(results, null, 2));
  return results;
}

/** Reports whether the one-minute trigger is installed. */
function getCcscAutomationStatus(): object {
  const triggerCount = ScriptApp.getProjectTriggers().filter(
    (trigger) =>
      trigger.getHandlerFunction() === 'processCcscReservationEmails',
  ).length;
  const actualOwner = String(
    requireCcscGmailService().Users.getProfile('me').emailAddress || '',
  );
  const targetCalendar = getCcscCalendar();
  const status = {
    triggerInstalled: triggerCount === 1,
    triggerCount,
    expectedOwner: CCSC_CONFIG.expectedOwnerEmail,
    actualOwner,
    guest: CCSC_CONFIG.guestEmail,
    calendarSetting: CCSC_CONFIG.calendarId || 'default',
    targetCalendarId: targetCalendar.getId(),
    targetCalendarName: targetCalendar.getName(),
    pollingMinutes: CCSC_CONFIG.triggerEveryMinutes,
  };
  console.log(JSON.stringify(status, null, 2));
  return status;
}
