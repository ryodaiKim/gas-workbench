import { readFileSync } from 'fs';
import vm from 'vm';

const scriptProperties = {};
const events = [];
const calendarPatchCalls = [];
let messageIds = ['message-1', 'message-2', 'message-3'];
let nextEventId = 1;
const projectTriggers = [];

const firstSampleBody = [
  '以下の内容で、予約を受け付けました。',
  '',
  '貸出ID      65083',
  '代表者氏名  テスト 太郎',
  '連絡先      08000000000',
  '代表者所属  大学医学部 学生（所属先：なし）',
  '使用場所    診察シミュレーション室（5）',
  '使用目的    個人トレーニング',
  '件　名      自主練',
  '使用人数    2人 （学生 2人）',
  '備　考      救急蘇生法',
  '使用日時    2026/08/04 10:00 から 2026/08/04 10:30 まで',
].join('\r\n');
const sampleBodies = {
  'message-1': firstSampleBody,
  'message-2': firstSampleBody
    .replace('65083', '65084')
    .replace('診察シミュレーション室（5）', '診察シミュレーション室（7）')
    .replace('救急蘇生法', '十二誘導心電図')
    .replace(
      '2026/08/04 10:00 から 2026/08/04 10:30',
      '2026/08/04 10:30 から 2026/08/04 11:00',
    ),
  'message-3': firstSampleBody
    .replace('65083', '65085')
    .replace('診察シミュレーション室（5）', '診察シミュレーション室（11）')
    .replace('救急蘇生法', '採血')
    .replace(
      '2026/08/04 10:00 から 2026/08/04 10:30',
      '2026/08/04 11:00 から 2026/08/04 11:30',
    ),
  // A resent copy of message-1: a different Gmail ID but the same 貸出ID.
  'message-4': firstSampleBody,
  // A confirmation used to upgrade a legacy provisional event in place.
  'message-5': firstSampleBody,
};

function createMockMessage(id) {
  const body = sampleBodies[id];
  if (!body) throw new Error(`Unknown mock message: ${id}`);
  return {
    id,
    threadId: 'thread-1',
    internalDate: '1784347200000',
    payload: {
      mimeType: 'text/plain',
      headers: [
        {
          name: 'From',
          value: 'byo-hp@system.ho.u-chiba.jp',
        },
        {
          name: 'To',
          value: '23mb1095@student.gs.chiba-u.jp',
        },
        {
          name: 'Subject',
          value: '【CCSC】予約申請 確認証および貸出証・利用記録',
        },
        {
          name: 'Content-Type',
          value: 'text/plain; charset="UTF-8"',
        },
      ],
      body: {
        data: Buffer.from(body, 'utf8').toString('base64url'),
      },
    },
  };
}

const calendar = {
  createEvent(title, start, end, options) {
    const tags = {};
    const event = {
      id: `event-${nextEventId}`,
      title,
      start,
      end,
      options,
      description: options.description,
      location: options.location,
      popupReminders: [],
      tags,
      setTag(key, value) {
        tags[key] = value;
        return this;
      },
      getTag(key) {
        return tags[key] || '';
      },
      getTitle() {
        return this.title;
      },
      getDescription() {
        return this.description;
      },
      getStartTime() {
        return this.start;
      },
      getEndTime() {
        return this.end;
      },
      getGuestList() {
        return String(this.options.guests || '')
          .split(',')
          .filter(Boolean)
          .map((email) => ({
            getEmail() {
              return email;
            },
          }));
      },
      addGuest(email) {
        this.options.guests = [this.options.guests, email]
          .filter(Boolean)
          .join(',');
        return this;
      },
      setTitle(value) {
        this.title = value;
        return this;
      },
      setDescription(value) {
        this.description = value;
        return this;
      },
      setLocation(value) {
        this.location = value;
        return this;
      },
      setTime(start, end) {
        this.start = start;
        this.end = end;
        return this;
      },
      deleteEvent() {
        const index = events.indexOf(this);
        if (index < 0) {
          throw new Error(
            'The calendar event does not exist, or it has already been deleted.',
          );
        }
        events.splice(index, 1);
      },
      deleteTag(key) {
        delete tags[key];
        return this;
      },
      addPopupReminder(minutes) {
        this.popupReminders.push(minutes);
        return this;
      },
      getId() {
        return this.id;
      },
    };
    nextEventId += 1;
    events.push(event);
    return event;
  },
  getEventById(id) {
    return events.find((event) => event.id === id) || null;
  },
  getEvents() {
    return events;
  },
  getId() {
    return '23mb1095@student.gs.chiba-u.jp';
  },
  getName() {
    return '23mb1095@student.gs.chiba-u.jp';
  },
};

function toMockApiEvent(event) {
  return {
    id: `api-${event.id}`,
    iCalUID: event.id,
    status: 'confirmed',
    summary: event.title,
    description: event.description,
    location: event.location,
    start: { dateTime: event.start.toISOString() },
    end: { dateTime: event.end.toISOString() },
    extendedProperties: { private: { ...event.tags } },
  };
}

function createMockTrigger(handlerFunction) {
  return {
    getHandlerFunction() {
      return handlerFunction;
    },
  };
}

projectTriggers.push(createMockTrigger('processCcscReservationEmails'));

const context = vm.createContext({
  console,
  ccscTestCalendar: calendar,
  ccscTestEvents: events,
  Gmail: {
    Users: {
      Messages: {
        list(_user, options) {
          if (
            !String(options.q || '').includes(
              '【CCSC】予約申請 確認証および貸出証・利用記録',
            )
          ) {
            throw new Error('Gmail query is not limited to confirmed mail.');
          }
          return {
            messages: messageIds.map((id) => ({ id })),
          };
        },
        get(_user, id) {
          return createMockMessage(id);
        },
      },
      getProfile() {
        return {
          emailAddress: '23mb1095@student.gs.chiba-u.jp',
        };
      },
    },
  },
  Utilities: {
    base64DecodeWebSafe(data) {
      if (data.length % 4 !== 0) {
        throw new Error('Mock decoder requires padded input.');
      }
      return Array.from(Buffer.from(data, 'base64url'));
    },
    base64Decode(data) {
      return Array.from(Buffer.from(data, 'base64'));
    },
    newBlob(bytes) {
      return {
        getDataAsString() {
          return Buffer.from(bytes).toString('utf8');
        },
      };
    },
    formatDate(date) {
      return date.toISOString();
    },
  },
  CalendarApp: {
    getDefaultCalendar() {
      return calendar;
    },
    getCalendarById() {
      return calendar;
    },
  },
  Calendar: {
    Events: {
      insert(resource, calendarId, options) {
        if (calendarId !== '23mb1095@student.gs.chiba-u.jp') {
          throw new Error(`Unexpected target calendar: ${calendarId}`);
        }
        const event = calendar.createEvent(
          resource.summary,
          new Date(resource.start.dateTime),
          new Date(resource.end.dateTime),
          {
            description: resource.description,
            location: resource.location,
            guests: (resource.attendees || [])
              .map((attendee) => attendee.email)
              .join(','),
            sendInvites: options.sendUpdates === 'all',
          },
        );
        Object.assign(
          event.tags,
          resource.extendedProperties?.private || {},
        );
        const popup = (resource.reminders?.overrides || []).find(
          (reminder) => reminder.method === 'popup',
        );
        if (popup) event.popupReminders.push(popup.minutes);
        return toMockApiEvent(event);
      },
      get(_calendarId, apiEventId) {
        const event = events.find(
          (candidate) => `api-${candidate.id}` === apiEventId,
        );
        if (!event) throw new Error(`Unknown API event: ${apiEventId}`);
        return toMockApiEvent(event);
      },
      list(_calendarId, options) {
        let matchingEvents = events;
        // Reproduce the cloud failure reported for the older provisional
        // events: direct iCalUID lookup returns nothing, so application tags or
        // a narrow time query must resolve the Calendar API event ID.
        if (options.iCalUID) matchingEvents = [];
        if (options.privateExtendedProperty) {
          const [key, value] = options.privateExtendedProperty.split('=');
          matchingEvents = events.filter(
            (candidate) => candidate.tags[key] === value,
          );
        }
        if (options.timeMin || options.timeMax) {
          const min = options.timeMin
            ? new Date(options.timeMin).getTime()
            : -Infinity;
          const max = options.timeMax
            ? new Date(options.timeMax).getTime()
            : Infinity;
          matchingEvents = events.filter(
            (candidate) =>
              candidate.end.getTime() > min &&
              candidate.start.getTime() < max,
          );
        }
        return {
          items: matchingEvents.map(toMockApiEvent),
        };
      },
      patch(resource, calendarId, apiEventId, options) {
        const event = events.find(
          (candidate) => `api-${candidate.id}` === apiEventId,
        );
        if (!event) throw new Error(`Unknown API event: ${apiEventId}`);
        if (resource.summary !== undefined) event.title = resource.summary;
        if (resource.description !== undefined) {
          event.description = resource.description;
        }
        if (resource.location !== undefined) {
          event.location = resource.location;
        }
        if (resource.start?.dateTime) {
          event.start = new Date(resource.start.dateTime);
        }
        if (resource.end?.dateTime) {
          event.end = new Date(resource.end.dateTime);
        }
        calendarPatchCalls.push({
          resource,
          calendarId,
          apiEventId,
          options,
        });
        return {
          id: apiEventId,
          description: event.description,
        };
      },
    },
  },
  PropertiesService: {
    getScriptProperties() {
      return {
        getProperty(key) {
          return scriptProperties[key] || null;
        },
        setProperties(values) {
          Object.assign(scriptProperties, values);
        },
        setProperty(key, value) {
          scriptProperties[key] = value;
        },
        getProperties() {
          return { ...scriptProperties };
        },
        deleteProperty(key) {
          delete scriptProperties[key];
        },
      };
    },
  },
  LockService: {
    getScriptLock() {
      return {
        tryLock() {
          return true;
        },
        releaseLock() {},
      };
    },
  },
  ScriptApp: {
    getProjectTriggers() {
      return [...projectTriggers];
    },
    deleteTrigger(trigger) {
      const index = projectTriggers.indexOf(trigger);
      if (index >= 0) projectTriggers.splice(index, 1);
    },
    newTrigger(handlerFunction) {
      return {
        timeBased() {
          return this;
        },
        everyMinutes() {
          return this;
        },
        create() {
          const trigger = createMockTrigger(handlerFunction);
          projectTriggers.push(trigger);
          return trigger;
        },
      };
    },
  },
});
const files = [
  'types.js',
  'config.js',
  'parser.js',
  'main.js',
  'tests.js',
];

for (const file of files) {
  const url = new URL(`../build/${file}`, import.meta.url);
  vm.runInContext(readFileSync(url, 'utf8'), context, { filename: file });
}

const result = vm.runInContext('runCcscSelfTests()', context);
if (result !== 'All CCSC parser tests passed.') {
  throw new Error(`Unexpected test result: ${String(result)}`);
}

const representationTest = vm.runInContext(
  `(() => {
    const text = '貸出ID 65083\\n使用日時 2026/08/04 10:00';
    const bytes = Array.from(text).flatMap((character) => {
      const encoded = unescape(encodeURIComponent(character));
      return Array.from(encoded).map((item) => item.charCodeAt(0));
    });
    const signedBytes = bytes.map((value) => value > 127 ? value - 256 : value);
    return {
      byteArray: decodeCcscBodyData(signedBytes, 'UTF-8') === text,
      numericText: decodeCcscBodyData(signedBytes.join(','), 'UTF-8') === text,
      plainText: decodeCcscBodyData(text, 'UTF-8') === text,
    };
  })()`,
  context,
);
if (
  !representationTest.byteArray ||
  !representationTest.numericText ||
  !representationTest.plainText
) {
  throw new Error(
    `MIME representation compatibility failed: ${JSON.stringify(
      representationTest,
    )}`,
  );
}

const firstRun = vm.runInContext(
  'processCcscReservationEmails()',
  context,
);
if (firstRun.created !== 3 || events.length !== 3) {
  throw new Error(
    `Expected three created events, received ${JSON.stringify(firstRun)}`,
  );
}
if (
  firstRun.updated !== 0 ||
  events.some((event) => !event.title.startsWith('【CCSC予約確定】')) ||
  events.some((event) => event.description.includes('仮受付'))
) {
  throw new Error('Confirmed events are not labeled as confirmed.');
}
if (events.some(
  (event) =>
    event.options.guests !== '22m1084c@student.gs.chiba-u.jp' ||
    event.options.sendInvites !== true,
)) {
  throw new Error('Calendar guest or sendInvites option is incorrect.');
}
if (events.some((event) => event.popupReminders[0] !== 30)) {
  throw new Error('Expected a 30-minute organizer popup reminder.');
}
if (
  events.some(
    (event) =>
      !event.description.includes(
        '元メールを開く: https://mail.google.com/mail/u/0/#all/thread-1',
      ),
  )
) {
  throw new Error('New event description is missing the Gmail link.');
}

events[0].description = events[0].description.replace(
  /\n元メールを開く: [^\n]+/,
  '',
);
const backfillResult = vm.runInContext(
  'backfillCcscEmailLinks()',
  context,
);
if (
  backfillResult.updated !== 1 ||
  backfillResult.alreadyLinked !== 2 ||
  calendarPatchCalls.length !== 1 ||
  calendarPatchCalls[0].options.sendUpdates !== 'none'
) {
  throw new Error(
    `Email-link backfill failed: ${JSON.stringify(backfillResult)}`,
  );
}

messageIds = ['message-1', 'message-2', 'message-3', 'message-4'];
const secondRun = vm.runInContext(
  'processCcscReservationEmails()',
  context,
);
if (
  secondRun.alreadyProcessed !== 3 ||
  secondRun.updated !== 0 ||
  secondRun.duplicates !== 1 ||
  events.length !== 3
) {
  throw new Error(
    `Duplicate suppression failed: ${JSON.stringify(secondRun)}`,
  );
}

events[0].title = '【CCSC仮予約】自主練｜救急蘇生法';
events[0].description = '仮受付の古い説明';
events[0].deleteTag('ccscReservationFingerprint');
const legacyLoanRecordKey = 'CCSC_PROCESSED_LOAN_65083';
const legacyLoanRecord = JSON.parse(scriptProperties[legacyLoanRecordKey]);
delete legacyLoanRecord.apiEventId;
scriptProperties[legacyLoanRecordKey] = JSON.stringify(legacyLoanRecord);
messageIds = [
  'message-1',
  'message-2',
  'message-3',
  'message-4',
  'message-5',
];
const upgradeRun = vm.runInContext(
  'processCcscReservationEmails()',
  context,
);
const confirmationPatch = calendarPatchCalls.find(
  (call) => call.options.sendUpdates === 'all',
);
if (
  upgradeRun.updated !== 1 ||
  upgradeRun.created !== 0 ||
  events.length !== 3 ||
  !events[0].title.startsWith('【CCSC予約確定】') ||
  events[0].description.includes('仮受付') ||
  !confirmationPatch
) {
  throw new Error(
    `Legacy confirmation upgrade failed: ${JSON.stringify(upgradeRun)}`,
  );
}

const missingEventIndex = events.findIndex(
  (event) => event.getTag('ccscLoanId') === '65084',
);
if (missingEventIndex < 0) {
  throw new Error('Could not prepare the stale-record recovery test.');
}
events.splice(missingEventIndex, 1);
const reconciliationRun = vm.runInContext(
  'reconcileCcscReservationEmails()',
  context,
);
if (
  reconciliationRun.created !== 1 ||
  reconciliationRun.failed !== 0 ||
  events.length !== 3 ||
  !events.some((event) => event.getTag('ccscLoanId') === '65084')
) {
  throw new Error(
    `Missing-event reconciliation failed: ${JSON.stringify(
      reconciliationRun,
    )}`,
  );
}

const status = vm.runInContext('getCcscAutomationStatus()', context);
if (
  status.actualOwner !== '23mb1095@student.gs.chiba-u.jp' ||
  status.targetCalendarId !== '23mb1095@student.gs.chiba-u.jp'
) {
  throw new Error(`Runtime account diagnostics failed: ${JSON.stringify(status)}`);
}

const resetResult = vm.runInContext('resetCcscAutomation()', context);
if (
  resetResult.deletedEvents !== 3 ||
  resetResult.deletedProperties < 1 ||
  resetResult.processing.created !== 3 ||
  resetResult.processing.failed !== 0 ||
  resetResult.verifiedEvents.length !== 3 ||
  resetResult.verifiedEvents.some((event) => !event.found) ||
  resetResult.calendarId !== '23mb1095@student.gs.chiba-u.jp' ||
  resetResult.triggerCount !== 1 ||
  events.length !== 3
) {
  throw new Error(`Hard reset failed: ${JSON.stringify(resetResult)}`);
}
for (const loanId of ['65083', '65084', '65085']) {
  const record = JSON.parse(
    scriptProperties[`CCSC_PROCESSED_LOAN_${loanId}`],
  );
  if (!record.apiEventId) {
    throw new Error(`Reset did not store an API event ID for ${loanId}.`);
  }
}

const idempotentDeletion = vm.runInContext(
  `(() => {
    const event = ccscTestCalendar.createEvent(
      'throwaway',
      new Date('2026-08-05T00:00:00.000Z'),
      new Date('2026-08-05T00:30:00.000Z'),
      { description: '', location: '', guests: '', sendInvites: false },
    );
    const deleted = deleteCcscEventsForReset([event, event]);
    return { deleted, remains: ccscTestEvents.includes(event) };
  })()`,
  context,
);
if (idempotentDeletion.deleted !== 1 || idempotentDeletion.remains) {
  throw new Error(
    `Idempotent reset deletion failed: ${JSON.stringify(idempotentDeletion)}`,
  );
}

console.log('All CCSC integration tests passed.');
