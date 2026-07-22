type CcscReservation = {
  loanId: string;
  representativeName: string;
  contact: string;
  affiliation: string;
  location: string;
  purpose: string;
  title: string;
  attendeeCount: number | null;
  attendeeBreakdown: string;
  notes: string;
  start: Date;
  end: Date;
};

type CcscMailMessage = {
  id: string;
  threadId: string;
  from: string;
  to: string;
  subject: string;
  body: string;
  receivedAt: Date;
};

type CcscProcessingSummary = {
  matched: number;
  created: number;
  updated: number;
  duplicates: number;
  alreadyProcessed: number;
  failed: number;
};

type CcscProcessedRecord = {
  processedAt: string;
  loanId: string;
  eventId: string;
  apiEventId?: string;
  duplicate: boolean;
};

type CcscCreatedEvent = {
  eventId: string;
  apiEventId: string;
};

type CcscEventVerification = {
  loanId: string;
  found: boolean;
  apiEventId: string;
  title: string;
  start: string;
};

type CcscResetSummary = {
  calendarId: string;
  deletedEvents: number;
  deletedProperties: number;
  processing: CcscProcessingSummary;
  verifiedEvents: CcscEventVerification[];
  triggerCount: number;
};

type CcscEmailLinkBackfillSummary = {
  matched: number;
  updated: number;
  alreadyLinked: number;
  missingEvents: number;
  failed: number;
};
