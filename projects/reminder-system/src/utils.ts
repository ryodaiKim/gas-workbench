// Utility helpers for Apps Script

function toDateOnly(d: Date | string | null | undefined): string | null {
  if (!d) return null;
  const dt = typeof d === 'string' ? new Date(d) : d;
  if (isNaN(dt.getTime())) return null;
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function todayStr(): string {
  const now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function nowStr(): string {
  const now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

function renderTemplate(tpl: string, ctx: Record<string, string>): string {
  return tpl.replace(/\{\{\s*([^}]+)\s*\}\}/g, (_m, key) => {
    const k = String(key).trim();
    return k in ctx ? String(ctx[k]) : '';
  });
}

function splitEmails(s?: string): string[] {
  if (!s) return [];
  return s
    .split(/[;,\n]/)
    .map((x) => x.trim())
    .filter(Boolean);
}
