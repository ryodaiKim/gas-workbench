import { mkdirSync, writeFileSync } from 'fs';
import { dirname } from 'path';

const manifest = {
  timeZone: 'Asia/Tokyo',
  dependencies: {
    enabledAdvancedServices: [
      {
        userSymbol: 'Gmail',
        version: 'v1',
        serviceId: 'gmail',
      },
      {
        userSymbol: 'Calendar',
        version: 'v3',
        serviceId: 'calendar',
      },
    ],
  },
  oauthScopes: [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/script.scriptapp',
  ],
  exceptionLogging: 'STACKDRIVER',
  runtimeVersion: 'V8',
};

const out = new URL('../build/appsscript.json', import.meta.url).pathname;
mkdirSync(dirname(out), { recursive: true });
writeFileSync(out, JSON.stringify(manifest, null, 2));
console.log('Wrote build/appsscript.json');
