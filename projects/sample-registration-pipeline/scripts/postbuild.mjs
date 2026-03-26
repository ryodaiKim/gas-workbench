import { mkdirSync, writeFileSync } from 'fs';
import { dirname } from 'path';

const manifest = {
  timeZone: 'Asia/Tokyo',
  dependencies: {
    enabledAdvancedServices: [
      {
        userSymbol: 'Drive',
        serviceId: 'drive',
        version: 'v2',
      },
    ],
  },
  exceptionLogging: 'STACKDRIVER',
  runtimeVersion: 'V8',
  oauthScopes: [
    'https://www.googleapis.com/auth/script.external_request',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents.readonly',
    'https://www.googleapis.com/auth/script.scriptapp',
  ],
};

const out = new URL('../build/appsscript.json', import.meta.url).pathname;
mkdirSync(dirname(out), { recursive: true });
writeFileSync(out, JSON.stringify(manifest, null, 2));
console.log('Wrote build/appsscript.json');
