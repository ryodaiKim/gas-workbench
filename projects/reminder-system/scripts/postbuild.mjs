// Copy appsscript.json into build/
import { mkdirSync, writeFileSync } from 'fs';
import { dirname } from 'path';

const manifest = {
  timeZone: 'Asia/Tokyo',
  dependencies: {},
  exceptionLogging: 'STACKDRIVER',
  runtimeVersion: 'V8',
};

const out = new URL('../build/appsscript.json', import.meta.url).pathname;
mkdirSync(dirname(out), { recursive: true });
writeFileSync(out, JSON.stringify(manifest, null, 2));
console.log('Wrote build/appsscript.json');
