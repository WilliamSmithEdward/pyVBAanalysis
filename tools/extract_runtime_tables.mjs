// Mechanically extracts the VBA runtime constant + object tables to vendored JSON.
// Faithful (no hand-transcription), re-run after an XLIDE refresh:
//   npx -y tsx tools/extract_runtime_tables.mjs
import { writeFileSync } from 'node:fs';
import { VBA_RUNTIME_CONSTANTS, VBA_RUNTIME_OBJECTS } from '../../xlide_vscode/src/analyzer/runtime/vbaRuntime.ts';
const out = { constants: VBA_RUNTIME_CONSTANTS, objects: VBA_RUNTIME_OBJECTS };
writeFileSync('pyvbaanalysis/data/vba_runtime_tables.json', JSON.stringify(out));
console.log('constants:', VBA_RUNTIME_CONSTANTS.length, '| objects:', VBA_RUNTIME_OBJECTS.length,
  '| sample const:', JSON.stringify(VBA_RUNTIME_CONSTANTS[0]),
  '| sample object:', JSON.stringify(VBA_RUNTIME_OBJECTS[0]).slice(0,120));
