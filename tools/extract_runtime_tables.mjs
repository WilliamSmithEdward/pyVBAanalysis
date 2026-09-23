// Mechanically extracts the VBA runtime function, constant and object tables to
// vendored JSON. Faithful (no hand-transcription), re-run after an XLIDE refresh:
//   npx -y tsx tools/extract_runtime_tables.mjs
// Set XLIDE_ROOT to vendor from a pin instead of the sibling checkout.
import { writeFileSync } from 'node:fs';
import { importXlide } from './xlide_source.mjs';

const { VBA_RUNTIME_FUNCTIONS, VBA_RUNTIME_CONSTANTS, VBA_RUNTIME_OBJECTS } = await importXlide(
  'src/analyzer/runtime/vbaRuntime.ts',
);
const out = {
  constants: VBA_RUNTIME_CONSTANTS,
  objects: VBA_RUNTIME_OBJECTS,
  functions: VBA_RUNTIME_FUNCTIONS,
};
writeFileSync('pyvbaanalysis/data/vba_runtime_tables.json', JSON.stringify(out));
console.log('functions:', VBA_RUNTIME_FUNCTIONS.length, '| constants:', VBA_RUNTIME_CONSTANTS.length,
  '| objects:', VBA_RUNTIME_OBJECTS.length,
  '| sample function:', JSON.stringify(VBA_RUNTIME_FUNCTIONS[0]).slice(0, 120),
  '| sample object:', JSON.stringify(VBA_RUNTIME_OBJECTS[0]).slice(0, 120));
