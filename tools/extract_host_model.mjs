// Mechanically extracts the XLIDE Excel host object model to vendored JSON.
// Faithful (no hand-transcription): imports the generated host module from the
// sibling xlide_vscode checkout and serializes getExcelObjectModel() verbatim.
// Re-run after an XLIDE host-data refresh:  npx -y tsx tools/extract_host_model.mjs
import { writeFileSync } from 'node:fs';
import { getExcelObjectModel } from '../../xlide_vscode/src/analyzer/host/excelObjectModel.ts';
const model = getExcelObjectModel();
const out = 'pyvbaanalysis/data/excel_host_model.json';
writeFileSync(out, JSON.stringify(model));
console.log('wrote', out, '-', JSON.stringify(model).length, 'bytes;',
  Object.keys(model.types).length, 'types,', model.constants.length, 'constants,',
  model.globals.length, 'globals,', Object.keys(model.aliases).length, 'aliases');
