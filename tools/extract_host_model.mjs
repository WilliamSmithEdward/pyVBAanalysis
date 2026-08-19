// Mechanically extracts the XLIDE host object models to vendored JSON.
// Faithful (no hand-transcription): imports the generated host modules from the
// sibling xlide_vscode checkout and serializes each get*ObjectModel() verbatim.
// Re-run after an XLIDE host-data refresh:  npx -y tsx tools/extract_host_model.mjs
import { writeFileSync } from 'node:fs';
import { getExcelObjectModel } from '../../xlide_vscode/src/analyzer/host/excelObjectModel.ts';
import { getWordObjectModel } from '../../xlide_vscode/src/analyzer/host/wordObjectModel.ts';
import { getPowerPointObjectModel } from '../../xlide_vscode/src/analyzer/host/powerpointObjectModel.ts';
import { getAccessObjectModel } from '../../xlide_vscode/src/analyzer/host/accessObjectModel.ts';

// One entry per host the analyzer can select through the host registry. The
// file names are the ones pyvbaanalysis/host/host_model.py loads by host token.
const HOSTS = [
  ['excel', getExcelObjectModel, 'pyvbaanalysis/data/excel_host_model.json'],
  ['word', getWordObjectModel, 'pyvbaanalysis/data/word_host_model.json'],
  ['powerpoint', getPowerPointObjectModel, 'pyvbaanalysis/data/powerpoint_host_model.json'],
  ['access', getAccessObjectModel, 'pyvbaanalysis/data/access_host_model.json'],
];

for (const [token, getModel, out] of HOSTS) {
  const model = getModel();
  const json = JSON.stringify(model);
  writeFileSync(out, json);
  const count = (value) => (value ? Object.keys(value).length : 0);
  console.log(
    'wrote', out, '-', json.length, 'bytes;',
    count(model.types), 'types,', count(model.constants), 'constants,',
    count(model.globals), 'globals,', count(model.aliases), 'aliases',
    `(${token})`,
  );
}
