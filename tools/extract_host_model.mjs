// Mechanically extracts the XLIDE host object models to vendored JSON.
// Faithful (no hand-transcription): imports the generated host modules from the
// XLIDE checkout and serializes each get*ObjectModel() verbatim.
// Re-run after an XLIDE host-data refresh:  npx -y tsx tools/extract_host_model.mjs
// Set XLIDE_ROOT to vendor from a pin instead of the sibling checkout.
import { writeFileSync } from 'node:fs';
import { importXlide } from './xlide_source.mjs';

// One entry per host the analyzer can select through the host registry. The file
// names are the ones pyvbaanalysis/host/host_model.py loads by host token.
const HOSTS = [
  ['excel', 'excelObjectModel.ts', 'getExcelObjectModel', 'pyvbaanalysis/data/excel_host_model.json'],
  ['word', 'wordObjectModel.ts', 'getWordObjectModel', 'pyvbaanalysis/data/word_host_model.json'],
  ['powerpoint', 'powerpointObjectModel.ts', 'getPowerPointObjectModel', 'pyvbaanalysis/data/powerpoint_host_model.json'],
  ['access', 'accessObjectModel.ts', 'getAccessObjectModel', 'pyvbaanalysis/data/access_host_model.json'],
  // Not an Office host: a VB6 project's code-behind needs the VB runtime's surface
  // (App, Screen, Printer, the intrinsic controls), and the analyzer selects it by
  // the `vb6` token like any other.
  ['vb6', 'vb6ObjectModel.ts', 'getVb6ObjectModel', 'pyvbaanalysis/data/vb6_host_model.json'],
];

// Which Excel interfaces are NONEXTENSIBLE, measured upstream from the type
// library's TYPEFLAGS. A complete member list proves absence only for these: the
// rest of the object model is open, so VBA compiles the call and asks IDispatch
// for the name at run time. Extracted rather than transcribed, because the
// no-false-positive contract for member-not-found rests on the set being exact.
const { EXCEL_CLOSED_TYPE_NAMES } = await importXlide('src/analyzer/host/typeExtensibility.ts');
const closedOut = 'pyvbaanalysis/data/excel_closed_types.json';
writeFileSync(closedOut, JSON.stringify({ excelClosedTypes: [...EXCEL_CLOSED_TYPE_NAMES].sort() }, null, 2) + '\n');
console.log('wrote', closedOut, '-', EXCEL_CLOSED_TYPE_NAMES.length, 'closed Excel types');

// The Microsoft Forms members a UserForm and its controls carry: the generated
// type-library surface, and the members VBA wraps around a form (Show, Hide, Name,
// ...), which upstream verified on a live form. A form whose control list is
// authoritative proves a member absent (XLIDE issue #26), so a member missing here
// would be a false diagnostic on working code. Extracted, never transcribed.
const msforms = await importXlide('src/analyzer/host/msformsReferenceMembers.ts');
const userFormExtender = await importXlide('src/analyzer/host/userFormExtenderMembers.ts');
const msformsOut = 'pyvbaanalysis/data/msforms_members.json';
writeFileSync(msformsOut, JSON.stringify({
  userFormType: userFormExtender.VBA_USERFORM_TYPE,
  userFormExtenderMembers: userFormExtender.VBA_USERFORM_EXTENDER_MEMBERS,
  controlClassNames: msforms.MSFORMS_CONTROL_CLASS_NAMES,
  referenceMembers: msforms.MSFORMS_REFERENCE_MEMBERS,
}));
console.log(
  'wrote', msformsOut, '-',
  Object.keys(msforms.MSFORMS_REFERENCE_MEMBERS).length, 'MSForms types,',
  userFormExtender.VBA_USERFORM_EXTENDER_MEMBERS.length, 'UserForm extender members',
);

for (const [token, file, getter, out] of HOSTS) {
  const module = await importXlide(`src/analyzer/host/${file}`);
  const model = module[getter]();
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
