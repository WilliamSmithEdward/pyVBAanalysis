// Mechanically extracts the XLIDE diagnostic rule catalogue to vendored JSON.
// Faithful (no hand-transcription): imports DIAGNOSTIC_RULES from the XLIDE
// checkout and serializes it verbatim (pretty-printed, LF, ASCII).
// Re-run after an XLIDE rule-metadata refresh:  npx -y tsx tools/extract_rule_metadata.mjs
// Set XLIDE_ROOT to vendor from a pin instead of the sibling checkout.
import { writeFileSync } from 'node:fs';
import { importXlide } from './xlide_source.mjs';

const { DIAGNOSTIC_RULES } = await importXlide('src/analyzer/diagnostics/ruleMetadata.ts');
const out = 'pyvbaanalysis/data/rule_metadata.json';
const json = JSON.stringify(DIAGNOSTIC_RULES, null, 2) + '\n';
writeFileSync(out, json);
console.log('wrote', out, '-', json.length, 'bytes;',
  Object.keys(DIAGNOSTIC_RULES).length, 'rules');
