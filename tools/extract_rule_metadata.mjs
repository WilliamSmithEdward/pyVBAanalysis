// Mechanically extracts the XLIDE diagnostic rule catalogue to vendored JSON.
// Faithful (no hand-transcription): imports DIAGNOSTIC_RULES from the sibling
// xlide_vscode checkout and serializes it verbatim (pretty-printed, LF, ASCII).
// Re-run after an XLIDE rule-metadata refresh:  npx -y tsx tools/extract_rule_metadata.mjs
import { writeFileSync } from 'node:fs';
import { DIAGNOSTIC_RULES } from '../../xlide_vscode/src/analyzer/diagnostics/ruleMetadata.ts';
const out = 'pyvbaanalysis/data/rule_metadata.json';
const json = JSON.stringify(DIAGNOSTIC_RULES, null, 2) + '\n';
writeFileSync(out, json);
console.log('wrote', out, '-', json.length, 'bytes;',
  Object.keys(DIAGNOSTIC_RULES).length, 'rules');
