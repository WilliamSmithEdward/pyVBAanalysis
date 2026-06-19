// Mechanically extracts the Excel event-handler catalogue to vendored JSON.
// Faithful (no hand-transcription): imports ALL_EVENT_DEFINITIONS from the sibling
// xlide_vscode checkout and serializes name + owner + derived document type for
// each event. Re-run after an XLIDE event-catalogue refresh:
//   npx -y tsx tools/extract_event_definitions.mjs
import { writeFileSync } from 'node:fs';
import {
  ALL_EVENT_DEFINITIONS,
  eventHandlerProcedureForName,
} from '../../xlide_vscode/src/analyzer/completion/eventHandlers.ts';

// documentTypeForOwner is internal; eventHandlerProcedureForName exposes the same
// owner -> documentType mapping per definition, so reuse it (stays faithful).
const events = ALL_EVENT_DEFINITIONS.map((def) => {
  const match = eventHandlerProcedureForName(def.name);
  if (!match) {
    throw new Error(`event '${def.name}' did not resolve a procedure match`);
  }
  return { name: match.name, owner: match.owner, documentType: match.documentType };
});

const out = 'pyvbaanalysis/data/event_definitions.json';
writeFileSync(out, JSON.stringify({ events }));
console.log('wrote', out, '-', events.length, 'events;',
  'sample:', JSON.stringify(events[0]));
