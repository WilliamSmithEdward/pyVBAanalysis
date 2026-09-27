// The upstream half of `harness.py corpus`: every oracle case through the
// upstream analyzer, each module standalone ("S") and with the case's modules as
// a project ("P"), one JSON line per diagnostic. probes.py makes the port's half.
// Each module's findings are what XLIDE shows for it, cut to what the port
// reproduces (shown.mjs).
//
//     XLIDE_ROOT=<pin> npx -y tsx tools/differential/upstream/corpus_messages.mjs <cases.json>
//
// The project context is wired the way upstream's own no-false-positive sweep
// wires it (tests/oracleAcceptedNoFalsePositives.test.ts).
import { readFileSync } from 'node:fs';
import { importXlide } from '../../xlide_source.mjs';
import { shownFindings, thrown } from './shown.mjs';

const {
  buildVbaProjectIndex,
  effectiveModuleKind,
  projectAnalysisOptionsForModule,
  projectProcedureSignatures,
} = await importXlide('src/vbaProjectAnalysis.ts');

const corpus = JSON.parse(readFileSync(process.argv[2], 'utf8'));
const lines = [];
const emit = (pass, id, module, d) => lines.push(JSON.stringify({
  pass, id, module, code: d.code, start: d.span.start, end: d.span.end, message: d.message,
}));

for (const c of corpus.cases) {
  // The corpus keeps a module's kind under `type`, not `moduleType`.
  const raw = c.modules ?? [{ name: c.moduleName ?? 'Module1', type: 'standard', source: c.source }];
  const modules = raw.map((m) => ({ moduleName: m.name, type: m.type ?? 'standard', source: m.source }));
  for (const mod of modules) {
    const alone = shownFindings(mod.source, mod.type, { moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod) });
    for (const d of alone) emit('S', c.id, mod.moduleName, d);
  }
  let project;
  let procedures;
  try {
    project = buildVbaProjectIndex(modules);
    procedures = projectProcedureSignatures(project);
  } catch (error) {
    for (const mod of modules) emit('P', c.id, mod.moduleName, thrown(error));
    continue;
  }
  for (const mod of modules) {
    const found = shownFindings(mod.source, mod.type, {
      moduleName: mod.moduleName,
      moduleKind: effectiveModuleKind(mod),
      ...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
    });
    for (const d of found) emit('P', c.id, mod.moduleName, d);
  }
}
process.stdout.write(lines.join('\n') + '\n');
