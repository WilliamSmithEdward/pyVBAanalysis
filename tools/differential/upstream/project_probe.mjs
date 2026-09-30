// The upstream half of `harness.py projects`: labelled multi-module cases (see
// cases.py) through the upstream analyzer with project context, the case's host
// and its reference list, one JSON line per module. probes.py makes the port's half.
// Each module's findings are what XLIDE shows for it, cut to what the port
// reproduces (shown.mjs).
//
//     XLIDE_ROOT=<pin> npx -y tsx tools/differential/upstream/project_probe.mjs <cases.json>
import { readFileSync } from 'node:fs';
import { importXlide } from '../../xlide_source.mjs';
import { shownFindings, thrown } from './shown.mjs';

const {
  buildVbaProjectIndex,
  effectiveModuleKind,
  projectAnalysisOptionsForModule,
  projectProcedureSignatures,
} = await importXlide('src/vbaProjectAnalysis.ts');

const lines = [];
const line = (label, module, found) => lines.push(JSON.stringify({
  label,
  module,
  findings: found.map((d) => `${d.code} @${d.span.start}-${d.span.end}: ${d.message}`),
}));

for (const c of JSON.parse(readFileSync(process.argv[2], 'utf8'))) {
  const modules = c.modules.map((m) => ({
    moduleName: m.name,
    type: m.type ?? 'standard',
    source: m.source,
    ...(m.designerClass ? { designerClass: m.designerClass } : {}),
    ...(m.implicitMembers ? { implicitMembers: m.implicitMembers } : {}),
    ...(m.predeclaredId !== undefined ? { predeclaredId: m.predeclaredId } : {}),
  }));
  let project;
  let procedures;
  try {
    project = buildVbaProjectIndex(modules);
    procedures = projectProcedureSignatures(project);
  } catch (error) {
    for (const mod of modules) line(c.label, mod.moduleName, [thrown(error)]);
    continue;
  }
  for (const mod of modules) {
    line(c.label, mod.moduleName, shownFindings(mod.source, mod.type, {
      moduleName: mod.moduleName,
      moduleKind: effectiveModuleKind(mod),
      ...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
      // The extension's analysis worker passes the designer's class beside the
      // project options (analysisWorkerLogic.ts); projectAnalysisOptionsForModule
      // does not carry it, and the port's analyze_project does.
      ...(mod.designerClass ? { designerClass: mod.designerClass } : {}),
      ...(c.host ? { host: c.host } : {}),
      // Absent means "nothing referenced" upstream; the port is given the same.
      referencedHosts: c.referenced ?? [],
    }));
  }
}
process.stdout.write(lines.join('\n') + '\n');
