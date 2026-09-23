// The upstream half of `harness.py projects`: labelled multi-module cases (see
// cases.py) through the upstream analyzer with project context, the case's host
// and its reference list, one JSON line per module. probes.py makes the port's half.
//
//     XLIDE_ROOT=<pin> npx -y tsx tools/differential/upstream/project_probe.mjs <cases.json>
import { readFileSync } from 'node:fs';
import { importXlide } from '../../xlide_source.mjs';

const { analyzeModule } = await importXlide('src/analyzer/index.ts');
const {
  buildVbaProjectIndex,
  effectiveModuleKind,
  projectAnalysisOptionsForModule,
  projectProcedureSignatures,
} = await importXlide('src/vbaProjectAnalysis.ts');

const lines = [];
for (const c of JSON.parse(readFileSync(process.argv[2], 'utf8'))) {
  const modules = c.modules.map((m) => ({
    moduleName: m.name,
    type: m.type ?? 'standard',
    source: m.source,
    ...(m.designerClass ? { designerClass: m.designerClass } : {}),
    ...(m.implicitMembers ? { implicitMembers: m.implicitMembers } : {}),
    ...(m.predeclaredId !== undefined ? { predeclaredId: m.predeclaredId } : {}),
  }));
  const project = buildVbaProjectIndex(modules);
  const procedures = projectProcedureSignatures(project);
  for (const mod of modules) {
    const found = analyzeModule(mod.source, {
      moduleName: mod.moduleName,
      moduleKind: effectiveModuleKind(mod),
      ...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
      ...(c.host ? { host: c.host } : {}),
      // Absent means "nothing referenced" upstream; the port is given the same.
      referencedHosts: c.referenced ?? [],
    });
    lines.push(JSON.stringify({
      label: c.label,
      module: mod.moduleName,
      findings: found.map((d) => `${d.code} @${d.span.start}-${d.span.end}: ${d.message}`),
    }));
  }
}
process.stdout.write(lines.join('\n') + '\n');
