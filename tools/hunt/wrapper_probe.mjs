// Everything XLIDE's analyzeVbaModuleSource shows for each module of each
// case, its structural pass included, with project context.
//
//     XLIDE_ROOT=<tree> npx -y tsx wrapper_probe.mjs CASES.json [LABEL...]
import { readFileSync } from 'node:fs';

const { importXlide } = await import(new URL('../xlide_source.mjs', import.meta.url).href);
const { analyzeVbaModuleSource } = await importXlide('src/vbaModuleAnalysis.ts');
const { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } =
  await importXlide('src/vbaProjectAnalysis.ts');

const wanted = new Set(process.argv.slice(3));
for (const c of JSON.parse(readFileSync(process.argv[2], 'utf8'))) {
  if (wanted.size && !wanted.has(c.label)) continue;
  const modules = c.modules.map((m) => ({ moduleName: m.name, type: m.type ?? 'standard', source: m.source }));
  const project = buildVbaProjectIndex(modules);
  const procedures = projectProcedureSignatures(project);
  for (const mod of modules) {
    const found = analyzeVbaModuleSource({
      source: mod.source,
      moduleType: mod.type,
      moduleName: mod.moduleName,
      moduleKind: effectiveModuleKind(mod),
      ...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
      host: c.host ?? 'excel',
      referencedHosts: [],
    }).diagnostics;
    console.log(JSON.stringify({ label: c.label, module: mod.moduleName, codes: found.map((d) => `${d.code}: ${d.message.slice(0, 70)}`) }));
  }
}
