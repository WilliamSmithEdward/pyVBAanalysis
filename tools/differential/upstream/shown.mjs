// What XLIDE shows for one module, cut to what the port reproduces. The corpus
// and projects probes both compare this.
//
// Every XLIDE surface runs the rules through analyzeVbaModuleSource
// (src/vbaModuleAnalysis.ts), which drops a runtime-error finding under
// On Error Resume Next and merges findings with one code and span; the port does
// both (pyvbaanalysis/diagnostics/module_analysis.py). The wrapper also adds
// findings of its own the port leaves out: its structural block-balance pass, the
// @xlide-test directive check and malformed @xlide-analysis directives. So its
// list is cut to the codes and spans the rules produced, each keeping the
// wrapper's verdict. The one wrapper step that removes rule findings and the port
// lacks is the expected-error suppression of @xlide-test procedures.
import { importXlide } from '../../xlide_source.mjs';

const { analyzeModule } = await importXlide('src/analyzer/index.ts');
const { analyzeVbaModuleSource } = await importXlide('src/vbaModuleAnalysis.ts');

const key = (d) => `${d.code}:${d.span.start}:${d.span.end}`;

/** A finding in place of what an upstream call threw (XLIDE issue #148). */
export function thrown(error) {
  return { code: '(throw)', span: { start: 0, end: 0 }, message: String(error?.message ?? error) };
}

export function shownFindings(source, moduleType, options) {
  const fromRules = new Set(analyzeModule(source, options).map(key));
  try {
    return analyzeVbaModuleSource({ source, moduleType, ...options })
      .diagnostics.filter((d) => fromRules.has(key(d)));
  } catch (error) {
    // The wrapper parses outside its own guards, so an input that crashes the
    // parser throws here where analyzeModule returned nothing.
    return [thrown(error)];
  }
}
