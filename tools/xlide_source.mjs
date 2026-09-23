// Where the extractors read XLIDE from.
//
// THE SIBLING CHECKOUT BY DEFAULT, A PIN WHEN ONE IS ASKED FOR. Day to day nobody
// sets XLIDE_ROOT and this reads the checkout next door. A sync sets it to a pin
// (tools/pin_analyzer.py), because that checkout is a working tree somebody is
// usually working in and vendored data must not depend on what is half-written in
// it this minute.
//
// THIS IS WHAT MAKES THE PIN REAL. The extractors used static `../../xlide_vscode/...`
// imports, so an environment variable could only ever have changed a check while
// esbuild-style resolution went on reading the tree next door. xlide_vbide shipped a
// release believing it was pinned and was not, for exactly that reason (its
// engine/build.mjs records it). Every read goes through importXlide below, resolved
// against the root, so pointing the root elsewhere moves the read with it.

import { existsSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));

/** The XLIDE checkout root: a pin when XLIDE_ROOT is set, the sibling otherwise. */
export const XLIDE_ROOT = process.env.XLIDE_ROOT
    ? resolve(process.env.XLIDE_ROOT)
    : resolve(here, '..', '..', 'xlide_vscode');

/** Proves the root is the XLIDE repository rather than some other folder. */
const MARKER = 'src/analyzer/diagnostics/ruleMetadata.ts';

if (!existsSync(resolve(XLIDE_ROOT, MARKER))) {
    console.error(`No XLIDE checkout at ${XLIDE_ROOT} (looked for ${MARKER}).`);
    console.error('Clone xlide_vscode next to this repository, or set XLIDE_ROOT to a pin:');
    console.error('  python tools/pin_analyzer.py --ref v10.5.0');
    process.exit(1);
}

/**
 * Imports one module out of the XLIDE checkout, by path relative to its root.
 *
 * A file: URL rather than a bare path, because a Windows drive letter in a bare
 * specifier is read as a protocol and the import fails.
 */
export async function importXlide(relativePath) {
    const full = resolve(XLIDE_ROOT, relativePath);
    if (!existsSync(full)) {
        // Named, not swallowed: a root missing a file the extractor needs is a
        // broken pin, and falling back to the sibling would hide what the pin is for.
        console.error(`The XLIDE root at ${XLIDE_ROOT} has no ${relativePath}.`);
        process.exit(1);
    }
    return import(pathToFileURL(full).href);
}
