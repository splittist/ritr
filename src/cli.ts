import { readFile } from 'node:fs/promises';
import { basename } from 'node:path';
import { DocxPackage, comparePackages } from './package/docx';
import { readDocument, plainText } from './engine/document';
import { Workspace } from './engine/workspace';
import { exportWorkspace, saveAs } from './io/save';

const usage = `ritr — preservation-first DOCX workbench

  inspect <file.docx> [--json]            Stories, codes, and diagnostics
  compare <before.docx> <after.docx>      Per-part byte comparison (JSON)
  roundtrip <input.docx> <new.docx>       Validate and save without edits
  search <literal> <files...>            Source-bound matches (JSON)
  replace <literal> <replacement> <files...> [--out-dir <new-directory>]

Replace previews by default. --out-dir applies the preview and exports a new
directory. Search is case-sensitive and confined to individual text spans.
Originals are never overwritten. Use the desktop app for interactive editing.
`;
const print = (value: unknown) => console.log(JSON.stringify(value, null, 2));
async function main(args: string[]): Promise<void> {
  const [command, ...rest] = args;
  const load = async (file: string) => DocxPackage.open(await readFile(file));
  if (!command || command === '--help' || command === 'help') {
    console.log(usage);
    return;
  }
  if (command === 'inspect') {
    if (!rest[0] || rest.length > 2 || (rest[1] && rest[1] !== '--json')) throw new Error(usage);
    const pkg = await load(rest[0]);
    const model = readDocument(pkg);
    if (rest.includes('--json')) print({ parts: pkg.names(), ...model });
    else {
      console.log(`${rest[0]} · ${pkg.names().length} parts · ${model.stories.length} stories`);
      for (const story of model.stories) {
        console.log(`\n── ${story.label} (${story.part}) ──\n${plainText(story)}`);
        console.log(
          story.tokens
            .map((t) =>
              t.kind === 'text' ? t.span.text : `⟦${t.label}⟧${t.label === '¶' ? '\n' : ''}`,
            )
            .join(''),
        );
      }
      print(model.diagnostics);
    }
  } else if (command === 'compare' || command === 'roundtrip') {
    if (rest.length !== 2) throw new Error(usage);
    const before = await load(rest[0]!);
    if (command === 'roundtrip') await saveAs(rest[1]!, before);
    const report = comparePackages(before, await load(rest[1]!));
    print(report);
    if (command === 'compare' && report.some((p) => p.status !== 'unchanged')) process.exitCode = 2;
  } else if (command === 'search' || command === 'replace') {
    const query = rest.shift();
    const replacement = command === 'replace' ? rest.shift() : undefined;
    const outIndex = rest.indexOf('--out-dir');
    let output: string | undefined;
    if (outIndex >= 0) {
      if (command !== 'replace' || outIndex !== rest.length - 2) throw new Error(usage);
      output = rest[outIndex + 1];
      rest.splice(outIndex, 2);
    }
    if (
      !query ||
      !rest.length ||
      rest.some((p) => p.startsWith('--')) ||
      (command === 'replace' && replacement === undefined)
    )
      throw new Error(usage);
    const workspace = new Workspace();
    for (const file of rest) workspace.open(basename(file), await readFile(file));
    if (command === 'search') print(workspace.search(query));
    else {
      const preview = workspace.previewReplace(query, replacement!);
      print(preview);
      if (output) {
        workspace.commit(preview.id);
        await exportWorkspace(
          output,
          workspace.documents().map((d) => ({ name: d.name, pkg: workspace.package(d.id) })),
        );
        console.log(`Exported ${workspace.documents().length} documents to ${output}`);
      }
    }
  } else throw new Error(usage);
}
main(process.argv.slice(2)).catch((error) => {
  console.error(String(error));
  process.exitCode = 1;
});
