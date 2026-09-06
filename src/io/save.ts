import { link, mkdir, mkdtemp, open, readFile, rename, rm, lstat } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { randomUUID } from 'node:crypto';
import { DocxPackage, comparePackages } from '../package/docx';

async function stage(path: string, pkg: DocxPackage): Promise<void> {
  const bytes = pkg.save();
  const reopened = DocxPackage.open(bytes);
  if (comparePackages(pkg, reopened).some((p) => p.status !== 'unchanged'))
    throw new Error('Save validation changed package parts');
  const file = await open(path, 'wx');
  try {
    await file.writeFile(bytes);
    await file.sync();
  } finally {
    await file.close();
  }
  const disk = DocxPackage.open(await readFile(path));
  if (comparePackages(pkg, disk).some((p) => p.status !== 'unchanged'))
    throw new Error('Saved file verification failed');
}

/** Publish using an exclusive hard link: an existing destination can never be replaced. */
export async function saveAs(destination: string, pkg: DocxPackage): Promise<void> {
  const target = resolve(destination);
  const temporary = join(dirname(target), `.ritr-${randomUUID()}.tmp`);
  try {
    await stage(temporary, pkg);
    await link(temporary, target);
  } finally {
    await rm(temporary, { force: true });
  }
}

/** Batch Save As publishes a completely staged NEW directory, never replaces originals. */
export async function exportWorkspace(
  destination: string,
  outputs: { name: string; pkg: DocxPackage }[],
): Promise<void> {
  const target = resolve(destination);
  try {
    await lstat(target);
    throw new Error('Output directory already exists; choose a new directory');
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
  }
  const names = new Set<string>();
  for (const output of outputs) {
    if (
      basename(output.name) !== output.name ||
      !output.name.toLowerCase().endsWith('.docx') ||
      names.has(output.name.toLowerCase())
    )
      throw new Error(`Invalid or duplicate output filename: ${output.name}`);
    names.add(output.name.toLowerCase());
  }
  await mkdir(dirname(target), { recursive: true });
  const temporary = await mkdtemp(join(dirname(target), '.ritr-export-'));
  try {
    for (const output of outputs) await stage(join(temporary, output.name), output.pkg);
    await rename(temporary, target);
  } finally {
    // Only this function's freshly-created staging directory is eligible for cleanup.
    await rm(temporary, { recursive: true, force: true });
  }
}
