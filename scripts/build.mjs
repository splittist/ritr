import { build as bundle } from 'esbuild';
import { build as ui } from 'vite';

await bundle({
  entryPoints: ['src/desktop/main.ts', 'src/desktop/preload.ts'],
  outdir: 'dist/desktop',
  outExtension: { '.js': '.cjs' },
  bundle: true,
  platform: 'node',
  format: 'cjs',
  external: ['electron'],
});
await ui();
