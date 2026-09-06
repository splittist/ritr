import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync } from 'node:fs';
import { resolve, join } from 'node:path';

// Point DOCXFIX_ROOT at another checkout when docxfix is not a sibling repository.
const root = resolve(process.env.DOCXFIX_ROOT ?? '../docxfix');
const python = join(
  root,
  '.venv',
  process.platform === 'win32' ? 'Scripts/python.exe' : 'bin/python',
);
if (!existsSync(python)) throw new Error(`Missing docxfix virtual environment: ${python}`);
mkdirSync('fixtures/generated', { recursive: true });
for (const name of ['plain', 'review', 'sections']) {
  const result = spawnSync(
    python,
    [
      '-m',
      'docxfix.cli',
      'create',
      resolve(`fixtures/generated/${name}.docx`),
      '--spec',
      resolve(`fixtures/specs/${name}.json`),
    ],
    {
      cwd: root,
      stdio: 'inherit',
      env: { ...process.env, PYTHONPATH: join(root, 'src'), PYTHONUTF8: '1' },
    },
  );
  if (result.status !== 0)
    throw new Error(`docxfix failed for ${name}: ${result.error ?? result.status}`);
}
