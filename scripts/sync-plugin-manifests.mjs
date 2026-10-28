import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const source = JSON.parse(await readFile(resolve(root, 'plugins/planner-cli/plugin.json'), 'utf8'));
const generated = new Map([
  ['.claude-plugin/marketplace.json', {
    name: source.name,
    description: 'Planner CLI agent skills',
    owner: { name: source.developerName },
    plugins: [{
      name: source.name,
      source: './plugins/planner-cli',
      description: source.description,
    }],
  }],
  ['.agents/plugins/marketplace.json', {
    name: source.name,
    interface: { displayName: source.displayName },
    plugins: [{
      name: source.name,
      source: { source: 'local', path: './plugins/planner-cli' },
      policy: { installation: 'AVAILABLE', authentication: 'ON_USE' },
      category: source.category,
    }],
  }],
  ['plugins/planner-cli/.claude-plugin/plugin.json', {
    name: source.name,
    description: source.description,
    author: { name: source.developerName },
    license: source.license,
  }],
  ['plugins/planner-cli/.codex-plugin/plugin.json', {
    name: source.name,
    description: source.description,
    author: { name: source.developerName },
    license: source.license,
    skills: './skills/',
    interface: {
      displayName: source.displayName,
      shortDescription: source.shortDescription,
      developerName: source.developerName,
      category: source.category,
      capabilities: [],
    },
  }],
]);

const check = process.argv.includes('--check');
let stale = false;
for (const [relativePath, value] of generated) {
  const path = resolve(root, relativePath);
  const contents = `${JSON.stringify(value, null, 2)}\n`;
  if (check) {
    let current;
    try { current = await readFile(path, 'utf8'); }
    catch { current = ''; }
    if (current !== contents) {
      console.error(`${relativePath} is stale; run node scripts/sync-plugin-manifests.mjs`);
      stale = true;
    }
  } else {
    await mkdir(dirname(path), { recursive: true });
    await writeFile(path, contents);
  }
}
if (stale) process.exitCode = 1;
