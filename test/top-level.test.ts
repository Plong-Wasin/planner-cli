import assert from 'node:assert/strict';
import { after, before, it } from 'node:test';
import { readFileSync } from 'node:fs';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

const pkg = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')) as { version: string };

let servers: FakeServers;
let cli: Cli;

before(async () => {
  servers = await startFakeServers();
  cli = createCli({
    PLANNER_GRAPH_URL: servers.graph.url,
    PLANNER_LOGIN_URL: servers.login.url,
  });
});

after(async () => {
  await servers.close();
  cli.cleanup();
});

it('prints the manifest version from the bundled script', async () => {
  const result = await cli.run(['--version']);
  assert.equal(result.code, 0);
  assert.equal(result.stdout.trim(), pkg.version);
  assert.equal(result.stderr, '');
});

it('prints help, and reports an unknown command as a usage error', async () => {
  const help = await cli.run(['--help']);
  assert.equal(help.code, 0);
  assert.match(help.stdout, /Commands:/);
  for (const command of ['auth', 'config', 'me', 'plans', 'buckets', 'members', 'tasks', 'task', 'add', 'update', 'done', 'delete', 'api']) {
    assert.match(help.stdout, new RegExp(`\\b${command}\\b`));
  }
  for (const option of ['--profile', '--client-id', '--plan', '--bucket', '--assign', '--priority', '--due', '--start', '--description', '--checklist', '--data', '--no-assign', '--percent', '--yes', '--mine', '--open', '--created-after', '--completed-after', '--fields', '--details', '--cursor', '--all', '--json']) {
    assert.match(help.stdout, new RegExp(option.replaceAll('-', '\\-')));
  }
  assert.equal(help.stderr, '');

  const unknown = await cli.run(['not-a-planner-command']);
  assert.equal(unknown.code, 2);
  assert.match(unknown.stderr, /^planner: /);
  assert.equal(unknown.stdout, '');
});
