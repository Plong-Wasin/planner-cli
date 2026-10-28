import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

// Real Planner ids are 28-character opaque strings, not GUIDs.
const planId = 'Pl4nAAAAAAAAAAAAAAAAAAAAAAAA';
const bucketId = 'BuckBBBBBBBBBBBBBBBBBBBBBBBB';
const taskId = '4Q81bX_cmkWVWe3vseex5ckAIV4d';
const userId = '00000000-0000-4000-8000-000000000003';

let servers: FakeServers;
let cli: Cli;

before(async () => {
  servers = await startFakeServers();
  cli = createCli({ PLANNER_GRAPH_URL: servers.graph.url });
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: {
    accessToken: 'fix-token', defaultPlan: planId, defaultBucket: bucketId,
  } } }));
});

after(async () => {
  await servers.close();
  cli.cleanup();
});

function json(response: { writeHead(code: number, headers: Record<string, string>): void; end(body: string): void }, body: unknown, code = 200): void {
  response.writeHead(code, { 'Content-Type': 'application/json' });
  response.end(JSON.stringify(body));
}

function lookups(request: { method: string; path: string }, response: Parameters<typeof json>[0]): boolean {
  if (request.method !== 'GET') return false;
  if (request.path === '/me/planner/plans') json(response, { value: [{ id: planId, title: 'Roadmap', owner: 'grp' }] });
  else if (request.path === `/planner/plans/${planId}/buckets`) json(response, { value: [{ id: bucketId, name: 'Ready' }] });
  else if (request.path === '/me') json(response, { id: userId });
  else return false;
  return true;
}

it('accepts real 28-character Task ids in task, done and delete, and still rejects malformed ids', async () => {
  servers.graph.setReply((request, response) => {
    if (request.method === 'GET') json(response, { id: taskId, title: 'Real id', '@odata.etag': 'e1' });
    else json(response, {});
  });
  servers.graph.reset();

  const shown = await cli.run(['task', taskId]);
  assert.equal(shown.code, 0, shown.stderr);
  const done = await cli.run(['done', taskId]);
  assert.equal(done.code, 0, done.stderr);
  const deleted = await cli.run(['delete', taskId, '--yes']);
  assert.equal(deleted.code, 0, deleted.stderr);
  const viaUrl = await cli.run(['task', `https://planner.example.test/Home/Task/${taskId}`]);
  assert.equal(viaUrl.code, 0, viaUrl.stderr);

  const short = await cli.run(['task', 'abc123']);
  assert.equal(short.code, 2);
  assert.match(short.stderr, /Invalid Task id/);
});

it('prints a compact key/value block for one Task instead of the raw JSON', async () => {
  servers.graph.setReply((_request, response) => json(response, {
    id: taskId, title: 'Compact view', percentComplete: 40, priority: 3, dueDateTime: '2026-11-01T00:00:00Z', bucketId,
    details: { description: 'Body text', checklist: { a: { title: 'Step one', isChecked: true } } },
  }));
  servers.graph.reset();

  const result = await cli.run(['task', taskId]);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /^title: Compact view$/m);
  assert.match(result.stdout, /^percent: 40%$/m);
  assert.match(result.stdout, /Body text/);
  assert.match(result.stdout, /\[x\] Step one/);
  assert.doesNotMatch(result.stdout, /^\{/);
});

it('bulk create accepts a checklist array and a numeric priority', async () => {
  servers.graph.setReply((request, response) => {
    if (lookups(request, response)) return;
    if (request.method === 'POST') json(response, { id: 'NewTaskAAAAAAAAAAAAAAAAAAAA' });
    else if (request.path.endsWith('/details') && request.method === 'GET') json(response, { '@odata.etag': 'd1' });
    else json(response, {});
  });
  servers.graph.reset();

  const result = await cli.run(['add', '--data', '-'], {}, JSON.stringify([
    { title: 'Typed', priority: 5, checklist: ['one', 'two'] },
  ]));
  assert.equal(result.code, 0, result.stderr);
  const post = servers.graph.requests.find(({ method }) => method === 'POST');
  assert.equal(JSON.parse(post?.body ?? '{}').priority, 5);
  const patch = servers.graph.requests.find(({ method, path }) => method === 'PATCH' && path.endsWith('/details'));
  const items = Object.values(JSON.parse(patch?.body ?? '{}').checklist) as Array<{ title: string }>;
  assert.deepEqual(items.map((item) => item.title), ['one', 'two']);
});

it('bulk create prints the new id before a Details failure so the Task is not orphaned silently', async () => {
  servers.graph.setReply((request, response) => {
    if (lookups(request, response)) return;
    if (request.method === 'POST') json(response, { id: 'OrphanAAAAAAAAAAAAAAAAAAAAA' });
    else if (request.method === 'GET') json(response, { '@odata.etag': 'd1' });
    else json(response, { error: 'boom' }, 500);
  });
  servers.graph.reset();

  const result = await cli.run(['add', '--data', '-'], {}, `${JSON.stringify({ title: 'Half done', description: 'text' })}\n`);
  assert.equal(result.code, 1);
  assert.match(result.stdout, /OrphanAAAAAAAAAAAAAAAAAAAAA/);
  assert.match(result.stderr, /line 1: .*OrphanAAAAAAAAAAAAAAAAAAAAA/);
});

it('bulk create resolves each Plan, Bucket and the current user once per run', async () => {
  servers.graph.setReply((request, response) => {
    if (request.method === 'POST') json(response, { id: `T${servers.graph.requests.length}`.padEnd(28, 'X') });
    else if (request.path === '/me/planner/plans') json(response, { value: [{ id: planId, title: 'Roadmap', owner: 'grp' }] });
    else if (request.path === `/planner/plans/${planId}/buckets`) json(response, { value: [{ id: bucketId, name: 'Ready' }] });
    else if (request.path === '/me') json(response, { id: userId });
    else json(response, {});
  });
  servers.graph.reset();

  const lines = [1, 2, 3].map((n) => JSON.stringify({ title: `Task ${n}`, bucket: 'Ready' })).join('\n');
  const result = await cli.run(['--plan', 'Roadmap', 'add', '--data', '-'], {}, lines);
  assert.equal(result.code, 0, result.stderr);
  const count = (path: string): number => servers.graph.requests.filter((request) => request.method === 'GET' && request.path === path).length;
  assert.equal(count('/me/planner/plans'), 1);
  assert.equal(count(`/planner/plans/${planId}/buckets`), 1);
  assert.equal(count('/me'), 1);
});
