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

it('raw api PATCH fetches the etag itself, and --header overrides it', async () => {
  servers.graph.setReply((request, response) => {
    if (request.method === 'GET') json(response, { '@odata.etag': 'auto-etag' });
    else if (!request.headers['if-match']) json(response, { error: { message: 'missing If-Match' } }, 428);
    else json(response, {});
  });
  servers.graph.reset();
  const auto = await cli.run(['api', 'PATCH', `/planner/tasks/${taskId}`, '--data', '{"percentComplete":50}']);
  assert.equal(auto.code, 0, auto.stderr);
  assert.equal(servers.graph.requests.find(({ method }) => method === 'PATCH')?.headers['if-match'], 'auto-etag');

  servers.graph.reset();
  const manual = await cli.run(['api', 'PATCH', `/planner/tasks/${taskId}`, '--header', 'If-Match: W/"mine"', '--data', '{}']);
  assert.equal(manual.code, 0, manual.stderr);
  assert.deepEqual(servers.graph.requests.map(({ method }) => method), ['PATCH']);
  assert.equal(servers.graph.requests[0]?.headers['if-match'], 'W/"mine"');
});

it('includes the Graph error message and retries a 429 after Retry-After', async () => {
  servers.graph.setReply((_request, response) => json(response, { error: { code: 'BadRequest', message: 'planId is invalid' } }, 400));
  servers.graph.reset();
  const bad = await cli.run(['api', 'POST', '/planner/tasks', '--data', '{}']);
  assert.equal(bad.code, 1);
  assert.match(bad.stderr, /planId is invalid/);

  let calls = 0;
  servers.graph.setReply((_request, response) => {
    calls += 1;
    if (calls === 1) { response.writeHead(429, { 'Retry-After': '0', 'Content-Type': 'application/json' }); response.end('{}'); }
    else json(response, { ok: true });
  });
  servers.graph.reset();
  const throttled = await cli.run(['api', 'GET', '/me']);
  assert.equal(throttled.code, 0, throttled.stderr);
  assert.equal(servers.graph.requests.length, 2);
});

it('shows Assignee names for one Task', async () => {
  servers.graph.setReply((request, response) => {
    if (request.path.startsWith('/planner/tasks/')) json(response, { id: taskId, title: 'Owned', planId, assignments: { [userId]: {} } });
    else if (request.path === `/planner/plans/${planId}`) json(response, { id: planId, owner: 'grp' });
    else if (request.path.startsWith('/groups/grp/members')) json(response, { value: [{ id: userId, displayName: 'Alex Example' }] });
    else json(response, { value: [] });
  });
  servers.graph.reset();
  const result = await cli.run(['task', taskId]);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /^assignees: Alex Example$/m);
});

it('applies tasks filters across pages instead of only the first page', async () => {
  servers.graph.setReply((request, response) => {
    if (request.path === '/me/planner/plans') json(response, { value: [] });
    else if (request.path.startsWith(`/planner/plans/${planId}/tasks?`)) {
      json(response, { value: [{ id: 'done1', title: 'Finished', percentComplete: 100 }], '@odata.nextLink': `${servers.graph.url}/page2` });
    } else if (request.path === '/page2') json(response, { value: [{ id: 'open1', title: 'Still open', percentComplete: 0 }] });
    else json(response, { value: [] });
  });
  servers.graph.reset();
  const result = await cli.run(['tasks', '--open']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /Still open/);
  assert.doesNotMatch(result.stdout, /Finished/);
});

it('treats bulk create strings as literal text, never as @file or stdin', async () => {
  servers.graph.setReply((request, response) => {
    if (lookups(request, response)) return;
    if (request.method === 'POST') json(response, { id: 'LiteralAAAAAAAAAAAAAAAAAAAA' });
    else if (request.method === 'GET') json(response, { '@odata.etag': 'd1' });
    else json(response, {});
  });
  servers.graph.reset();
  const result = await cli.run(['add', '--data', '-'], {}, `${JSON.stringify({ title: 'At sign', description: '@alice please review' })}\n`);
  assert.equal(result.code, 0, result.stderr);
  const patch = servers.graph.requests.find(({ method, path }) => method === 'PATCH' && path.endsWith('/details'));
  assert.equal(JSON.parse(patch?.body ?? '{}').description, '@alice please review');
});

it('uses a 28-character Plan id directly when it is not in the Plan list', async () => {
  servers.graph.setReply((request, response) => {
    if (request.path === '/me/planner/plans') json(response, { value: [] });
    else json(response, { value: [] });
  });
  servers.graph.reset();
  const result = await cli.run(['tasks']);
  assert.equal(result.code, 0, result.stderr);
  assert.ok(servers.graph.requests.some(({ path }) => path.startsWith(`/planner/plans/${planId}/tasks`)));
});

it('bulk update keeps going after a per-Task usage error', async () => {
  const other = 'OtherTaskAAAAAAAAAAAAAAAAAAA';
  servers.graph.setReply((request, response) => {
    if (lookups(request, response)) return;
    if (request.method === 'GET') json(response, { id: taskId, planId, '@odata.etag': 'e1' });
    else json(response, {});
  });
  servers.graph.reset();
  const result = await cli.run(['update', taskId, other, '--bucket', 'Missing']);
  assert.equal(result.code, 1);
  assert.match(result.stderr, new RegExp(`#${taskId}:`));
  assert.match(result.stderr, new RegExp(`#${other}:`));
});
