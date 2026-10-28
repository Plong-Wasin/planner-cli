import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

const planId = '00000000-0000-4000-8000-000000000001';
const bucketId = '00000000-0000-4000-8000-000000000002';
const userId = '00000000-0000-4000-8000-000000000003';

let servers: FakeServers;
let cli: Cli;

before(async () => {
  servers = await startFakeServers();
  cli = createCli({ PLANNER_GRAPH_URL: servers.graph.url });
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: {
    accessToken: 'add-token', defaultPlan: planId, defaultBucket: bucketId,
  } } }));
});

after(async () => {
  await servers.close();
  cli.cleanup();
});

it('creates a Task in the Default Bucket assigned to me at important priority', async () => {
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/me') response.end(JSON.stringify({ id: userId }));
    else if (request.path === '/planner/tasks') response.end(JSON.stringify({ id: 'task-created' }));
    else response.end(JSON.stringify({ value: [] }));
  });
  servers.graph.reset();

  const result = await cli.run(['add', 'Write release notes']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, 'task-created\n');
  assert.equal(result.stderr, '');
  assert.deepEqual(servers.graph.requests.map(({ method, path }) => [method, path]), [
    ['GET', '/me'], ['POST', '/planner/tasks'],
  ]);
  assert.deepEqual(JSON.parse(servers.graph.requests[1]?.body ?? ''), {
    planId,
    bucketId,
    title: 'Write release notes',
    assignments: { [userId]: { '@odata.type': 'microsoft.graph.plannerAssignment', orderHint: ' !' } },
    priority: 3,
  });
});

it('resolves names and writes optional fields and Details with the current etag', async () => {
  const checklistPath = join(cli.home, 'checklist.txt');
  writeFileSync(checklistPath, 'Draft outline\nReview with team\n');
  let detailsWrites = 0;
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/me/planner/plans') response.end(JSON.stringify({ value: [{ id: planId, title: 'Roadmap' }] }));
    else if (request.path === `/planner/plans/${planId}/buckets`) response.end(JSON.stringify({ value: [{ id: bucketId, name: 'Ready' }] }));
    else if (request.path === `/planner/plans/${planId}`) response.end(JSON.stringify({ id: planId, owner: 'group-1' }));
    else if (request.path === '/groups/group-1/members') response.end(JSON.stringify({ value: [{ id: userId, displayName: 'Ada Planner', mail: 'ada@example.test' }] }));
    else if (request.method === 'POST' && request.path === '/planner/tasks') response.end(JSON.stringify({ id: 'task-named' }));
    else if (request.method === 'GET' && request.path === '/planner/tasks/task-named/details') response.end(JSON.stringify({ '@odata.etag': 'details-etag' }));
    else if (request.method === 'PATCH' && request.path === '/planner/tasks/task-named/details') {
      detailsWrites += 1;
      response.end('{}');
    } else response.end('{}');
  });
  servers.graph.reset();

  const result = await cli.run([
    'add', 'Prepare roadmap', '--bucket', 'Ready', '--assign', 'ada@example.test', '--priority', 'urgent',
    '--due', '2026-02-03', '--start', '2026-02-01T08:30:00+07:00',
    '--description', 'Coordinate the release', '--checklist', `@${checklistPath}`,
  ]);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, 'task-named\n');
  const create = servers.graph.requests.find((request) => request.method === 'POST' && request.path === '/planner/tasks');
  assert.deepEqual(JSON.parse(create?.body ?? ''), {
    planId,
    bucketId,
    title: 'Prepare roadmap',
    priority: 1,
    dueDateTime: '2026-02-03T00:00:00.000Z',
    startDateTime: '2026-02-01T08:30:00+07:00',
    assignments: { [userId]: { '@odata.type': 'microsoft.graph.plannerAssignment', orderHint: ' !' } },
  });
  const detailsGet = servers.graph.requests.find((request) => request.method === 'GET' && request.path === '/planner/tasks/task-named/details');
  const detailsPatch = servers.graph.requests.find((request) => request.method === 'PATCH' && request.path === '/planner/tasks/task-named/details');
  assert.equal(detailsGet?.headers.authorization, 'Bearer add-token');
  assert.equal(detailsPatch?.headers['if-match'], 'details-etag');
  const writtenDetails = JSON.parse(detailsPatch?.body ?? '') as { description: string; checklist: Record<string, { title: string; isChecked: boolean; '@odata.type': string }> };
  assert.equal(writtenDetails.description, 'Coordinate the release');
  assert.deepEqual(Object.values(writtenDetails.checklist).map(({ title, isChecked, '@odata.type': type }) => [title, isChecked, type]), [
    ['Draft outline', false, 'microsoft.graph.plannerChecklistItem'],
    ['Review with team', false, 'microsoft.graph.plannerChecklistItem'],
  ]);
  assert.equal(detailsWrites, 1);
});

it('leaves a new Task unassigned when --no-assign is set and accepts a raw priority number', async () => {
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.method === 'POST') response.end(JSON.stringify({ id: 'task-unassigned' }));
    else response.end('{}');
  });
  servers.graph.reset();

  const result = await cli.run(['add', 'Unassigned capture', '--no-assign', '--priority', '7']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, 'task-unassigned\n');
  const body = JSON.parse(servers.graph.requests.find(({ method }) => method === 'POST')?.body ?? '') as Record<string, unknown>;
  assert.equal(body.priority, 7);
  assert.equal('assignments' in body, false);
  assert.equal(servers.graph.requests.some(({ path }) => path === '/me'), false);
});

it('reads Details text from stdin', async () => {
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.method === 'POST') response.end(JSON.stringify({ id: 'task-stdin' }));
    else if (request.method === 'GET' && request.path.endsWith('/details')) response.end(JSON.stringify({ '@odata.etag': 'stdin-etag' }));
    else response.end('{}');
  });
  servers.graph.reset();

  const result = await cli.run(['add', 'Task from stdin', '--no-assign', '--description', '-', '--checklist', 'First item'], {}, 'Description from stdin');

  assert.equal(result.code, 0, result.stderr);
  const patch = servers.graph.requests.find(({ method }) => method === 'PATCH');
  assert.equal(JSON.parse(patch?.body ?? '').description, 'Description from stdin');
});

it('re-fetches Details once after a 412 before retrying with the fresh etag', async () => {
  let etagRead = 0;
  let patchCount = 0;
  servers.graph.setReply((request, response) => {
    if (request.method === 'PATCH') {
      patchCount += 1;
      if (patchCount === 1) {
        response.writeHead(412, { 'Content-Type': 'application/json' });
        response.end('{}');
      } else {
        response.writeHead(200, { 'Content-Type': 'application/json' });
        response.end('{}');
      }
      return;
    }
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.method === 'POST') response.end(JSON.stringify({ id: 'task-retry' }));
    else if (request.path.endsWith('/details')) {
      etagRead += 1;
      response.end(JSON.stringify({ '@odata.etag': `etag-${etagRead}` }));
    } else response.end('{}');
  });
  servers.graph.reset();

  const result = await cli.run(['add', 'Add description', '--no-assign', '--description', 'Details text']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(patchCount, 2);
  assert.deepEqual(servers.graph.requests.filter(({ method }) => method === 'PATCH').map(({ headers }) => headers['if-match']), ['etag-1', 'etag-2']);
});

it('reports a second Details conflict after exactly one retry', async () => {
  let etagRead = 0;
  let patchCount = 0;
  servers.graph.setReply((request, response) => {
    if (request.method === 'PATCH') {
      patchCount += 1;
      response.writeHead(412, { 'Content-Type': 'application/json' });
      response.end('{}');
      return;
    }
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.method === 'POST') response.end(JSON.stringify({ id: 'task-conflict' }));
    else if (request.path.endsWith('/details')) {
      etagRead += 1;
      response.end(JSON.stringify({ '@odata.etag': `etag-${etagRead}` }));
    } else response.end('{}');
  });
  servers.graph.reset();

  const result = await cli.run(['add', 'Conflicting details', '--no-assign', '--description', 'Some text']);

  assert.equal(result.code, 1);
  assert.match(result.stderr, /Task Details for 'task-conflict' changed twice/);
  assert.equal(patchCount, 2);
});

it('creates valid NDJSON Tasks around a bad line and reports the line error', async () => {
  let taskNumber = 0;
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.method === 'POST' && request.path === '/planner/tasks') {
      taskNumber += 1;
      response.end(JSON.stringify({ id: `bulk-${taskNumber}` }));
    } else if (request.path === '/me') response.end(JSON.stringify({ id: userId }));
    else response.end(JSON.stringify({ value: [] }));
  });
  servers.graph.reset();

  const result = await cli.run(['add', '--data', '-'], {}, [
    JSON.stringify({ title: 'First task', assign: 'me' }),
    JSON.stringify({ title: 'Bad priority', priority: 'immediate' }),
    JSON.stringify({ title: 'Last task', assign: 'me' }),
  ].join('\n'));

  assert.equal(result.code, 1);
  assert.equal(result.stdout, 'bulk-1\nbulk-2\n');
  assert.match(result.stderr, /line 2: Priority must be urgent, important, medium, low, or a number from 0 to 10\./);
  assert.equal(servers.graph.requests.filter(({ method, path }) => method === 'POST' && path === '/planner/tasks').length, 2);
});

it('creates Tasks from a JSON array with CLI names and default values', async () => {
  const dataPath = join(cli.home, 'tasks.json');
  writeFileSync(dataPath, JSON.stringify([
    { title: 'Named bucket task', bucket: 'Ready', assign: 'Ada Planner', priority: 'urgent', due: '2026-04-05', start: '2026-04-01', description: 'Bulk details', checklist: 'Step one\nStep two' },
    { title: 'Default task' },
  ]));
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/me/planner/plans') response.end(JSON.stringify({ value: [{ id: planId, title: 'Roadmap' }] }));
    else if (request.path === `/planner/plans/${planId}/buckets`) response.end(JSON.stringify({ value: [{ id: bucketId, name: 'Ready' }] }));
    else if (request.path === `/planner/plans/${planId}`) response.end(JSON.stringify({ id: planId, owner: 'group-1' }));
    else if (request.path === '/groups/group-1/members') response.end(JSON.stringify({ value: [{ id: userId, displayName: 'Ada Planner', mail: 'ada@example.test' }] }));
    else if (request.path === '/me') response.end(JSON.stringify({ id: userId }));
    else if (request.method === 'POST' && request.path === '/planner/tasks') {
      const count = servers.graph.requests.filter(({ method, path }) => method === 'POST' && path === '/planner/tasks').length;
      response.end(JSON.stringify({ id: `array-${count}` }));
    } else if (request.path.endsWith('/details')) response.end(JSON.stringify({ '@odata.etag': 'bulk-etag' }));
    else response.end('{}');
  });
  servers.graph.reset();

  const result = await cli.run(['add', '--data', `@${dataPath}`]);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, 'array-1\narray-2\n');
  const creates = servers.graph.requests.filter(({ method, path }) => method === 'POST' && path === '/planner/tasks');
  assert.deepEqual(JSON.parse(creates[0]?.body ?? ''), {
    planId, bucketId, title: 'Named bucket task', priority: 1,
    dueDateTime: '2026-04-05T00:00:00.000Z',
    startDateTime: '2026-04-01T00:00:00.000Z',
    assignments: { [userId]: { '@odata.type': 'microsoft.graph.plannerAssignment', orderHint: ' !' } },
  });
  assert.deepEqual(JSON.parse(creates[1]?.body ?? ''), {
    planId, bucketId, title: 'Default task', priority: 3,
    assignments: { [userId]: { '@odata.type': 'microsoft.graph.plannerAssignment', orderHint: ' !' } },
  });
  const detailsPatch = servers.graph.requests.find(({ method }) => method === 'PATCH');
  const detailsBody = JSON.parse(detailsPatch?.body ?? '') as { description: string; checklist: Record<string, { title: string }> };
  assert.equal(detailsBody.description, 'Bulk details');
  assert.deepEqual(Object.values(detailsBody.checklist).map(({ title }) => title), ['Step one', 'Step two']);
});
