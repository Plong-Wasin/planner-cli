import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

const planId = '00000000-0000-4000-8000-000000000001';
const bucketId = '00000000-0000-4000-8000-000000000002';
const task1 = '00000000-0000-4000-8000-000000000011';
const task2 = '00000000-0000-4000-8000-000000000012';
const task3 = '00000000-0000-4000-8000-000000000013';
const task4 = '00000000-0000-4000-8000-000000000014';
const task5 = '00000000-0000-4000-8000-000000000015';
const task6 = '00000000-0000-4000-8000-000000000016';
const taskBad = '00000000-0000-4000-8000-0000000000ba';
const taskGood = '00000000-0000-4000-8000-0000000000bc';

let servers: FakeServers;
let cli: Cli;

before(async () => {
  servers = await startFakeServers();
  cli = createCli({ PLANNER_GRAPH_URL: servers.graph.url });
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: {
    accessToken: 'bulk-token', defaultPlan: planId, defaultBucket: bucketId,
  } } }));
});

after(async () => {
  await servers.close();
  cli.cleanup();
});

it('updates a Task with an etag precondition and add-style fields', async () => {
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.method === 'GET' && request.path === `/planner/tasks/${task1}`) {
      response.end(JSON.stringify({ id: task1, planId, '@odata.etag': 'task-etag-1' }));
    } else {
      response.end('{}');
    }
  });
  servers.graph.reset();

  const result = await cli.run(['update', task1, '--bucket', bucketId, '--priority', 'urgent', '--percent', '50', '--due', '2026-03-04']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stderr, '');
  assert.deepEqual(servers.graph.requests.map(({ method, path }) => [method, path]), [
    ['GET', `/planner/tasks/${task1}`], ['GET', `/planner/tasks/${task1}`], ['PATCH', `/planner/tasks/${task1}`],
  ]);
  const patch = servers.graph.requests[2];
  assert.equal(patch?.headers['if-match'], 'task-etag-1');
  assert.deepEqual(JSON.parse(patch?.body ?? ''), {
    bucketId,
    priority: 1,
    percentComplete: 50,
    dueDateTime: '2026-03-04T00:00:00.000Z',
  });
});

it('marks each done Task complete with an etag precondition', async () => {
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.method === 'GET') response.end(JSON.stringify({ '@odata.etag': 'done-etag' }));
    else response.end('{}');
  });
  servers.graph.reset();

  const result = await cli.run(['done', task2]);

  assert.equal(result.code, 0, result.stderr);
  const patch = servers.graph.requests.find(({ method }) => method === 'PATCH');
  assert.equal(patch?.path, `/planner/tasks/${task2}`);
  assert.equal(patch?.headers['if-match'], 'done-etag');
  assert.deepEqual(JSON.parse(patch?.body ?? ''), { percentComplete: 100 });
});

it('re-fetches the Task etag once after an update conflict', async () => {
  let reads = 0;
  let patches = 0;
  servers.graph.setReply((request, response) => {
    if (request.method === 'PATCH') {
      patches += 1;
      if (patches === 1) response.writeHead(412).end('{}');
      else response.writeHead(204).end();
      return;
    }
    response.writeHead(200, { 'Content-Type': 'application/json' });
    reads += 1;
    response.end(JSON.stringify({ '@odata.etag': `update-etag-${reads}` }));
  });
  servers.graph.reset();

  const result = await cli.run(['update', task5, '--percent', '25']);

  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.filter(({ method }) => method === 'PATCH').map(({ headers }) => headers['if-match']), [
    'update-etag-1', 'update-etag-2',
  ]);
});

it('updates assignment and Details fields with their separate etag preconditions', async () => {
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.method === 'GET' && request.path.endsWith('/details')) {
      response.end(JSON.stringify({ '@odata.etag': 'details-etag' }));
    } else if (request.method === 'GET') {
      response.end(JSON.stringify({ assignments: { 'user-1': {} }, '@odata.etag': 'task-etag' }));
    } else response.end('{}');
  });
  servers.graph.reset();

  const result = await cli.run(['update', task6, '--no-assign', '--description', 'Updated notes', '--checklist', 'Review\nShip']);

  assert.equal(result.code, 0, result.stderr);
  const taskPatch = servers.graph.requests.find(({ method, path }) => method === 'PATCH' && path === `/planner/tasks/${task6}`);
  assert.equal(taskPatch?.headers['if-match'], 'task-etag');
  assert.deepEqual(JSON.parse(taskPatch?.body ?? ''), { assignments: { 'user-1': null } });
  const detailsPatch = servers.graph.requests.find(({ method, path }) => method === 'PATCH' && path.endsWith('/details'));
  assert.equal(detailsPatch?.headers['if-match'], 'details-etag');
  const detailBody = JSON.parse(detailsPatch?.body ?? '') as { description: string; checklist: Record<string, { title: string }> };
  assert.equal(detailBody.description, 'Updated notes');
  assert.deepEqual(Object.values(detailBody.checklist).map(({ title }) => title), ['Review', 'Ship']);
  assert.equal(Object.keys(detailBody.checklist).every((id) => /^[0-9a-f-]{36}$/i.test(id)), true);
});

it('refuses delete without --yes without making a Graph request', async () => {
  servers.graph.reset();

  const result = await cli.run(['delete', task3]);

  assert.equal(result.code, 2);
  assert.match(result.stderr, /without '--yes'/);
  assert.equal(servers.graph.requests.length, 0);
});

it('deletes with an etag and retries once after a 412 using a fresh etag', async () => {
  let reads = 0;
  let deletes = 0;
  servers.graph.setReply((request, response) => {
    if (request.method === 'DELETE') {
      deletes += 1;
      if (deletes === 1) response.writeHead(412).end('{}');
      else response.writeHead(204).end();
      return;
    }
    response.writeHead(200, { 'Content-Type': 'application/json' });
    reads += 1;
    response.end(JSON.stringify({ '@odata.etag': `delete-etag-${reads}` }));
  });
  servers.graph.reset();

  const result = await cli.run(['delete', task4, '--yes']);

  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.filter(({ method }) => method === 'DELETE').map(({ headers }) => headers['if-match']), [
    'delete-etag-1', 'delete-etag-2',
  ]);
});

it('continues after a Task failure and reports each failed id and the final failure count', async () => {
  servers.graph.setReply((request, response) => {
    if (request.path.endsWith(`/${taskBad}`)) {
      response.writeHead(404, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify({ error: 'Not found' }));
    } else {
      response.writeHead(200, { 'Content-Type': 'application/json' });
      if (request.method === 'GET') response.end(JSON.stringify({ '@odata.etag': 'good-etag' }));
      else response.end('{}');
    }
  });
  servers.graph.reset();

  const result = await cli.run(['done', taskBad, taskGood]);

  assert.equal(result.code, 1);
  assert.match(result.stderr, new RegExp(`#${taskBad}: Not found: GET \\/planner\\/tasks\\/${taskBad}`));
  assert.equal(result.stderr.trimEnd().split('\n').at(-1), '1 Task failed.');
  assert.equal(servers.graph.requests.some(({ path }) => path === `/planner/tasks/${taskGood}`), true);
});
