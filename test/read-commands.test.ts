import assert from 'node:assert/strict';
import { readFileSync, writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;

before(async () => {
  servers = await startFakeServers();
  cli = createCli({ PLANNER_GRAPH_URL: servers.graph.url });
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: { accessToken: 'read-token' } } }));
});

after(async () => {
  await servers.close();
  cli.cleanup();
});

it('lists Plans in a compact table from the current user Plans endpoint', async () => {
  servers.graph.setReply((request, response) => {
    assert.equal(request.method, 'GET');
    assert.equal(request.path, '/me/planner/plans');
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ value: [{ id: 'plan-1', title: 'Roadmap', owner: 'group-1' }] }));
  });

  const result = await cli.run(['plans']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /ID\s+NAME/);
  assert.match(result.stdout, /plan-1\s+Roadmap/);
  assert.equal(result.stderr, '');
  assert.equal(servers.graph.requests[0]?.headers.authorization, 'Bearer read-token');
});

it('lists Tasks compactly by default and omits Details from the Graph request', async () => {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: {
    accessToken: 'read-token', defaultPlan: '00000000-0000-4000-8000-000000000001',
  } } }));
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path.startsWith('/planner/plans/00000000-0000-4000-8000-000000000001/tasks')) {
      response.end(JSON.stringify({ value: [{
        id: 'task-1', title: 'Prepare the release', bucketId: 'bucket-1', percentComplete: 50,
        dueDateTime: '2026-11-03T00:00:00Z', assignments: { 'user-1': {} },
        details: { description: 'private heavy field', checklist: { item1: { title: 'Check dates', isChecked: false } } },
      }] }));
    } else if (request.path === '/planner/plans/00000000-0000-4000-8000-000000000001/buckets') {
      response.end(JSON.stringify({ value: [{ id: 'bucket-1', name: 'Ready' }] }));
    } else if (request.path === '/planner/plans/00000000-0000-4000-8000-000000000001') {
      response.end(JSON.stringify({ owner: 'group-1' }));
    } else if (request.path === '/groups/group-1/members') {
      response.end(JSON.stringify({ value: [{ id: 'user-1', displayName: 'Ada Planner' }] }));
    } else response.end(JSON.stringify({ value: [] }));
  });
  servers.graph.reset();

  const result = await cli.run(['tasks']);

  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /ID\s+TITLE\s+BUCKET\s+PERCENT\s+DUE\s+ASSIGNEES/);
  assert.match(result.stdout, /task-1\s+Prepare the release\s+Ready\s+50%\s+2026-11-03\s+Ada Planner/);
  assert.doesNotMatch(result.stdout, /private heavy field/);
  assert.ok(servers.graph.requests.some(({ path }) => path.startsWith('/planner/plans/00000000-0000-4000-8000-000000000001/tasks?')));
  assert.ok(servers.graph.requests.every(({ path }) => !path.includes('$expand=details')));

  servers.graph.reset();
  const detailed = await cli.run(['tasks', '--details']);
  assert.equal(detailed.code, 0, detailed.stderr);
  assert.match(detailed.stdout, /DESCRIPTION\s+CHECKLIST/);
  assert.match(detailed.stdout, /private heavy field/);
  assert.match(detailed.stdout, /\[ \] Check dates/);
  assert.equal(new URLSearchParams(servers.graph.requests[0]!.path.split('?')[1]).get('$expand'), 'details');
});

it('applies Task filters using the fields those filters need', async () => {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: {
    accessToken: 'read-token', defaultPlan: '00000000-0000-4000-8000-000000000001',
  } } }));
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path.includes('/tasks?')) response.end(JSON.stringify({ value: [
      { id: 'one', title: 'Keep', bucketId: '00000000-0000-4000-8000-000000000002', percentComplete: 25, createdDateTime: '2026-10-02T00:00:00Z', completedDateTime: '2026-10-04T00:00:00Z', assignments: { 'user-1': {} } },
      { id: 'two', title: 'Wrong bucket', bucketId: 'other', percentComplete: 25, createdDateTime: '2026-10-02T00:00:00Z', completedDateTime: '2026-10-04T00:00:00Z', assignments: { 'user-1': {} } },
      { id: 'three', title: 'Completed', bucketId: '00000000-0000-4000-8000-000000000002', percentComplete: 100, createdDateTime: '2026-10-02T00:00:00Z', completedDateTime: '2026-10-04T00:00:00Z', assignments: { 'user-1': {} } },
      { id: 'four', title: 'Not mine', bucketId: '00000000-0000-4000-8000-000000000002', percentComplete: 25, createdDateTime: '2026-10-02T00:00:00Z', completedDateTime: '2026-10-04T00:00:00Z', assignments: { 'user-2': {} } },
      { id: 'five', title: 'Too old', bucketId: '00000000-0000-4000-8000-000000000002', percentComplete: 25, createdDateTime: '2026-09-30T00:00:00Z', completedDateTime: '2026-10-04T00:00:00Z', assignments: { 'user-1': {} } },
      { id: 'six', title: 'Not completed recently', bucketId: '00000000-0000-4000-8000-000000000002', percentComplete: 25, createdDateTime: '2026-10-02T00:00:00Z', completedDateTime: '2026-09-30T00:00:00Z', assignments: { 'user-1': {} } },
    ] }));
    else if (request.path === '/me') response.end(JSON.stringify({ id: 'user-1' }));
    else response.end(JSON.stringify({ value: [] }));
  });
  servers.graph.reset();

  const result = await cli.run(['tasks', '--mine', '--open', '--bucket', '00000000-0000-4000-8000-000000000002', '--created-after', '2026-10-01', '--completed-after', '2026-10-01']);

  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /Keep/);
  for (const rejected of ['Wrong bucket', 'Completed', 'Not mine', 'Too old', 'Not completed recently']) assert.doesNotMatch(result.stdout, new RegExp(rejected));
  const taskRequest = servers.graph.requests.find(({ path }) => path.includes('/tasks?'));
  assert.ok(taskRequest);
  const query = new URLSearchParams(taskRequest.path.slice(taskRequest.path.indexOf('?') + 1));
  assert.equal(query.get('$top'), '50');
  assert.deepEqual(new Set(query.get('$select')?.split(',')), new Set(['id', 'title', 'bucketId', 'percentComplete', 'dueDateTime', 'assignments', 'createdDateTime', 'completedDateTime']));
  assert.ok(servers.graph.requests.some(({ path }) => path === '/me'));
});

it('selects table fields, truncates only table titles, and prints raw Task JSON', async () => {
  const title = 'A'.repeat(70);
  const task = { value: [{ id: 'task-1', title }] };
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify(task));
  });
  servers.graph.reset();

  const tableResult = await cli.run(['tasks', '--fields', 'title,id']);
  assert.equal(tableResult.code, 0, tableResult.stderr);
  assert.match(tableResult.stdout, /TITLE\s+ID/);
  assert.match(tableResult.stdout, new RegExp(`A{47}…`));
  assert.doesNotMatch(tableResult.stdout, new RegExp(`A{48}`));
  const query = new URLSearchParams(servers.graph.requests[0]!.path.split('?')[1]);
  assert.equal(query.get('$select'), 'title,id');

  servers.graph.reset();
  const jsonResult = await cli.run(['tasks', '--fields', 'title,id', '--json']);
  assert.equal(jsonResult.code, 0, jsonResult.stderr);
  assert.deepEqual(JSON.parse(jsonResult.stdout), task);
});

it('shows one Task with Details and accepts only valid ids or Task URLs', async () => {
  const id = '00000000-0000-4000-8000-000000000010';
  const task = { id, title: 'Review plan', details: { description: 'Review the Plan', checklist: { item1: { title: 'Check dates', isChecked: false } } } };
  servers.graph.setReply((request, response) => {
    assert.equal(request.method, 'GET');
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify(task));
  });
  servers.graph.reset();

  const bare = await cli.run(['task', id]);
  assert.equal(bare.code, 0, bare.stderr);
  assert.match(bare.stdout, /Review the Plan/);
  assert.match(bare.stdout, /Check dates/);
  assert.equal(servers.graph.requests[0]?.path, `/planner/tasks/${id}?$expand=details`);

  servers.graph.reset();
  const url = await cli.run(['task', `https://planner.example.test/Home/Task/${id}`]);
  assert.equal(url.code, 0, url.stderr);
  assert.equal(servers.graph.requests[0]?.path, `/planner/tasks/${id}?$expand=details`);

  servers.graph.reset();
  const malformed = await cli.run(['task', 'abc123']);
  assert.equal(malformed.code, 2);
  assert.match(malformed.stderr, /Invalid Task id/);
  assert.equal(servers.graph.requests.length, 0);
});

it('pages through Tasks by default and follows all pages up to the hard cap', async () => {
  let page = 0;
  const nextPage = `${servers.graph.url}/planner/plans/00000000-0000-4000-8000-000000000001/tasks?page=2`;
  servers.graph.setReply((_request, response) => {
    page += 1;
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ value: [{ id: `task-${page}`, title: `Task ${page}` }], '@odata.nextLink': nextPage }));
  });
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: {
    accessToken: 'read-token', defaultPlan: '00000000-0000-4000-8000-000000000001',
  } } }));
  servers.graph.reset();

  const first = await cli.run(['tasks']);
  assert.equal(first.code, 0, first.stderr);
  assert.match(first.stdout, /Task 1/);
  assert.match(first.stderr, /more available: --cursor .* \(or --all\)/);
  assert.equal(new URLSearchParams(servers.graph.requests[0]!.path.split('?')[1]).get('$top'), '50');

  servers.graph.reset();
  page = 0;
  const all = await cli.run(['tasks', '--all']);
  assert.equal(all.code, 0, all.stderr);
  assert.equal(servers.graph.requests.length, 100);
  assert.match(all.stdout, /Task 100/);
  assert.match(all.stderr, /more available: --cursor/);
});

it('prints a cursor hint and follows Graph continuation links with --all', async () => {
  const nextPage = `${servers.graph.url}/me/planner/plans?$skiptoken=next-page`;
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/me/planner/plans') {
      response.end(JSON.stringify({ value: [{ id: 'plan-one', title: 'One' }], '@odata.nextLink': nextPage }));
    } else if (request.path === '/me/planner/plans?$skiptoken=next-page') {
      response.end(JSON.stringify({ value: [{ id: 'plan-two', title: 'Two' }] }));
    } else response.end(JSON.stringify({ value: [] }));
  });
  servers.graph.reset();

  const firstPage = await cli.run(['plans']);
  assert.equal(firstPage.code, 0, firstPage.stderr);
  assert.match(firstPage.stdout, /One/);
  assert.doesNotMatch(firstPage.stdout, /Two/);
  assert.match(firstPage.stderr, /more available: --cursor .* \(or --all\)/);

  servers.graph.reset();
  const all = await cli.run(['plans', '--all']);
  assert.equal(all.code, 0, all.stderr);
  assert.match(all.stdout, /One/);
  assert.match(all.stdout, /Two/);
  assert.equal(all.stderr, '');
  assert.deepEqual(servers.graph.requests.map(({ path }) => path), ['/me/planner/plans', '/me/planner/plans?$skiptoken=next-page']);
});

it('accepts the opaque cursor printed for a later page', async () => {
  const nextPage = `${servers.graph.url}/me/planner/plans?$skiptoken=next-page`;
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify(request.path.includes('skiptoken')
      ? { value: [{ id: 'plan-two', title: 'Two' }] }
      : { value: [{ id: 'plan-one', title: 'One' }], '@odata.nextLink': nextPage }));
  });
  servers.graph.reset();

  const first = await cli.run(['plans']);
  const cursor = first.stderr.match(/--cursor ([^ ]+)/)?.[1];
  assert.ok(cursor);
  servers.graph.reset();
  const next = await cli.run(['plans', '--cursor', cursor]);
  assert.equal(next.code, 0, next.stderr);
  assert.match(next.stdout, /plan-two\s+Two/);
  assert.equal(servers.graph.requests[0]?.path, '/me/planner/plans?$skiptoken=next-page');
});

it('stores named Plan and Bucket defaults as Graph ids', async () => {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: { accessToken: 'read-token' } } }));
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/me/planner/plans') {
      response.end(JSON.stringify({ value: [{ id: '00000000-0000-4000-8000-000000000001', title: 'Roadmap' }] }));
    } else if (request.path === '/planner/plans/00000000-0000-4000-8000-000000000001/buckets') {
      response.end(JSON.stringify({ value: [{ id: '00000000-0000-4000-8000-000000000002', name: 'Ready' }] }));
    } else response.end(JSON.stringify({ value: [] }));
  });

  const plan = await cli.run(['config', 'set', 'plan', 'Roadmap']);
  assert.equal(plan.code, 0, plan.stderr);
  const bucket = await cli.run(['config', 'set', 'bucket', 'Ready']);
  assert.equal(bucket.code, 0, bucket.stderr);
  const saved = JSON.parse(readFileSync(`${cli.home}/config.json`, 'utf8')) as { profiles: Record<string, { defaultPlan?: string; defaultBucket?: string }> };
  assert.equal(saved.profiles.default?.defaultPlan, '00000000-0000-4000-8000-000000000001');
  assert.equal(saved.profiles.default?.defaultBucket, '00000000-0000-4000-8000-000000000002');
});

it('uses --plan before PLANNER_PLAN and Profile defaults', async () => {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: { accessToken: 'read-token', defaultPlan: '00000000-0000-4000-8000-000000000001' } } }));
  const plans = [
    { id: '00000000-0000-4000-8000-000000000001', title: 'Profile Plan' },
    { id: '00000000-0000-4000-8000-000000000002', title: 'Environment Plan' },
    { id: '00000000-0000-4000-8000-000000000003', title: 'Flag Plan' },
  ];
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/me/planner/plans') response.end(JSON.stringify({ value: plans }));
    else response.end(JSON.stringify({ value: [{ id: 'bucket-1', name: 'Ready' }] }));
  });

  const flag = await cli.run(['--plan', 'Flag Plan', 'buckets'], { PLANNER_PLAN: 'Environment Plan' });
  assert.equal(flag.code, 0, flag.stderr);
  assert.ok(servers.graph.requests.some(({ path }) => path === '/planner/plans/00000000-0000-4000-8000-000000000003/buckets'));

  servers.graph.reset();
  const env = await cli.run(['buckets'], { PLANNER_PLAN: 'Environment Plan' });
  assert.equal(env.code, 0, env.stderr);
  assert.ok(servers.graph.requests.some(({ path }) => path === '/planner/plans/00000000-0000-4000-8000-000000000002/buckets'));
});

it('reports unknown and ambiguous Plan names with candidate ids as usage errors', async () => {
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ value: [
      { id: 'plan-a', title: 'Roadmap' },
      { id: 'plan-b', title: 'Roadmap' },
      { id: 'plan-c', title: 'Backlog' },
    ] }));
  });
  const unknown = await cli.run(['config', 'set', 'plan', 'Missing']);
  assert.equal(unknown.code, 2);
  assert.match(unknown.stderr, /Unknown Plan 'Missing'.*plan-a.*plan-b.*plan-c/);
  const ambiguous = await cli.run(['config', 'set', 'plan', 'Roadmap']);
  assert.equal(ambiguous.code, 2);
  assert.match(ambiguous.stderr, /Ambiguous Plan 'Roadmap'.*plan-a.*plan-b/);
});

it('reports unknown and ambiguous Bucket names with candidate ids as usage errors', async () => {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: {
    accessToken: 'read-token', defaultPlan: '00000000-0000-4000-8000-000000000001',
  } } }));
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ value: [
      { id: 'bucket-a', name: 'Ready' },
      { id: 'bucket-b', name: 'Ready' },
      { id: 'bucket-c', name: 'Doing' },
    ] }));
  });
  const unknown = await cli.run(['config', 'set', 'bucket', 'Missing']);
  assert.equal(unknown.code, 2);
  assert.match(unknown.stderr, /Unknown Bucket 'Missing'.*bucket-a.*bucket-b.*bucket-c/);
  const ambiguous = await cli.run(['config', 'set', 'bucket', 'Ready']);
  assert.equal(ambiguous.code, 2);
  assert.match(ambiguous.stderr, /Ambiguous Bucket 'Ready'.*bucket-a.*bucket-b/);
});

it('stops --all at the hard page cap and leaves a cursor for remaining pages', async () => {
  let page = 0;
  servers.graph.setReply((_request, response) => {
    page += 1;
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({
      value: [{ id: `plan-${page}`, title: `Plan ${page}` }],
      '@odata.nextLink': `${servers.graph.url}/me/planner/plans?page=${page + 1}`,
    }));
  });
  servers.graph.reset();

  const result = await cli.run(['plans', '--all']);
  assert.equal(result.code, 0, result.stderr);
  assert.equal(servers.graph.requests.length, 100);
  assert.match(result.stdout, /Plan 100/);
  assert.match(result.stderr, /more available: --cursor .* \(or --all\)/);
});

it('prints the current user compactly and preserves the raw Graph response with --json', async () => {
  servers.graph.setReply((request, response) => {
    assert.equal(request.method, 'GET');
    assert.equal(request.path, '/me');
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ id: 'user-1', displayName: 'Ada Planner', mail: 'ada@example.test' }));
  });

  const compact = await cli.run(['me']);
  assert.equal(compact.code, 0, compact.stderr);
  assert.match(compact.stdout, /Ada Planner/);
  assert.match(compact.stdout, /ada@example\.test/);
  const json = await cli.run(['me', '--json']);
  assert.equal(json.code, 0, json.stderr);
  assert.deepEqual(JSON.parse(json.stdout), { id: 'user-1', displayName: 'Ada Planner', mail: 'ada@example.test' });
});

it('requires an actionable Default Plan for Buckets and members', async () => {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: { accessToken: 'read-token' } } }));
  for (const command of ['buckets', 'members']) {
    const result = await cli.run([command]);
    assert.equal(result.code, 2);
    assert.match(result.stderr, /Default Plan.*--plan|--plan.*Default Plan/);
    assert.equal(result.stdout, '');
  }
});

it('lists Buckets from the selected Default Plan', async () => {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: { accessToken: 'read-token', defaultPlan: '00000000-0000-4000-8000-000000000001' } } }));
  servers.graph.setReply((request, response) => {
    assert.equal(request.method, 'GET');
    assert.equal(request.path, '/planner/plans/00000000-0000-4000-8000-000000000001/buckets');
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ value: [{ id: 'bucket-1', name: 'Ready' }] }));
  });
  servers.graph.reset();

  const result = await cli.run(['buckets']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /ID\s+NAME/);
  assert.match(result.stdout, /bucket-1\s+Ready/);
  assert.equal(servers.graph.requests.length, 1);
});

it('resolves a Default Plan name to its id before listing Buckets', async () => {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: { accessToken: 'read-token', defaultPlan: 'Roadmap' } } }));
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/me/planner/plans') {
      response.end(JSON.stringify({ value: [{ id: '00000000-0000-4000-8000-000000000001', title: 'Roadmap' }] }));
    } else if (request.path === '/planner/plans/00000000-0000-4000-8000-000000000001/buckets') {
      response.end(JSON.stringify({ value: [{ id: 'bucket-1', name: 'Ready' }] }));
    } else response.end(JSON.stringify({ value: [] }));
  });
  servers.graph.reset();

  const result = await cli.run(['buckets']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /bucket-1\s+Ready/);
  assert.deepEqual(servers.graph.requests.map(({ path }) => path), ['/me/planner/plans', '/planner/plans/00000000-0000-4000-8000-000000000001/buckets']);
});

it('lists the selected Plan group members', async () => {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: { accessToken: 'read-token', defaultPlan: '00000000-0000-4000-8000-000000000001' } } }));
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/planner/plans/00000000-0000-4000-8000-000000000001') {
      response.end(JSON.stringify({ id: '00000000-0000-4000-8000-000000000001', title: 'Roadmap', owner: 'group-1' }));
    } else if (request.path === '/groups/group-1/members') {
      response.end(JSON.stringify({ value: [{ id: 'user-1', displayName: 'Ada Planner', mail: 'ada@example.test' }] }));
    } else response.end(JSON.stringify({ value: [] }));
  });
  servers.graph.reset();

  const result = await cli.run(['members']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /ID\s+NAME\s+EMAIL/);
  assert.match(result.stdout, /user-1\s+Ada Planner\s+ada@example\.test/);
  assert.deepEqual(servers.graph.requests.map(({ method, path }) => [method, path]), [
    ['GET', '/planner/plans/00000000-0000-4000-8000-000000000001'],
    ['GET', '/groups/group-1/members'],
  ]);
});

it('prints raw Graph Bucket and member responses with --json', async () => {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: {
    accessToken: 'read-token', defaultPlan: '00000000-0000-4000-8000-000000000001',
  } } }));
  const buckets = { value: [{ id: 'bucket-1', name: 'Ready', planId: 'plan-id' }] };
  const members = { value: [{ id: 'member-1', displayName: 'Ada Planner', mail: 'ada@example.test' }] };
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path.endsWith('/buckets')) response.end(JSON.stringify(buckets));
    else if (request.path === '/planner/plans/00000000-0000-4000-8000-000000000001') response.end(JSON.stringify({ owner: 'group-1' }));
    else response.end(JSON.stringify(members));
  });

  const bucketOutput = await cli.run(['buckets', '--json']);
  assert.equal(bucketOutput.code, 0, bucketOutput.stderr);
  assert.deepEqual(JSON.parse(bucketOutput.stdout), buckets);
  const memberOutput = await cli.run(['members', '--json']);
  assert.equal(memberOutput.code, 0, memberOutput.stderr);
  assert.deepEqual(JSON.parse(memberOutput.stdout), members);
});
