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
