import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;

const PLAN_A = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAA';
const PLAN_B = 'BBBBBBBBBBBBBBBBBBBBBBBBBBBB';

before(async () => {
  servers = await startFakeServers();
  cli = createCli({ PLANNER_GRAPH_URL: servers.graph.url });
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: { accessToken: 'search-token', defaultPlan: PLAN_A } } }));
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    const path = request.path;
    // Like Graph, send Details only when the request expands them.
    const details = <T>(value: T): T | undefined => (path.includes('expand=details') ? value : undefined);
    if (path === '/me/planner/plans') {
      response.end(JSON.stringify({ value: [
        { id: PLAN_A, title: 'Dev', owner: 'group-a' },
        { id: PLAN_B, title: 'Omni Channel', owner: 'group-b' },
      ] }));
    } else if (path.startsWith(`/planner/plans/${PLAN_A}/tasks`)) {
      response.end(JSON.stringify({ value: [
        { id: 'task-a1', title: 'Fix login', planId: PLAN_A, bucketId: 'bucket-a', percentComplete: 0, assignments: {},
          details: details({ description: 'nothing here' }) },
      ] }));
    } else if (path.startsWith(`/planner/plans/${PLAN_B}/tasks`)) {
      response.end(JSON.stringify({ value: [
        { id: 'task-b1', title: 'การจัดการป้ายสินค้า', planId: PLAN_B, percentComplete: 0, assignments: { 'user-1': {} },
          details: details({ description: 'Super Deal badge' }) },
        { id: 'task-b2', title: 'Other work', planId: PLAN_B, percentComplete: 0, assignments: {},
          details: details({ description: 'unrelated' }) },
      ] }));
    } else if (path === `/planner/plans/${PLAN_B}/buckets`) {
      response.end(JSON.stringify({ value: [] }));
    } else if (path === '/groups/group-b/members') {
      response.end(JSON.stringify({ value: [{ id: 'user-1', displayName: 'Ada Planner' }] }));
    } else response.end(JSON.stringify({ value: [] }));
  });
});

after(async () => {
  await servers.close();
  cli.cleanup();
});

it('--search keeps only Tasks whose title contains the text, ignoring case', async () => {
  servers.graph.reset();
  const result = await cli.run(['tasks', '--search', 'FIX LOGIN']);

  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /task-a1\s+Fix login/);
  assert.doesNotMatch(result.stdout, /task-b/);
});

it('--search also matches the description when --details loads it', async () => {
  servers.graph.reset();
  const withoutDetails = await cli.run(['tasks', '--all-plans', '--search', 'super deal']);
  assert.equal(withoutDetails.code, 0, withoutDetails.stderr);
  assert.doesNotMatch(withoutDetails.stdout, /task-b1/);

  const withDetails = await cli.run(['tasks', '--all-plans', '--search', 'super deal', '--details']);
  assert.equal(withDetails.code, 0, withDetails.stderr);
  assert.match(withDetails.stdout, /task-b1/);
  assert.doesNotMatch(withDetails.stdout, /task-b2|task-a1/);
});

it('--all-plans lists Tasks from every Plan with a PLAN column and resolves Assignee names per Plan', async () => {
  servers.graph.reset();
  const result = await cli.run(['tasks', '--all-plans', '--search', 'ป้าย']);

  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /PLAN\s+ID\s+TITLE\s+BUCKET\s+PERCENT\s+DUE\s+ASSIGNEES/);
  assert.match(result.stdout, /Omni Channel\s+task-b1\s+การจัดการป้ายสินค้า\s+0%\s+Ada Planner/);
  assert.ok(servers.graph.requests.some(({ path }) => path.startsWith(`/planner/plans/${PLAN_A}/tasks`)));
  assert.ok(servers.graph.requests.some(({ path }) => path.startsWith(`/planner/plans/${PLAN_B}/tasks`)));
});

it('--all-plans --json prints one merged list that carries planId', async () => {
  const result = await cli.run(['tasks', '--all-plans', '--json']);

  assert.equal(result.code, 0, result.stderr);
  const parsed = JSON.parse(result.stdout) as { value: { id: string; planId: string }[] };
  assert.deepEqual(parsed.value.map((task) => `${task.planId}:${task.id}`).sort(), [`${PLAN_A}:task-a1`, `${PLAN_B}:task-b1`, `${PLAN_B}:task-b2`].sort());
});

it('--all-plans cannot be combined with --plan, --bucket or --cursor', async () => {
  for (const extra of [['--plan', 'Dev'], ['--bucket', 'Todo'], ['--cursor', 'abc']]) {
    const result = await cli.run(['tasks', '--all-plans', ...extra]);
    assert.equal(result.code, 2, `${extra[0]}: ${result.stderr}`);
    assert.match(result.stderr, /--all-plans/);
  }
});
