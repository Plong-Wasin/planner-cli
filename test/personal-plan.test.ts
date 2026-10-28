import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import type { ServerResponse } from 'node:http';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

// A personal Plan has no owning group: Graph reports the owning user in `owner`
// and in `container.url`, and `/groups/<user id>` does not exist.
const planId = 'Pl4nAAAAAAAAAAAAAAAAAAAAAAAA';
const taskId = '4Q81bX_cmkWVWe3vseex5ckAIV4d';
const ownerId = '00000000-0000-4000-8000-000000000009';

let servers: FakeServers;
let cli: Cli;
let withContainer = true;

function json(response: ServerResponse, body: unknown, status = 200): void {
  response.writeHead(status, { 'Content-Type': 'application/json' });
  response.end(JSON.stringify(body));
}

before(async () => {
  servers = await startFakeServers();
  cli = createCli({ PLANNER_GRAPH_URL: servers.graph.url });
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: { accessToken: 'personal-token', defaultPlan: planId } } }));
  servers.graph.setReply((request, response) => {
    const path = request.path.split('?')[0] ?? '';
    if (path === `/planner/plans/${planId}`) {
      json(response, { id: planId, title: 'Mine', owner: ownerId, ...(withContainer ? { container: { type: 'unknownFutureValue', url: `${servers.graph.url}/users/${ownerId}` } } : {}) });
    } else if (path === '/me/planner/plans') json(response, { value: [{ id: planId, title: 'Mine', owner: ownerId }] });
    else if (path === `/planner/plans/${planId}/tasks`) json(response, { value: [{ id: taskId, title: 'Owned', percentComplete: 0, assignments: { [ownerId]: {} } }] });
    else if (path === `/planner/tasks/${taskId}`) json(response, { id: taskId, planId, title: 'Owned', percentComplete: 0, assignments: { [ownerId]: {} } });
    else if (path === `/users/${ownerId}`) json(response, { id: ownerId, displayName: 'Plan Owner', mail: 'owner@example.test' });
    else if (path.startsWith('/groups/')) json(response, { error: { code: 'Request_ResourceNotFound', message: `Resource '${path.split('/')[2]}' does not exist` } }, 404);
    else json(response, { value: [] });
  });
});

after(async () => {
  await servers.close();
  cli.cleanup();
});

it('lists tasks of a personal Plan and names the owner as Assignee', async () => {
  const result = await cli.run(['tasks']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /Plan Owner/);
});

it('shows a task of a personal Plan with the owner as Assignee', async () => {
  const result = await cli.run(['task', taskId]);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /assignees: Plan Owner/);
});

it('lists the owner as the only member of a personal Plan', async () => {
  const result = await cli.run(['members']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /Plan Owner/);
  assert.match(result.stdout, /owner@example\.test/);
});

it('falls back to the owner when a Plan carries no container and the group lookup is a 404', async () => {
  withContainer = false;
  try {
    const result = await cli.run(['members']);
    assert.equal(result.code, 0, result.stderr);
    assert.match(result.stdout, /Plan Owner/);
  } finally {
    withContainer = true;
  }
});

it('still reports a missing group as an error when the owner is not a user either', async () => {
  servers.graph.setReply((request, response) => {
    const path = request.path.split('?')[0] ?? '';
    if (path === `/planner/plans/${planId}`) json(response, { id: planId, owner: ownerId });
    else json(response, { error: { code: 'Request_ResourceNotFound', message: 'gone' } }, 404);
  });
  const result = await cli.run(['members']);
  assert.equal(result.code, 1);
  assert.match(result.stderr, /Not found/);
});

it('resolves an Assignee by name on a personal Plan', async () => {
  servers.graph.setReply((request, response) => {
    const path = request.path.split('?')[0] ?? '';
    if (path === `/planner/plans/${planId}`) json(response, { id: planId, owner: ownerId, container: { url: `${servers.graph.url}/users/${ownerId}` } });
    else if (path === `/planner/plans/${planId}/buckets`) json(response, { value: [{ id: 'BuckBBBBBBBBBBBBBBBBBBBBBBBB', name: 'To do' }] });
    else if (path === `/users/${ownerId}`) json(response, { id: ownerId, displayName: 'Plan Owner', mail: 'owner@example.test' });
    else if (path === '/planner/tasks' && request.method === 'POST') json(response, { id: taskId }, 201);
    else json(response, { value: [] });
  });
  servers.graph.reset();
  const result = await cli.run(['add', 'Named', '--bucket', 'To do', '--assign', 'Plan Owner']);
  assert.equal(result.code, 0, result.stderr);
  const post = servers.graph.requests.find((request) => request.method === 'POST' && request.path === '/planner/tasks');
  assert.ok(post, 'task was created');
  assert.deepEqual(Object.keys(JSON.parse(post.body).assignments), [ownerId]);
});

it('does not mask a failed --cursor continuation as an owner-only member list', async () => {
  servers.graph.setReply((request, response) => {
    const path = request.path.split('?')[0] ?? '';
    if (path === `/planner/plans/${planId}`) json(response, { id: planId, owner: ownerId });
    else if (path === '/me/planner/plans') json(response, { value: [{ id: planId, title: 'Mine', owner: ownerId }] });
    else if (path === `/users/${ownerId}`) json(response, { id: ownerId, displayName: 'Plan Owner' });
    else json(response, { error: { code: 'Request_ResourceNotFound', message: 'expired' } }, 404);
  });
  const cursor = Buffer.from(`${servers.graph.url}/groups/some-group/members?page=2`).toString('base64url');
  const result = await cli.run(['members', '--cursor', cursor]);
  assert.equal(result.code, 1, result.stderr);
  assert.match(result.stderr, /Not found: GET .*\/groups\/some-group/);
  assert.doesNotMatch(result.stdout, /Plan Owner/);
});
