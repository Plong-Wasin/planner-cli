import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;

before(async () => {
  servers = await startFakeServers();
  cli = createCli({ PLANNER_GRAPH_URL: servers.graph.url, PLANNER_TOKEN: 'test-token' });
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ value: [{ id: 'message-1' }] }));
  });
});

after(async () => {
  await servers.close();
  cli.cleanup();
});

it('sends raw GET query pairs and prints the Graph response', async () => {
  const result = await cli.run(['api', 'GET', '/me/messages', '$select=id,subject', 'top=5']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, '{"value":[{"id":"message-1"}]}\n');
  assert.equal(result.stderr, '');
  assert.equal(servers.graph.requests[0]?.method, 'GET');
  assert.equal(servers.graph.requests[0]?.path, '/me/messages?%24select=id%2Csubject&top=5');
  assert.equal(servers.graph.requests[0]?.headers.authorization, 'Bearer test-token');
});

it('sends raw POST JSON supplied as an argument', async () => {
  servers.graph.reset();
  servers.graph.setReply((request, response) => {
    response.writeHead(201, { 'Content-Type': 'application/json' });
    response.end(request.body);
  });
  const result = await cli.run(['api', 'POST', '/planner/tasks', '--data', '{"planId":"plan-1","title":"Review"}']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, '{"planId":"plan-1","title":"Review"}\n');
  assert.equal(servers.graph.requests[0]?.method, 'POST');
  assert.equal(servers.graph.requests[0]?.path, '/planner/tasks');
  assert.equal(servers.graph.requests[0]?.headers['content-type'], 'application/json');
  assert.deepEqual(JSON.parse(servers.graph.requests[0]?.body ?? ''), { planId: 'plan-1', title: 'Review' });
});

it('loads raw POST JSON from a file or stdin', async () => {
  servers.graph.reset();
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(request.body);
  });
  const filePath = join(cli.home, 'payload.json');
  writeFileSync(filePath, '{"title":"From file"}');
  const fromFile = await cli.run(['api', 'POST', '/planner/tasks', '--data', `@${filePath}`]);
  assert.equal(fromFile.code, 0, fromFile.stderr);
  assert.deepEqual(JSON.parse(servers.graph.requests[0]?.body ?? ''), { title: 'From file' });

  servers.graph.reset();
  const fromStdin = await cli.run(['api', 'POST', '/planner/tasks', '--data', '-'], {}, '{"title":"From stdin"}');
  assert.equal(fromStdin.code, 0, fromStdin.stderr);
  assert.deepEqual(JSON.parse(servers.graph.requests[0]?.body ?? ''), { title: 'From stdin' });
});

it('follows Graph continuation links with --all and combines the pages', async () => {
  servers.graph.reset();
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify(request.path === '/me/messages?%24top=1'
      ? { value: [{ id: 'first' }], '@odata.nextLink': `${servers.graph.url}/me/messages?$skiptoken=next` }
      : { value: [{ id: 'second' }] }));
  });
  const result = await cli.run(['api', 'GET', '/me/messages', '$top=1', '--all']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, '{"value":[{"id":"first"},{"id":"second"}]}\n');
  assert.equal(servers.graph.requests.length, 2);
  assert.equal(servers.graph.requests[0]?.path, '/me/messages?%24top=1');
  assert.equal(servers.graph.requests[1]?.path, '/me/messages?$skiptoken=next');
});

it('stops --all at the 50-page cap', async () => {
  servers.graph.reset();
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({
      value: [{ id: String(servers.graph.requests.length) }],
      '@odata.nextLink': `${servers.graph.url}/me/messages?$skiptoken=repeat`,
    }));
  });
  const result = await cli.run(['api', 'GET', '/me/messages', '--all']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(servers.graph.requests.length, 50);
  assert.equal(JSON.parse(result.stdout).value.length, 50);
});

it('treats an empty 204 as success and maps Graph errors with the request shape', async () => {
  servers.graph.reset();
  servers.graph.setReply((_request, response) => { response.writeHead(204).end(); });
  const empty = await cli.run(['api', 'DELETE', '/planner/tasks/task-1']);
  assert.equal(empty.code, 0, empty.stderr);
  assert.equal(empty.stdout, '');
  assert.equal(empty.stderr, '');

  servers.graph.reset();
  servers.graph.setReply((_request, response) => {
    response.writeHead(404, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ error: 'Not found' }));
  });
  const missing = await cli.run(['api', 'GET', '/planner/tasks/missing']);
  assert.equal(missing.code, 1);
  assert.match(missing.stderr, /planner: Not found: GET \/planner\/tasks\/missing/);
});
