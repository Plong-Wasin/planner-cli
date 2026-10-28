import assert from 'node:assert/strict';
import { readFileSync, statSync, writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;

before(async () => {
  servers = await startFakeServers();
  cli = createCli({ PLANNER_LOGIN_URL: servers.login.url, PLANNER_GRAPH_URL: servers.graph.url });
  let polls = 0;
  servers.login.setReply(async (request, response) => {
    if (request.path.endsWith('/devicecode')) {
      response.writeHead(200, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify({ device_code: 'device-secret', user_code: 'ABCD-EFGH', verification_uri: 'https://login.example.test/device', interval: 0.01, expires_in: 30 }));
      return;
    }
    if (request.path.endsWith('/token')) {
      polls += 1;
      response.writeHead(200, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify(polls === 1
        ? { error: 'authorization_pending' }
        : { access_token: 'access-secret', refresh_token: 'refresh-secret', expires_in: 3600 }));
      return;
    }
    response.writeHead(404).end();
  });
});

after(async () => {
  await servers.close();
  cli.cleanup();
});

function setConfig(config: unknown): void {
  writeFileSync(`${cli.home}/config.json`, JSON.stringify(config), { mode: 0o600 });
}

function graphReply(status = 200): void {
  servers.graph.setReply((request, response) => {
    if (status !== 200) {
      response.writeHead(status, { 'Content-Type': 'application/json', ...(status === 429 ? { 'Retry-After': '2' } : {}) });
      response.end(JSON.stringify({ error: 'fake failure' }));
      return;
    }
    assert.equal(request.method, 'GET');
    assert.equal(request.path, '/me');
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ displayName: 'Planner User', userPrincipalName: 'user@example.test' }));
  });
  servers.graph.reset();
}

it('logs in with device code and stores credentials in a private config', async () => {
  servers.login.reset();
  const result = await cli.run(['auth', 'login', '--client-id', 'test-client']);
  assert.equal(result.code, 0);
  assert.match(result.stdout, /https:\/\/login\.example\.test\/device/);
  assert.match(result.stdout, /ABCD-EFGH/);
  assert.match(result.stdout, /Signed in/);
  assert.equal(result.stderr, '');

  assert.equal(servers.login.requests.length, 3);
  const deviceRequest = servers.login.requests[0]!;
  assert.equal(deviceRequest.method, 'POST');
  assert.equal(deviceRequest.path, '/common/oauth2/v2.0/devicecode');
  const deviceBody = new URLSearchParams(deviceRequest.body);
  assert.equal(deviceBody.get('client_id'), 'test-client');
  assert.match(deviceBody.get('scope') ?? '', /offline_access/);

  const tokenRequest = servers.login.requests[2]!;
  const tokenBody = new URLSearchParams(tokenRequest.body);
  assert.equal(tokenBody.get('grant_type'), 'urn:ietf:params:oauth:grant-type:device_code');
  assert.equal(tokenBody.get('device_code'), 'device-secret');

  const config = JSON.parse(readFileSync(`${cli.home}/config.json`, 'utf8')) as { profiles: Record<string, { refreshToken: string }> };
  assert.equal(config.profiles.default?.refreshToken, 'refresh-secret');
  assert.equal(statSync(`${cli.home}/config.json`).mode & 0o777, 0o600);
});

it('refreshes an expired Profile before checking identity and rotates its refresh token', async () => {
  setConfig({ defaultProfile: 'first', profiles: { first: { accessToken: 'expired', refreshToken: 'old-refresh', expiresAt: 1 } } });
  servers.login.reset();
  servers.login.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ access_token: 'fresh-access', refresh_token: 'new-refresh', expires_in: 3600 }));
  });
  graphReply();

  const result = await cli.run(['auth', 'check', '-p', 'first', '--client-id', 'custom-client'], { PLANNER_PROFILE: 'ignored' });
  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, 'Signed in as Planner User\n');
  assert.equal(servers.login.requests.length, 1);
  const refresh = new URLSearchParams(servers.login.requests[0]!.body);
  assert.equal(refresh.get('grant_type'), 'refresh_token');
  assert.equal(refresh.get('client_id'), 'custom-client');
  assert.equal(refresh.get('refresh_token'), 'old-refresh');
  assert.equal(servers.graph.requests[0]!.headers.authorization, 'Bearer fresh-access');
  const saved = JSON.parse(readFileSync(`${cli.home}/config.json`, 'utf8')) as { profiles: Record<string, { refreshToken: string }> };
  assert.equal(saved.profiles.first?.refreshToken, 'new-refresh');

  setConfig({ profiles: { second: { accessToken: 'expired', refreshToken: 'env-refresh', expiresAt: 1 } } });
  servers.login.reset();
  const viaEnvironment = await cli.run(['auth', 'check', '-p', 'second'], { PLANNER_CLIENT_ID: 'environment-client' });
  assert.equal(viaEnvironment.code, 0, viaEnvironment.stderr);
  assert.equal(new URLSearchParams(servers.login.requests[0]!.body).get('client_id'), 'environment-client');
});

it('allows PLANNER_TOKEN without a signed-in Profile and names method and path for API failures', async () => {
  graphReply();
  const ok = await cli.run(['auth', 'check'], { PLANNER_TOKEN: 'headless-token' });
  assert.equal(ok.code, 0, ok.stderr);
  assert.equal(servers.graph.requests[0]!.headers.authorization, 'Bearer headless-token');

  for (const [status, phrase] of [[401, 'Authentication failed'], [403, 'Access denied'], [404, 'Not found'], [429, 'Rate limited']] as const) {
    graphReply(status);
    const result = await cli.run(['auth', 'check'], { PLANNER_TOKEN: 'headless-token' });
    assert.equal(result.code, 1);
    assert.match(result.stderr, new RegExp(`${phrase}.*GET \/me`));
  }
});

it('reports failed refresh with an auth login hint', async () => {
  setConfig({ profiles: { default: { accessToken: 'expired', refreshToken: 'bad-refresh', expiresAt: 1 } } });
  servers.login.reset();
  servers.graph.reset();
  servers.login.setReply((_request, response) => {
    response.writeHead(400, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ error: 'invalid_grant' }));
  });
  const result = await cli.run(['auth', 'check']);
  assert.equal(result.code, 1);
  assert.match(result.stderr, /Run 'planner auth login'/);
  assert.equal(servers.graph.requests.length, 0);
});

it('sets and shows Profile defaults while masking credentials', async () => {
  setConfig({ defaultProfile: 'work', profiles: { work: { accessToken: 'visible-no', refreshToken: 'visible-no-too' } } });
  const plan = await cli.run(['config', 'set', 'plan', 'Roadmap Plan', '-p', 'work']);
  assert.equal(plan.code, 0);
  const bucket = await cli.run(['config', 'set', 'bucket', 'Ready', '-p', 'work']);
  assert.equal(bucket.code, 0);
  const shown = await cli.run(['config', 'show']);
  assert.equal(shown.code, 0);
  assert.match(shown.stdout, /Roadmap Plan/);
  assert.match(shown.stdout, /Ready/);
  assert.doesNotMatch(shown.stdout, /visible-no/);
  assert.equal((JSON.parse(shown.stdout) as { profiles: Record<string, { accessToken: string }> }).profiles.work?.accessToken, '********');
});

it('checks the selected Profile and logout removes only its credentials', async () => {
  setConfig({ defaultProfile: 'primary', profiles: { primary: { accessToken: 'primary-token', refreshToken: 'primary-refresh' }, backup: { accessToken: 'backup-token', refreshToken: 'backup-refresh' } } });
  graphReply();
  const fromEnvironment = await cli.run(['auth', 'check'], { PLANNER_PROFILE: 'backup' });
  assert.equal(fromEnvironment.code, 0, fromEnvironment.stderr);
  assert.equal(servers.graph.requests[0]!.headers.authorization, 'Bearer backup-token');

  graphReply();
  const checked = await cli.run(['auth', 'check', '-p', 'backup'], { PLANNER_PROFILE: 'primary' });
  assert.equal(checked.code, 0, checked.stderr);
  assert.equal(servers.graph.requests[0]!.headers.authorization, 'Bearer backup-token');

  const loggedOut = await cli.run(['auth', 'logout', '-p', 'backup']);
  assert.equal(loggedOut.code, 0);
  const saved = JSON.parse(readFileSync(`${cli.home}/config.json`, 'utf8')) as { profiles: Record<string, Record<string, unknown>> };
  assert.equal(saved.profiles.backup?.accessToken, undefined);
  assert.equal(saved.profiles.primary?.accessToken, 'primary-token');
});
