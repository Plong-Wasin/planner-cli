import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http';

export interface FakeRequest {
  method: string;
  path: string;
  headers: IncomingMessage['headers'];
  body: string;
}

export type FakeReply = (request: FakeRequest, response: ServerResponse) => void | Promise<void>;

export interface FakeEndpoint {
  url: string;
  requests: FakeRequest[];
  setReply(reply: FakeReply): void;
  reset(): void;
  close(): Promise<void>;
}

export interface FakeServers {
  graph: FakeEndpoint;
  login: FakeEndpoint;
  close(): Promise<void>;
}

async function startEndpoint(): Promise<FakeEndpoint> {
  const requests: FakeRequest[] = [];
  let reply: FakeReply = (_request, response) => {
    response.writeHead(404, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ error: 'Not found' }));
  };
  const server: Server = createServer(async (request, response) => {
    let body = '';
    for await (const chunk of request) body += chunk.toString();
    const recorded = { method: request.method ?? 'GET', path: request.url ?? '/', headers: request.headers, body };
    requests.push(recorded);
    await reply(recorded, response);
  });

  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  if (!address || typeof address === 'string') throw new Error('Fake server did not bind to a TCP port');

  return {
    url: `http://127.0.0.1:${address.port}`,
    requests,
    setReply: (nextReply) => { reply = nextReply; },
    reset: () => { requests.length = 0; },
    close: () => new Promise((resolve, reject) => {
      server.close((error) => error ? reject(error) : resolve());
      server.closeAllConnections();
    }),
  };
}

export async function startFakeServers(): Promise<FakeServers> {
  const graph = await startEndpoint();
  const login = await startEndpoint();
  return {
    graph,
    login,
    close: async () => {
      await Promise.all([graph.close(), login.close()]);
    },
  };
}
