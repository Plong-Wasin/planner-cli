import { createServer, type Server } from 'node:http';

export interface FakeEndpoint {
  url: string;
  close(): Promise<void>;
}

export interface FakeServers {
  graph: FakeEndpoint;
  login: FakeEndpoint;
  close(): Promise<void>;
}

async function startEndpoint(): Promise<FakeEndpoint> {
  const server: Server = createServer((_request, response) => {
    response.writeHead(404, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ error: 'Not found' }));
  });

  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  if (!address || typeof address === 'string') throw new Error('Fake server did not bind to a TCP port');

  return {
    url: `http://127.0.0.1:${address.port}`,
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
