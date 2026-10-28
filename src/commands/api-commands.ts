import { readFile } from 'node:fs/promises';
import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';

export async function runApi(ctx: Context, args: string[]): Promise<void> {
  const method = args[0]?.toUpperCase();
  const path = args[1];
  if (!method || !path) usage("Use 'planner api <METHOD> <path> [key=value ...]'.");
  if (!['GET', 'POST', 'PUT', 'PATCH', 'DELETE'].includes(method)) usage(`Unsupported HTTP method: ${args[0]}.`);
  const query = new URLSearchParams();
  let body: string | undefined;
  let all = false;
  const headers: Record<string, string> = {};
  for (let index = 2; index < args.length; index += 1) {
    const pair = args[index]!;
    if (pair === '--all') {
      all = true;
      continue;
    }
    if (pair === '--header' || pair.startsWith('--header=')) {
      const header = pair === '--header' ? args[++index] : pair.slice('--header='.length);
      const colon = header?.indexOf(':') ?? -1;
      if (!header || colon <= 0) usage("Option '--header' requires 'Name: value'.");
      headers[header.slice(0, colon).trim()] = header.slice(colon + 1).trim();
      continue;
    }
    if (pair === '--data' || pair.startsWith('--data=')) {
      if (body !== undefined) usage("Option '--data' may only be provided once.");
      const source = pair === '--data' ? args[++index] : pair.slice('--data='.length);
      if (source === undefined) usage("Option '--data' requires JSON, @file, or '-'.");
      const text = source === '-' ? await readStdin() : source.startsWith('@') ? await readFile(source.slice(1), 'utf8') : source;
      try {
        body = JSON.stringify(JSON.parse(text))!;
      } catch {
        usage("Option '--data' must contain valid JSON.");
      }
      continue;
    }
    const separator = pair.indexOf('=');
    if (separator <= 0) usage(`Expected a query pair in key=value form: ${pair}.`);
    query.append(pair.slice(0, separator), pair.slice(separator + 1));
  }
  if (all && method !== 'GET') usage("Option '--all' is only supported with GET requests.");
  const hasIfMatch = Object.keys(headers).some((name) => name.toLowerCase() === 'if-match');
  if ((method === 'PATCH' || method === 'DELETE') && path.startsWith('/planner/') && !hasIfMatch) {
    // Planner rejects writes without an If-Match etag, so fetch the current one.
    const current = await graphRequest(ctx, 'GET', path) as { '@odata.etag'?: string } | undefined;
    if (current?.['@odata.etag']) headers['If-Match'] = current['@odata.etag'];
  }
  let data = await graphRequest(ctx, method, path, { query, body, headers });
  if (all) data = await collectPages(ctx, path, query, data);
  if (data !== undefined) process.stdout.write(`${JSON.stringify(data)}\n`);
}

async function collectPages(ctx: Context, path: string, query: URLSearchParams, initial: unknown): Promise<unknown> {
  if (!isRecord(initial) || !Array.isArray(initial.value)) return initial;
  const values = [...initial.value];
  const result: Record<string, unknown> = { ...initial, value: values };
  delete result['@odata.nextLink'];
  let nextLink = initial['@odata.nextLink'];
  for (let page = 1; typeof nextLink === 'string' && page < 50; page += 1) {
    const nextUrl = new URL(nextLink);
    const base = new URL(ctx.env.PLANNER_GRAPH_URL || 'https://graph.microsoft.com/v1.0');
    const basePath = base.pathname.replace(/\/$/, '');
    const inBasePath = !basePath || nextUrl.pathname === basePath || nextUrl.pathname.startsWith(`${basePath}/`);
    if (nextUrl.origin !== base.origin || !inBasePath) {
      usage('Graph returned a continuation URL outside the configured Graph URL.');
    }
    const data: unknown = await graphRequest(ctx, 'GET', path, { rawUrl: nextUrl.href });
    if (!isRecord(data)) break;
    if (Array.isArray(data.value)) values.push(...data.value);
    nextLink = data['@odata.nextLink'];
  }
  return result;
}

function isRecord(value: unknown): value is Record<string, unknown> & { value: unknown[] } {
  return typeof value === 'object' && value !== null && Array.isArray((value as { value?: unknown }).value);
}

async function readStdin(): Promise<string> {
  const chunks: Buffer[] = [];
  for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString('utf8');
}
