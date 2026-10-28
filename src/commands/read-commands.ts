import { graphRequest, type Context } from './auth-commands.js';
import { CliError, usage } from '../utils/errors.js';
import { isGuid, isPersonalPlan, personalPlanMember, resolvePlanId } from '../utils/name-resolution.js';

interface Plan {
  id?: string;
  title?: string;
  name?: string;
  owner?: string;
  container?: { url?: string };
}

interface Bucket {
  id?: string;
  name?: string;
}

interface Member {
  id?: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

interface Collection<T> {
  value?: T[];
  '@odata.nextLink'?: string;
  [key: string]: unknown;
}

interface ReadCollection<T> {
  response: Collection<T>;
  items: T[];
  nextLink?: string;
}

const ALL_PAGE_CAP = 100;

function table(headers: string[], rows: string[][]): string {
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values: string[]) => values.map((value, index) => value.padEnd(widths[index]!)).join('  ').trimEnd();
  return `${line(headers)}\n${rows.map(line).join('\n')}${rows.length ? '\n' : ''}`;
}

export async function runReadCommand(ctx: Context, command: string, args: string[]): Promise<void> {
  const requestedPlan = command === 'buckets' || command === 'members'
    ? option(args, '--plan') ?? ctx.options.plan ?? ctx.env.PLANNER_PLAN ?? ctx.profile.defaultPlan
    : undefined;
  if (command === 'buckets' || command === 'members') {
    if (!requestedPlan) usage("A Default Plan is required. Set one with 'planner config set plan <name|id>' or pass --plan <name|id>.");
  }
  const selectedPlan = requestedPlan && !isGuid(requestedPlan) ? await resolvePlanId(ctx, requestedPlan) : requestedPlan;
  let response: unknown;
  let rows: unknown[] | undefined;
  let nextLink: string | undefined;
  if (command === 'me') response = await graphRequest(ctx, 'GET', '/me');
  else if (command === 'buckets') {
    const result = await readCollection<Bucket>(ctx, `/planner/plans/${encodeURIComponent(selectedPlan!)}/buckets`, args);
    response = result.response;
    rows = result.items;
    nextLink = result.nextLink;
  }
  else if (command === 'members') {
    const plan = await graphRequest(ctx, 'GET', `/planner/plans/${encodeURIComponent(selectedPlan!)}`) as Plan;
    if (!plan.owner) usage(`Could not find the owning group for Plan '${selectedPlan}'. Check the Plan id.`);
    const owner = plan.owner;
    const ownerOnly = async (): Promise<ReadCollection<Member>> => {
      const member = await personalPlanMember(ctx, owner);
      return { response: { value: [member] }, items: [member] };
    };
    let result: ReadCollection<Member>;
    if (isPersonalPlan(plan)) result = await ownerOnly();
    else {
      try { result = await readCollection<Member>(ctx, `/groups/${encodeURIComponent(owner)}/members`, args); }
      catch (error) {
        if (!(error instanceof CliError && error.status === 404)) throw error;
        result = await ownerOnly();
      }
    }
    response = result.response;
    rows = result.items;
    nextLink = result.nextLink;
  } else if (command === 'plans') {
    const result = await readCollection<Plan>(ctx, '/me/planner/plans', args);
    response = result.response;
    rows = result.items;
    nextLink = result.nextLink;
  } else usage(`Unknown read command: ${command}`);
  if (args.includes('--json')) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}\n`);
    printContinuation(nextLink);
    return;
  }
  if (command === 'me') {
    const user = response as { displayName?: string; userPrincipalName?: string; mail?: string };
    process.stdout.write(table(['NAME', 'EMAIL'], [[user.displayName ?? 'Current user', user.mail ?? user.userPrincipalName ?? '']]));
    return;
  }
  if (command === 'members') {
    const members = (rows as Member[] | undefined) ?? (response as Collection<Member>).value ?? [];
    process.stdout.write(table(['ID', 'NAME', 'EMAIL'], members.map((member) => [member.id ?? '', member.displayName ?? '', member.mail ?? member.userPrincipalName ?? ''])));
    printContinuation(nextLink);
    return;
  }
  if (command === 'buckets') {
    const buckets = (rows as Bucket[] | undefined) ?? (response as Collection<Bucket>).value ?? [];
    process.stdout.write(table(['ID', 'NAME'], buckets.map((bucket) => [bucket.id ?? '', bucket.name ?? ''])));
    printContinuation(nextLink);
    return;
  }
  if (command !== 'plans') throw new Error(`Unknown read command: ${command}`);
  const plans = (rows as Plan[] | undefined) ?? (response as Collection<Plan>).value ?? [];
  process.stdout.write(table(['ID', 'NAME'], plans.map((plan) => [plan.id ?? '', plan.title ?? plan.name ?? ''])));
  printContinuation(nextLink);
}

async function readCollection<T>(ctx: Context, path: string, args: string[]): Promise<ReadCollection<T>> {
  const token = option(args, '--cursor');
  let url = token ? decodeCursor(token, ctx) : path;
  const all = args.includes('--all');
  const pages: Collection<T>[] = [];
  let nextLink: string | undefined;
  for (let page = 0; page < (all ? ALL_PAGE_CAP : 1); page += 1) {
    const result = await graphPage(ctx, url) as Collection<T>;
    pages.push(result);
    nextLink = result['@odata.nextLink'];
    if (!all || !nextLink) break;
    validateContinuation(nextLink, ctx);
    url = nextLink;
  }
  const first = pages[0] ?? { value: [] };
  const items = pages.flatMap((page) => page.value ?? []);
  const response = all ? { ...first, value: items, ...(nextLink ? { '@odata.nextLink': nextLink } : {}) } : first;
  if (all && !nextLink) delete response['@odata.nextLink'];
  return { response, items, nextLink };
}

function graphPage(ctx: Context, url: string): Promise<unknown> {
  return /^https?:\/\//i.test(url)
    ? graphRequest(ctx, 'GET', url, { rawUrl: url })
    : graphRequest(ctx, 'GET', url);
}

function graphOrigin(ctx: Context): string {
  return new URL(ctx.env.PLANNER_GRAPH_URL || 'https://graph.microsoft.com/v1.0').origin;
}

function validateContinuation(link: string, ctx: Context): void {
  let url: URL;
  try { url = new URL(link, ctx.env.PLANNER_GRAPH_URL || 'https://graph.microsoft.com/v1.0'); }
  catch { usage('Graph returned an invalid continuation URL.'); }
  const basePath = new URL(ctx.env.PLANNER_GRAPH_URL || 'https://graph.microsoft.com/v1.0').pathname.replace(/\/$/, '');
  const inBasePath = !basePath || url.pathname === basePath || url.pathname.startsWith(`${basePath}/`);
  if (url.origin !== graphOrigin(ctx) || !inBasePath) usage('Graph returned a continuation URL outside the configured Graph URL.');
}

function decodeCursor(token: string, ctx: Context): string {
  let link: string;
  try { link = Buffer.from(token, 'base64url').toString('utf8'); }
  catch { usage('Invalid cursor. Copy the cursor from the previous command output.'); }
  validateContinuation(link, ctx);
  return link;
}

function printContinuation(nextLink: string | undefined): void {
  if (nextLink) process.stderr.write(`more available: --cursor ${Buffer.from(nextLink).toString('base64url')} (or --all)\n`);
}


function option(args: string[], name: string): string | undefined {
  const index = args.indexOf(name);
  if (index < 0) return undefined;
  const value = args[index + 1];
  if (!value || value.startsWith('-')) usage(`Option '${name}' requires a value.`);
  return value;
}
