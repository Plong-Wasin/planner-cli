import { graphRequest, type Context } from '../commands/auth-commands.js';
import { usage } from './errors.js';

interface GraphCollection<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

interface Plan {
  id?: string;
  title?: string;
  name?: string;
  owner?: string;
}

interface Bucket {
  id?: string;
  name?: string;
}

interface Assignee {
  id?: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

const PAGE_CAP = 100;

export function isGuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}

export async function resolvePlanId(ctx: Context, value: string): Promise<string> {
  if (isGuid(value)) return value;
  const plans = await collection<Plan>(ctx, '/me/planner/plans');
  const matches = plans.filter((plan) => plan.id === value || equal(plan.title ?? plan.name, value));
  if (matches.length === 1 && matches[0]?.id) return matches[0].id;
  const choices = (matches.length ? matches : plans).map((plan) => `${plan.title ?? plan.name ?? 'Untitled'} (${plan.id ?? 'unknown id'})`);
  fail('Plan', value, matches.length > 1, choices);
}

export async function resolveBucketId(ctx: Context, plan: string, value: string): Promise<string> {
  const planId = await resolvePlanId(ctx, plan);
  if (isGuid(value)) return value;
  const buckets = await collection<Bucket>(ctx, `/planner/plans/${encodeURIComponent(planId)}/buckets`);
  const matches = buckets.filter((bucket) => bucket.id === value || equal(bucket.name, value));
  if (matches.length === 1 && matches[0]?.id) return matches[0].id;
  fail('Bucket', value, matches.length > 1, (matches.length ? matches : buckets).map((bucket) => `${bucket.name ?? 'Untitled'} (${bucket.id ?? 'unknown id'})`));
}

export async function resolveAssigneeId(ctx: Context, plan: string, value: string): Promise<string> {
  if (equal(value, 'me')) {
    const me = await graphRequest(ctx, 'GET', '/me') as Assignee;
    if (me.id) return me.id;
    usage("Could not identify the current user. Check the signed-in Profile.");
  }
  if (isGuid(value)) return value;
  const planId = await resolvePlanId(ctx, plan);
  const selectedPlan = await graphRequest(ctx, 'GET', `/planner/plans/${encodeURIComponent(planId)}`) as Plan;
  if (!selectedPlan.owner) usage(`Could not find the owning group for Plan '${plan}'. Check the Plan id.`);
  const members = await collection<Assignee>(ctx, `/groups/${encodeURIComponent(selectedPlan.owner)}/members`);
  const matches = members.filter((member) => member.id === value || equal(member.displayName, value)
    || equal(member.mail, value) || equal(member.userPrincipalName, value));
  if (matches.length === 1 && matches[0]?.id) return matches[0].id;
  fail('Assignee', value, matches.length > 1, (matches.length ? matches : members).map((member) => `${member.displayName ?? 'Unknown'} <${member.mail ?? member.userPrincipalName ?? 'no email'}> (${member.id ?? 'unknown id'})`));
}

async function collection<T>(ctx: Context, path: string): Promise<T[]> {
  const values: T[] = [];
  let next = path;
  for (let page = 0; page < PAGE_CAP; page += 1) {
    const response = await graphPage(ctx, next) as GraphCollection<T>;
    values.push(...(response.value ?? []));
    if (!response['@odata.nextLink']) return values;
    next = response['@odata.nextLink'];
    validateNext(next, ctx);
  }
  return values;
}

function graphPage(ctx: Context, url: string): Promise<unknown> {
  return /^https?:\/\//i.test(url)
    ? graphRequest(ctx, 'GET', url, { rawUrl: url })
    : graphRequest(ctx, 'GET', url);
}

function validateNext(value: string, ctx: Context): void {
  let link: URL;
  try { link = new URL(value, ctx.env.PLANNER_GRAPH_URL || 'https://graph.microsoft.com/v1.0'); }
  catch { usage('Graph returned an invalid continuation URL while resolving a name.'); }
  const base = new URL(ctx.env.PLANNER_GRAPH_URL || 'https://graph.microsoft.com/v1.0');
  const basePath = base.pathname.replace(/\/$/, '');
  const inBasePath = !basePath || link.pathname === basePath || link.pathname.startsWith(`${basePath}/`);
  if (link.origin !== base.origin || !inBasePath) usage('Graph returned a continuation URL outside the configured Graph URL.');
}

function equal(left: string | undefined, right: string): boolean {
  return (left ?? '').trim().toLocaleLowerCase() === right.trim().toLocaleLowerCase();
}

function fail(subject: string, value: string, ambiguous: boolean, candidates: string[]): never {
  usage(`${ambiguous ? 'Ambiguous' : 'Unknown'} ${subject} '${value}'. Candidates: ${candidates.join(', ') || 'none'}.`);
}
