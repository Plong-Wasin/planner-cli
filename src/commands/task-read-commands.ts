import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';
import { isGuid, resolveBucketId, resolvePlanId } from '../utils/name-resolution.js';

interface Task {
  id?: string;
  title?: string;
  bucketId?: string;
  percentComplete?: number;
  dueDateTime?: string;
  createdDateTime?: string;
  completedDateTime?: string;
  assignments?: Record<string, unknown>;
  details?: { description?: string; checklist?: Record<string, { title?: string; isChecked?: boolean }> };
  [key: string]: unknown;
}

interface Collection<T> {
  value?: T[];
  '@odata.nextLink'?: string;
  [key: string]: unknown;
}

interface Plan { owner?: string }
interface Bucket { id?: string; name?: string }
interface Member { id?: string; displayName?: string; mail?: string; userPrincipalName?: string }

const DEFAULT_PAGE_SIZE = 50;
const PAGE_CAP = 100;
const TITLE_WIDTH = 48;
const FIELD_MAP: Record<string, string> = {
  id: 'id', title: 'title', bucket: 'bucketId', percent: 'percentComplete',
  due: 'dueDateTime', assignees: 'assignments', created: 'createdDateTime',
  completed: 'completedDateTime', priority: 'priority', start: 'startDateTime',
};
const DEFAULT_FIELDS = ['id', 'title', 'bucket', 'percent', 'due', 'assignees'];

export async function runTaskReadCommand(ctx: Context, command: 'tasks' | 'task', args: string[]): Promise<void> {
  if (command === 'task') return showTask(ctx, args);
  return listTasks(ctx, args);
}

async function listTasks(ctx: Context, args: string[]): Promise<void> {
  const planInput = option(args, '--plan') ?? ctx.options.plan ?? ctx.env.PLANNER_PLAN ?? ctx.profile.defaultPlan;
  const planId = planInput ? await resolvePlanId(ctx, planInput) : undefined;
  const bucketInput = option(args, '--bucket');
  const bucketId = bucketInput ? (isGuid(bucketInput) ? bucketInput : planId ? await resolveBucketId(ctx, planId, bucketInput) : usage("A Default Plan is required to resolve a Bucket name; pass --plan <name|id>.")) : undefined;
  const fieldsArg = option(args, '--fields');
  const fields = fieldsArg ? fieldsArg.split(',').map((field) => field.trim()).filter(Boolean) : DEFAULT_FIELDS;
  if (!fields.length || fields.some((field) => !Object.hasOwn(FIELD_MAP, field))) {
    usage(`Unknown field in --fields. Choose: ${Object.keys(FIELD_MAP).join(', ')}.`);
  }
  const selected = new Set(fields.map((field) => FIELD_MAP[field]!));
  if (bucketId) selected.add('bucketId');
  if (args.includes('--mine')) selected.add('assignments');
  if (args.includes('--open')) selected.add('percentComplete');
  if (option(args, '--created-after')) selected.add('createdDateTime');
  if (option(args, '--completed-after')) selected.add('completedDateTime');

  const query = new URLSearchParams();
  query.set('$top', String(DEFAULT_PAGE_SIZE));
  query.set('$select', [...selected].join(','));
  if (args.includes('--details')) query.set('$expand', 'details');
  const basePath = planId ? `/planner/plans/${encodeURIComponent(planId)}/tasks` : '/me/planner/tasks';
  const cursor = option(args, '--cursor');
  let url = cursor ? decodeCursor(cursor, ctx) : `${basePath}?${query}`;
  const all = args.includes('--all');
  const pages: Collection<Task>[] = [];
  let nextLink: string | undefined;
  for (let page = 0; page < (all ? PAGE_CAP : 1); page += 1) {
    const result = await graphPage(ctx, url) as Collection<Task>;
    pages.push(result);
    nextLink = result['@odata.nextLink'];
    if (!all || !nextLink) break;
    validateContinuation(nextLink, ctx);
    url = nextLink;
  }
  let tasks = pages.flatMap((page) => page.value ?? []);
  if (args.includes('--mine')) {
    const me = await graphRequest(ctx, 'GET', '/me') as { id?: string };
    if (!me.id) usage('Could not identify the current user for --mine.');
    tasks = tasks.filter((task) => Object.hasOwn(task.assignments ?? {}, me.id!));
  }
  if (args.includes('--open')) tasks = tasks.filter((task) => (task.percentComplete ?? 0) < 100);
  if (bucketId) tasks = tasks.filter((task) => task.bucketId === bucketId);
  const createdAfter = option(args, '--created-after');
  const createdBound = createdAfter ? dateBound(createdAfter, '--created-after') : undefined;
  if (createdBound !== undefined) tasks = tasks.filter((task) => after(task.createdDateTime, createdBound));
  const completedAfter = option(args, '--completed-after');
  const completedBound = completedAfter ? dateBound(completedAfter, '--completed-after') : undefined;
  if (completedBound !== undefined) tasks = tasks.filter((task) => after(task.completedDateTime, completedBound));

  const first = pages[0] ?? { value: [] };
  const response: Collection<Task> = all ? { ...first, value: tasks, ...(nextLink ? { '@odata.nextLink': nextLink } : {}) } : { ...first, value: tasks };
  if (all && !nextLink) delete response['@odata.nextLink'];
  if (args.includes('--json')) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}\n`);
    printContinuation(nextLink);
    return;
  }

  const bucketNames = new Map<string, string>();
  const assigneeNames = new Map<string, string>();
  if (planId && fields.includes('bucket') && tasks.some((task) => task.bucketId)) {
    const buckets = await graphRequest(ctx, 'GET', `/planner/plans/${encodeURIComponent(planId)}/buckets`) as Collection<Bucket>;
    for (const bucket of buckets.value ?? []) if (bucket.id) bucketNames.set(bucket.id, bucket.name ?? bucket.id);
  }
  if (planId && fields.includes('assignees') && tasks.some((task) => Object.keys(task.assignments ?? {}).length)) {
    const plan = await graphRequest(ctx, 'GET', `/planner/plans/${encodeURIComponent(planId)}`) as Plan;
    if (plan.owner) {
      const members = await graphRequest(ctx, 'GET', `/groups/${encodeURIComponent(plan.owner)}/members`) as Collection<Member>;
      for (const member of members.value ?? []) if (member.id) assigneeNames.set(member.id, member.displayName ?? member.mail ?? member.userPrincipalName ?? member.id);
    }
  }
  const detailFields = args.includes('--details') ? ['description', 'checklist'] : [];
  const displayFields = [...fields, ...detailFields];
  const labels = displayFields.map((field) => field.toUpperCase());
  const rows = tasks.map((task) => displayFields.map((field) => formatField(task, field, bucketNames, assigneeNames)));
  process.stdout.write(table(labels, rows));
  printContinuation(nextLink);
}

async function showTask(ctx: Context, args: string[]): Promise<void> {
  const idInput = args.find((arg) => !arg.startsWith('-'));
  if (!idInput) usage("Usage: planner task <id|url>.");
  const id = parseTaskId(idInput!);
  const task = await graphRequest(ctx, 'GET', `/planner/tasks/${encodeURIComponent(id)}?$expand=details`) as Task;
  if (args.includes('--json')) process.stdout.write(`${JSON.stringify(task, null, 2)}\n`);
  else process.stdout.write(`${JSON.stringify(task, null, 2)}\n`);
}

export function parseTaskId(value: string): string {
  if (isGuid(value)) return value;
  let parsed: URL;
  try { parsed = new URL(value); } catch { usage(`Invalid Task id '${value}'. Pass a Task GUID or Planner Task URL.`); }
  const parts = parsed!.pathname.split('/').filter(Boolean);
  const candidate = parts.at(-1) ?? '';
  if (!/^https?:$/.test(parsed!.protocol) || !isGuid(candidate)) usage(`Invalid Task id '${value}'. Pass a Task GUID or Planner Task URL.`);
  return candidate;
}

function formatField(task: Task, field: string, buckets: Map<string, string>, members: Map<string, string>): string {
  if (field === 'title') {
    const title = task.title ?? '';
    return title.length > TITLE_WIDTH ? `${title.slice(0, TITLE_WIDTH - 1)}…` : title;
  }
  if (field === 'bucket') return buckets.get(task.bucketId ?? '') ?? task.bucketId ?? '';
  if (field === 'percent') return task.percentComplete === undefined ? '' : `${task.percentComplete}%`;
  if (field === 'due') return task.dueDateTime?.slice(0, 10) ?? '';
  if (field === 'assignees') return Object.keys(task.assignments ?? {}).map((id) => members.get(id) ?? id).join(', ');
  if (field === 'description') return task.details?.description ?? '';
  if (field === 'checklist') return Object.values(task.details?.checklist ?? {}).map((item) => `${item.isChecked ? '[x]' : '[ ]'} ${item.title ?? ''}`).join('; ');
  const value = task[FIELD_MAP[field]!];
  return value === undefined || value === null ? '' : String(value);
}

function table(headers: string[], rows: string[][]): string {
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values: string[]) => values.map((value, index) => value.padEnd(widths[index]!)).join('  ').trimEnd();
  return `${line(headers)}\n${rows.map(line).join('\n')}${rows.length ? '\n' : ''}`;
}

function dateBound(date: string, flag: string): number {
  const match = /^(\d{4}-\d{2}-\d{2})$/.exec(date);
  const target = match ? Date.parse(`${date}T00:00:00Z`) : Number.NaN;
  if (!Number.isFinite(target) || new Date(target).toISOString().slice(0, 10) !== date) usage(`Option '${flag}' requires a valid YYYY-MM-DD date.`);
  return target!;
}

function after(value: string | undefined, target: number): boolean {
  if (!value) return false;
  const timestamp = Date.parse(value);
  if (!Number.isFinite(timestamp)) return false;
  return timestamp > target;
}

function option(args: string[], name: string): string | undefined {
  const index = args.indexOf(name);
  if (index < 0) return undefined;
  const value = args[index + 1];
  if (!value || value.startsWith('-')) usage(`Option '${name}' requires a value.`);
  return value;
}

function graphPage(ctx: Context, url: string): Promise<unknown> {
  return /^https?:\/\//i.test(url) ? graphRequest(ctx, 'GET', url, { rawUrl: url }) : graphRequest(ctx, 'GET', url);
}

function decodeCursor(token: string, ctx: Context): string {
  let link: string;
  try { link = Buffer.from(token, 'base64url').toString('utf8'); }
  catch { usage('Invalid cursor. Copy the cursor from the previous command output.'); }
  validateContinuation(link!, ctx);
  return link!;
}

function validateContinuation(link: string, ctx: Context): void {
  let url: URL;
  try { url = new URL(link, ctx.env.PLANNER_GRAPH_URL || 'https://graph.microsoft.com/v1.0'); }
  catch { usage('Graph returned an invalid continuation URL.'); }
  const base = new URL(ctx.env.PLANNER_GRAPH_URL || 'https://graph.microsoft.com/v1.0');
  const basePath = base.pathname.replace(/\/$/, '');
  if (url!.origin !== base.origin || (basePath && url!.pathname !== basePath && !url!.pathname.startsWith(`${basePath}/`))) {
    usage('Graph returned a continuation URL outside the configured Graph URL.');
  }
}

function printContinuation(nextLink: string | undefined): void {
  if (nextLink) process.stderr.write(`more available: --cursor ${Buffer.from(nextLink).toString('base64url')} (or --all)\n`);
}
