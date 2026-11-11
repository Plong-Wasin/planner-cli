import { graphRequest, type Context } from './auth-commands.js';
import { CliError, usage } from '../utils/errors.js';
import { collection, isGuid, planMembers, resolveBucketId, resolvePlanId } from '../utils/name-resolution.js';

interface Task {
  id?: string;
  title?: string;
  bucketId?: string;
  percentComplete?: number;
  dueDateTime?: string;
  createdDateTime?: string;
  completedDateTime?: string;
  assignments?: Record<string, unknown>;
  planId?: string;
  startDateTime?: string;
  details?: { description?: string; checklist?: Record<string, { title?: string; isChecked?: boolean }> };
  [key: string]: unknown;
}

interface Collection<T> {
  value?: T[];
  '@odata.nextLink'?: string;
  [key: string]: unknown;
}

interface Plan { owner?: string; container?: { url?: string } }
interface PlanRef extends Plan { id?: string; title?: string }
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
  const allPlans = args.includes('--all-plans');
  if (allPlans && (option(args, '--plan') !== undefined || ctx.options.plan !== undefined || option(args, '--bucket') !== undefined || option(args, '--cursor') !== undefined)) {
    usage("Option '--all-plans' cannot be combined with --plan, --bucket, or --cursor.");
  }
  const planInput = allPlans ? undefined : option(args, '--plan') ?? ctx.options.plan ?? ctx.env.PLANNER_PLAN ?? ctx.profile.defaultPlan;
  const planId = planInput ? await resolvePlanId(ctx, planInput) : undefined;
  const bucketInput = option(args, '--bucket');
  const bucketId = bucketInput ? (isGuid(bucketInput) ? bucketInput : planId ? await resolveBucketId(ctx, planId, bucketInput) : usage("A Default Plan is required to resolve a Bucket name; pass --plan <name|id>.")) : undefined;
  const search = option(args, '--search')?.toLocaleLowerCase();
  const cursor = option(args, '--cursor');
  const cursorUrl = cursor ? decodeCursor(cursor, ctx) : undefined;
  const fieldsArg = option(args, '--fields');
  const fields = fieldsArg ? fieldsArg.split(',').map((field) => field.trim()).filter(Boolean) : DEFAULT_FIELDS;
  if (!fields.length || fields.some((field) => !Object.hasOwn(FIELD_MAP, field))) {
    usage(`Unknown field in --fields. Choose: ${Object.keys(FIELD_MAP).join(', ')}.`);
  }
  const selected = new Set(fields.map((field) => FIELD_MAP[field]!));
  if (bucketId) selected.add('bucketId');
  if (search !== undefined) selected.add('title');
  if (allPlans) selected.add('planId');
  if (args.includes('--mine')) selected.add('assignments');
  if (args.includes('--open')) selected.add('percentComplete');
  if (option(args, '--created-after')) selected.add('createdDateTime');
  if (option(args, '--completed-after')) selected.add('completedDateTime');

  const query = new URLSearchParams();
  query.set('$top', String(DEFAULT_PAGE_SIZE));
  query.set('$select', [...selected].join(','));
  if (args.includes('--details')) query.set('$expand', 'details');
  const all = args.includes('--all') || allPlans;
  const createdAfter = option(args, '--created-after');
  const createdBound = createdAfter ? dateBound(createdAfter, '--created-after') : undefined;
  const completedAfter = option(args, '--completed-after');
  const completedBound = completedAfter ? dateBound(completedAfter, '--completed-after') : undefined;
  let meId: string | undefined;
  if (args.includes('--mine')) {
    const me = await graphRequest(ctx, 'GET', '/me') as { id?: string };
    if (!me.id) usage('Could not identify the current user for --mine.');
    meId = me.id;
  }
  // The search text is matched against the title, and against the description when --details loads it.
  const keep = (task: Task): boolean => (meId === undefined || Object.hasOwn(task.assignments ?? {}, meId))
    && (!args.includes('--open') || (task.percentComplete ?? 0) < 100)
    && (!bucketId || task.bucketId === bucketId)
    && (createdBound === undefined || after(task.createdDateTime, createdBound))
    && (completedBound === undefined || after(task.completedDateTime, completedBound))
    && (search === undefined || (task.title ?? '').toLocaleLowerCase().includes(search) || (task.details?.description ?? '').toLocaleLowerCase().includes(search));
  const filtered = meId !== undefined || args.includes('--open') || bucketId !== undefined
    || createdBound !== undefined || completedBound !== undefined || search !== undefined;

  const detailFields = args.includes('--details') ? ['description', 'checklist'] : [];
  const displayFields = [...fields, ...detailFields];

  if (allPlans) {
    const plans = await collection<PlanRef>(ctx, '/me/planner/plans');
    const groups: { plan: PlanRef; tasks: Task[] }[] = [];
    if (meId !== undefined) {
      // Graph already lists the Tasks assigned to the user across every Plan, so one paged request replaces a request per Plan.
      const result = await collectTasks(ctx, `/me/planner/tasks?${query}`, keep, all, filtered);
      if (result.nextLink) process.stderr.write('more Tasks than the safety cap; narrow it with --search or other filters.\n');
      const byPlan = new Map<string, Task[]>();
      for (const task of result.tasks) {
        const id = typeof task.planId === 'string' ? task.planId : '';
        const group = byPlan.get(id);
        if (group) group.push(task); else byPlan.set(id, [task]);
      }
      for (const [id, tasks] of byPlan) groups.push({ plan: plans.find((plan) => plan.id === id) ?? { id }, tasks });
    } else {
      for (const plan of plans) {
        if (!plan.id) continue;
        // A Plan can be listed yet unreadable (its group was deleted, access was removed); skip it instead of losing every other Plan.
        const result = await collectTasks(ctx, `/planner/plans/${encodeURIComponent(plan.id)}/tasks?${query}`, keep, all, filtered).catch((error: unknown) => {
          if (!isUnreadable(error)) throw error;
          process.stderr.write(`skipped plan '${plan.title ?? plan.id}': ${(error as CliError).message}\n`);
          return undefined;
        });
        if (!result) continue;
        if (result.nextLink) process.stderr.write(`plan '${plan.title ?? plan.id}' has more Tasks than the safety cap; narrow it with --search or other filters.\n`);
        if (result.tasks.length) groups.push({ plan, tasks: result.tasks });
      }
    }
    const found: Task[] = [];
    const rows: string[][] = [];
    for (const { plan, tasks } of groups) {
      if (!plan.id) continue;
      found.push(...tasks);
      const names = await nameMaps(ctx, plan.id, plan.owner ? plan : undefined, fields, tasks).catch((error: unknown) => {
        if (!isUnreadable(error)) throw error;
        process.stderr.write(`could not resolve Bucket or Assignee names for plan '${plan.title ?? plan.id}': ${(error as CliError).message}\n`);
        return { buckets: new Map<string, string>(), assignees: new Map<string, string>() };
      });
      for (const task of tasks) rows.push([plan.title ?? plan.id, ...displayFields.map((field) => formatField(task, field, names.buckets, names.assignees))]);
    }
    if (args.includes('--json')) {
      process.stdout.write(`${JSON.stringify({ value: found }, null, 2)}\n`);
      return;
    }
    process.stdout.write(table(['PLAN', ...displayFields.map((field) => field.toUpperCase())], rows));
    return;
  }

  const basePath = planId ? `/planner/plans/${encodeURIComponent(planId)}/tasks` : '/me/planner/tasks';
  const start = cursorUrl ?? `${basePath}?${query}`;
  const { pages, tasks, nextLink } = await collectTasks(ctx, start, keep, all, filtered);

  const first = pages[0] ?? { value: [] };
  const response: Collection<Task> = all ? { ...first, value: tasks, ...(nextLink ? { '@odata.nextLink': nextLink } : {}) } : { ...first, value: tasks };
  if (all && !nextLink) delete response['@odata.nextLink'];
  if (args.includes('--json')) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}\n`);
    printContinuation(nextLink);
    return;
  }

  const names = planId ? await nameMaps(ctx, planId, undefined, fields, tasks) : { buckets: new Map<string, string>(), assignees: new Map<string, string>() };
  const labels = displayFields.map((field) => field.toUpperCase());
  const rows = tasks.map((task) => displayFields.map((field) => formatField(task, field, names.buckets, names.assignees)));
  process.stdout.write(table(labels, rows));
  printContinuation(nextLink);
}

function isUnreadable(error: unknown): boolean {
  return error instanceof CliError && (error.status === 403 || error.status === 404);
}

// Filters run client-side, so keep paging until a full page of matches (or the end) instead of filtering one page.
async function collectTasks(ctx: Context, startUrl: string, keep: (task: Task) => boolean, all: boolean, filtered: boolean): Promise<{ pages: Collection<Task>[]; tasks: Task[]; nextLink: string | undefined }> {
  const pages: Collection<Task>[] = [];
  const tasks: Task[] = [];
  let nextLink: string | undefined;
  let url = startUrl;
  for (let page = 0; page < (all || filtered ? PAGE_CAP : 1); page += 1) {
    const result = await graphPage(ctx, url) as Collection<Task>;
    pages.push(result);
    tasks.push(...(result.value ?? []).filter(keep));
    nextLink = result['@odata.nextLink'];
    if (!nextLink || (!all && tasks.length >= DEFAULT_PAGE_SIZE)) break;
    if (!all && !filtered) break;
    validateContinuation(nextLink, ctx);
    url = nextLink;
  }
  return { pages, tasks, nextLink };
}

// Bucket and Assignee names are looked up only for columns that are shown and Tasks that have them.
async function nameMaps(ctx: Context, planId: string, knownPlan: PlanRef | undefined, fields: string[], tasks: Task[]): Promise<{ buckets: Map<string, string>; assignees: Map<string, string> }> {
  const buckets = new Map<string, string>();
  const assignees = new Map<string, string>();
  if (fields.includes('bucket') && tasks.some((task) => task.bucketId)) {
    for (const bucket of await collection<Bucket>(ctx, `/planner/plans/${encodeURIComponent(planId)}/buckets`)) if (bucket.id) buckets.set(bucket.id, bucket.name ?? bucket.id);
  }
  if (fields.includes('assignees') && tasks.some((task) => Object.keys(task.assignments ?? {}).length)) {
    const plan = knownPlan ?? await graphRequest(ctx, 'GET', `/planner/plans/${encodeURIComponent(planId)}`) as Plan;
    for (const member of await planMembers(ctx, plan)) if (member.id) assignees.set(member.id, member.displayName ?? member.mail ?? member.userPrincipalName ?? member.id);
  }
  return { buckets, assignees };
}

async function showTask(ctx: Context, args: string[]): Promise<void> {
  const idInput = args.find((arg) => !arg.startsWith('-'));
  if (!idInput) usage("Usage: planner task <id|url>.");
  const id = parseTaskId(idInput!);
  const task = await graphRequest(ctx, 'GET', `/planner/tasks/${encodeURIComponent(id)}?$expand=details`) as Task;
  if (args.includes('--json')) { process.stdout.write(`${JSON.stringify(task, null, 2)}\n`); return; }
  const planId = typeof task.planId === 'string' ? task.planId : undefined;
  const bucketNames = new Map<string, string>();
  const assigneeNames = new Map<string, string>();
  if (planId && task.bucketId) {
    for (const bucket of await collection<Bucket>(ctx, `/planner/plans/${encodeURIComponent(planId)}/buckets`)) if (bucket.id) bucketNames.set(bucket.id, bucket.name ?? bucket.id);
  }
  if (planId && Object.keys(task.assignments ?? {}).length) {
    const plan = await graphRequest(ctx, 'GET', `/planner/plans/${encodeURIComponent(planId)}`) as Plan;
    for (const member of await planMembers(ctx, plan)) {
      if (member.id) assigneeNames.set(member.id, member.displayName ?? member.mail ?? member.userPrincipalName ?? member.id);
    }
  }
  const lines = ['id', 'title', 'bucket', 'percent', 'priority', 'start', 'due', 'assignees', 'description', 'checklist']
    .map((field) => [field, formatField(task, field, bucketNames, assigneeNames, false)] as const)
    .filter(([, value]) => value !== '')
    .map(([field, value]) => `${field}: ${value}`);
  process.stdout.write(`${lines.join('\n')}\n`);
}

// Planner Task ids are 28-character opaque strings; a GUID is accepted too.
function isTaskId(value: string): boolean {
  return /^[A-Za-z0-9_-]{28}$/.test(value) || isGuid(value);
}

export function parseTaskId(value: string): string {
  if (isTaskId(value)) return value;
  let parsed: URL;
  try { parsed = new URL(value); } catch { usage(`Invalid Task id '${value}'. Pass a Planner Task id or Planner Task URL.`); }
  const parts = parsed!.pathname.split('/').filter(Boolean);
  const candidate = parts.at(-1) ?? '';
  if (!/^https?:$/.test(parsed!.protocol) || !isTaskId(candidate)) usage(`Invalid Task id '${value}'. Pass a Planner Task id or Planner Task URL.`);
  return candidate;
}

function formatField(task: Task, field: string, buckets: Map<string, string>, members: Map<string, string>, truncate = true): string {
  if (field === 'title') {
    const title = task.title ?? '';
    return truncate && title.length > TITLE_WIDTH ? `${title.slice(0, TITLE_WIDTH - 1)}…` : title;
  }
  if (field === 'bucket') return buckets.get(task.bucketId ?? '') ?? task.bucketId ?? '';
  if (field === 'percent') return task.percentComplete === undefined ? '' : `${task.percentComplete}%`;
  if (field === 'due') return task.dueDateTime?.slice(0, 10) ?? '';
  if (field === 'start') return task.startDateTime?.slice(0, 10) ?? '';
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
  const inline = args.find((arg) => arg.startsWith(`${name}=`));
  if (inline !== undefined) {
    const value = inline.slice(name.length + 1);
    if (!value) usage(`Option '${name}' requires a value.`);
    return value;
  }
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
