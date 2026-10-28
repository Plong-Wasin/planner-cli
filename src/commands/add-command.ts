import { randomUUID } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { graphRequest, type Context } from './auth-commands.js';
import { CliError, usage } from '../utils/errors.js';
import { resolveAssigneeId, resolveBucketId, resolvePlanId } from '../utils/name-resolution.js';

const PRIORITIES: Record<string, number> = { urgent: 1, important: 3, medium: 5, low: 9 };

interface TaskDetails {
  '@odata.etag'?: string;
}

export async function runAdd(ctx: Context, args: string[]): Promise<void> {
  const options = parseOptions(args);
  if (options.data !== undefined) {
    if (options.positionals.length || options.bucket !== undefined || options.assign !== undefined || options.noAssign
      || options.priority !== undefined || options.due !== undefined || options.start !== undefined
      || options.description !== undefined || options.checklist !== undefined) {
      usage("Use '--data' with fields on each input record, not Task flags or a title.");
    }
    const entries = bulkEntries(loadText(options.data, '--data'));
    let failed = false;
    for (const entry of entries) {
      try {
        if (entry.error !== undefined) throw new CliError(entry.error, 2);
        const record = entry.record;
        if (!record || Array.isArray(record) || typeof record !== 'object') usage('Each Task record must be a JSON object.');
        const task = bulkOptions(record as Record<string, unknown>);
        await createTask(ctx, task.title, task);
      } catch (error) {
        failed = true;
        const message = error instanceof Error ? error.message : String(error);
        process.stderr.write(`line ${entry.line}: ${message}\n`);
      }
    }
    if (failed) process.exitCode = 1;
    return;
  }
  const title = options.positionals[0];
  if (!title) usage("Usage: planner add <title> [--bucket <name|id>] [--assign <me|email|name|id>] [options]");
  if (options.positionals.length > 1) usage('Task titles with spaces must be quoted as one argument.');

  await createTask(ctx, title, options);
}

interface AddOptions {
  title?: string;
  bucket?: string;
  assign?: string;
  noAssign?: boolean;
  priority?: string;
  due?: string;
  start?: string;
  description?: string;
  checklist?: string;
}

async function createTask(ctx: Context, title: string | undefined, options: AddOptions): Promise<void> {
  if (!title) usage("Each Task record needs a non-empty 'title'.");
  if (options.noAssign && options.assign !== undefined) usage("Use either 'assign' or no assignment, not both.");

  const requestedPlan = ctx.options.plan || ctx.env.PLANNER_PLAN || ctx.profile.defaultPlan;
  if (!requestedPlan) usage("A Default Plan is required. Set one with 'planner config set plan <name|id>' or pass --plan <name|id>.");
  const planId = await resolvePlanId(ctx, requestedPlan);
  const requestedBucket = options.bucket ?? ctx.profile.defaultBucket;
  if (!requestedBucket) usage("A Default Bucket is required. Set one with 'planner config set bucket <name|id>' or pass --bucket <name|id>.");
  const bucketId = await resolveBucketId(ctx, planId, requestedBucket);

  const payload: Record<string, unknown> = {
    planId,
    bucketId,
    title,
    priority: options.priority === undefined ? 3 : parsePriority(options.priority),
  };
  if (options.assign !== undefined || !options.noAssign) {
    const assignee = await resolveAssigneeId(ctx, planId, options.assign ?? 'me');
    payload.assignments = {
      [assignee]: { '@odata.type': 'microsoft.graph.plannerAssignment', orderHint: ' !' },
    };
  }
  if (options.due !== undefined) payload.dueDateTime = parseDate(options.due, '--due');
  if (options.start !== undefined) payload.startDateTime = parseDate(options.start, '--start');

  const task = await graphRequest(ctx, 'POST', '/planner/tasks', { body: JSON.stringify(payload) }) as { id?: string };
  if (!task?.id) throw new CliError('Graph created the Task without returning its id.');
  const details: Record<string, unknown> = {};
  if (options.description !== undefined) details.description = loadText(options.description, '--description');
  if (options.checklist !== undefined) details.checklist = checklist(loadText(options.checklist, '--checklist'));
  // Print the id first: the Task exists even if its Details write fails, and the caller needs the id to fix it.
  process.stdout.write(`${task.id}\n`);
  if (Object.keys(details).length) {
    try { await writeDetails(ctx, task.id, details); }
    catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      throw new CliError(`Task ${task.id} was created but its Details were not saved: ${reason}`, error instanceof CliError ? error.exitCode : 1, error instanceof CliError ? error.status : undefined);
    }
  }
}

async function writeDetails(ctx: Context, taskId: string, body: Record<string, unknown>): Promise<void> {
  const path = `/planner/tasks/${encodeURIComponent(taskId)}/details`;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const current = await graphRequest(ctx, 'GET', path) as TaskDetails;
    if (!current?.['@odata.etag']) throw new CliError(`Task Details for '${taskId}' did not include an etag.`);
    try {
      await graphRequest(ctx, 'PATCH', path, {
        body: JSON.stringify(body),
        headers: { 'If-Match': current['@odata.etag'] },
      });
      return;
    } catch (error) {
      if (!(error instanceof CliError) || error.status !== 412) throw error;
      if (attempt === 1) throw new CliError(`Task Details for '${taskId}' changed twice while being updated. Retry the command.`);
    }
  }
}

function parseOptions(args: string[]): {
  positionals: string[]; bucket?: string; assign?: string; noAssign: boolean;
  priority?: string; due?: string; start?: string; description?: string; checklist?: string; data?: string;
} {
  const parsed: ReturnType<typeof parseOptions> = { positionals: [], noAssign: false };
  const valued: Record<string, keyof typeof parsed> = {
    '--bucket': 'bucket', '--assign': 'assign', '--priority': 'priority', '--due': 'due',
    '--start': 'start', '--description': 'description', '--checklist': 'checklist', '--data': 'data',
  };
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (arg === '--no-assign') { parsed.noAssign = true; continue; }
    if (arg === '--json') continue;
    const [name, inline] = arg.includes('=') ? [arg.slice(0, arg.indexOf('=')), arg.slice(arg.indexOf('=') + 1)] : [arg, undefined];
    const key = valued[name];
    if (key) {
      const value = inline ?? args[++index];
      if (value === undefined || (inline === undefined && value.startsWith('--'))) usage(`Option '${name}' requires a value.`);
      (parsed[key] as string | undefined) = value;
      continue;
    }
    if (arg.startsWith('-')) usage(`Unknown add option: ${arg}.`);
    parsed.positionals.push(arg);
  }
  if (parsed.noAssign && parsed.assign !== undefined) usage("Use either '--assign' or '--no-assign', not both.");
  return parsed;
}

function bulkEntries(input: string): Array<{ line: number; record?: unknown; error?: string }> {
  if (input.trimStart().startsWith('[')) {
    try {
      const value: unknown = JSON.parse(input);
      if (Array.isArray(value)) return value.map((record, index) => ({ line: index + 1, record }));
      return [{ line: 1, record: value }];
    } catch {
      // A malformed array is also malformed NDJSON; process each physical line below.
    }
  }
  const entries: Array<{ line: number; record?: unknown; error?: string }> = [];
  input.split(/\r?\n/).forEach((line, index) => {
    if (!line.trim()) return;
    try { entries.push({ line: index + 1, record: JSON.parse(line) as unknown }); }
    catch { entries.push({ line: index + 1, error: 'Invalid JSON.' }); }
  });
  return entries;
}

function bulkOptions(record: Record<string, unknown>): AddOptions {
  const allowed = new Set(['title', 'bucket', 'assign', 'priority', 'due', 'start', 'description', 'checklist']);
  for (const key of Object.keys(record)) {
    if (!allowed.has(key)) usage(`Unknown Task field '${key}'.`);
  }
  const options: AddOptions = {};
  for (const key of allowed) {
    const value = record[key];
    if (value === undefined) continue;
    if (key === 'checklist' && Array.isArray(value) && value.every((item) => typeof item === 'string')) {
      options.checklist = value.join('\n');
      continue;
    }
    if (key === 'priority' && typeof value === 'number') {
      options.priority = String(value);
      continue;
    }
    if (typeof value !== 'string') usage(`Task field '${key}' must be text${key === 'checklist' ? ' or a list of text' : key === 'priority' ? ' or a number' : ''}.`);
    if (key === 'title') options.title = value;
    else if (key === 'bucket') options.bucket = value;
    else if (key === 'assign') options.assign = value;
    else if (key === 'priority') options.priority = value;
    else if (key === 'due') options.due = value;
    else if (key === 'start') options.start = value;
    else if (key === 'description') options.description = value;
    else if (key === 'checklist') options.checklist = value;
  }
  return options;
}

function parsePriority(value: string): number {
  const named = PRIORITIES[value.toLocaleLowerCase()];
  if (named !== undefined) return named;
  if (/^(?:10|[0-9])$/.test(value)) return Number(value);
  usage("Priority must be urgent, important, medium, low, or a number from 0 to 10.");
}

function parseDate(value: string, flag: string): string {
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const date = new Date(`${value}T00:00:00.000Z`);
    if (date.toISOString().slice(0, 10) === value) return date.toISOString();
  } else if (/^\d{4}-\d\d-\d\dT/.test(value) && Number.isFinite(Date.parse(value))) {
    return value;
  }
  usage(`Option '${flag}' must be YYYY-MM-DD or a full ISO date and time.`);
}

function loadText(value: string, flag: string): string {
  if (value === '-') return readFileSync(0, 'utf8');
  if (value.startsWith('@')) {
    const path = value.slice(1);
    if (!path) usage(`Option '${flag}' needs a file path after '@'.`);
    try { return readFileSync(path, 'utf8'); }
    catch { usage(`Cannot read input file for '${flag}': ${path}.`); }
  }
  return value;
}

function checklist(input: string): Record<string, unknown> {
  return Object.fromEntries(input.split(/\r?\n/).filter((line) => line.length > 0).map((title) => [randomUUID(), {
    '@odata.type': 'microsoft.graph.plannerChecklistItem', title, isChecked: false,
  }]));
}
