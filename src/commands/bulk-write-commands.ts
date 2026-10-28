import { readFileSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import { graphRequest, type Context } from './auth-commands.js';
import { CliError, usage } from '../utils/errors.js';
import { resolveAssigneeId, resolveBucketId } from '../utils/name-resolution.js';
import { parseTaskId } from './task-read-commands.js';

interface Task {
  planId?: string;
  '@odata.etag'?: string;
  assignments?: Record<string, unknown>;
}

interface Options {
  positionals: string[];
  bucket?: string;
  assign?: string;
  noAssign: boolean;
  priority?: string;
  due?: string;
  start?: string;
  description?: string;
  checklist?: string;
  percent?: string;
  yes: boolean;
}

interface PreparedUpdate {
  fields: Record<string, unknown>;
  description?: string;
  checklist?: string;
  bucket?: string;
  assign?: string;
  noAssign: boolean;
}

const PRIORITIES: Record<string, number> = { urgent: 1, important: 3, medium: 5, low: 9 };

export async function runBulkWrite(ctx: Context, command: 'update' | 'done' | 'delete', args: string[]): Promise<void> {
  const options = parseOptions(command, args);
  const ids = options.positionals.map(parseTaskId);
  if (!ids.length) usage(`Usage: planner ${command} <id…>${command === 'delete' ? ' --yes' : ''} [options]`);
  if (command === 'delete' && !options.yes) usage("Refusing to delete Tasks without '--yes'.");
  const update = command === 'update' ? prepareUpdate(options) : undefined;

  const failures: string[] = [];
  for (const id of ids) {
    try {
      if (command === 'delete') await deleteTask(ctx, id);
      else if (command === 'done') await updateTask(ctx, id, { percentComplete: 100 });
      else {
        const change = await updateBody(ctx, id, update!);
        await updateTask(ctx, id, change.fields, change.details);
      }
      process.stdout.write(`${command === 'done' ? 'Completed' : command === 'delete' ? 'Deleted' : 'Updated'} #${id}\n`);
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      failures.push(id);
      process.stderr.write(`#${id}: ${reason}\n`);
    }
  }
  if (failures.length) {
    process.stderr.write(`${failures.length} Task${failures.length === 1 ? '' : 's'} failed.\n`);
    process.exitCode = 1;
  }
}

function prepareUpdate(options: Options): PreparedUpdate {
  const fields: Record<string, unknown> = {};
  if (options.priority !== undefined) fields.priority = parsePriority(options.priority);
  if (options.due !== undefined) fields.dueDateTime = parseDate(options.due, '--due');
  if (options.start !== undefined) fields.startDateTime = parseDate(options.start, '--start');
  if (options.percent !== undefined) fields.percentComplete = parsePercent(options.percent);
  const description = options.description === undefined ? undefined : loadText(options.description, '--description');
  const checklistText = options.checklist === undefined ? undefined : loadText(options.checklist, '--checklist');
  if (!Object.keys(fields).length && options.bucket === undefined && options.assign === undefined && !options.noAssign
      && description === undefined && checklistText === undefined) usage('Provide at least one field to update.');
  return { fields, description, checklist: checklistText, bucket: options.bucket, assign: options.assign, noAssign: options.noAssign };
}

async function updateBody(ctx: Context, id: string, update: PreparedUpdate): Promise<{ fields: Record<string, unknown>; details?: Record<string, unknown> }> {
  const fields = { ...update.fields };
  let task: Task | undefined;
  const currentTask = async (): Promise<Task> => {
    task ??= await graphRequest(ctx, 'GET', `/planner/tasks/${encodeURIComponent(id)}`) as Task;
    return task;
  };
  if (update.assign !== undefined || update.noAssign || update.bucket !== undefined) {
    const current = await currentTask();
    if (update.assign !== undefined) {
      const planId = current.planId || ctx.options.plan || ctx.env.PLANNER_PLAN || ctx.profile.defaultPlan;
      if (!planId) usage("A Default Plan is required to resolve an Assignee. Set one with 'planner config set plan <name|id>'.");
      const assigneeId = await resolveAssigneeId(ctx, planId, update.assign);
      fields.assignments = { [assigneeId]: { '@odata.type': 'microsoft.graph.plannerAssignment', orderHint: ' !' } };
    } else if (update.noAssign) {
      fields.assignments = Object.fromEntries(Object.keys(current.assignments ?? {}).map((assigneeId) => [assigneeId, null]));
    }
    if (update.bucket !== undefined) {
      const planId = current.planId || ctx.options.plan || ctx.env.PLANNER_PLAN || ctx.profile.defaultPlan;
      if (!planId) usage("A Default Plan is required to move a Task. Set one with 'planner config set plan <name|id>'.");
      fields.bucketId = await resolveBucketId(ctx, planId, update.bucket);
    }
  }
  const details: Record<string, unknown> = {};
  if (update.description !== undefined) details.description = update.description;
  if (update.checklist !== undefined) details.checklist = checklist(update.checklist);
  return { fields, ...(Object.keys(details).length ? { details } : {}) };
}

async function updateTask(ctx: Context, id: string, fields: Record<string, unknown>, details?: Record<string, unknown>): Promise<void> {
  if (Object.keys(fields).length) await writeTask(ctx, id, fields);
  if (details) await writeDetails(ctx, id, details);
}

async function writeTask(ctx: Context, id: string, body: Record<string, unknown>): Promise<void> {
  const path = `/planner/tasks/${encodeURIComponent(id)}`;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const task = await graphRequest(ctx, 'GET', path) as Task;
    if (!task?.['@odata.etag']) throw new CliError(`Task '${id}' did not include an etag.`);
    try {
      await graphRequest(ctx, 'PATCH', path, {
        body: JSON.stringify(body), headers: { 'If-Match': task['@odata.etag'] },
      });
      return;
    } catch (error) {
      if (!(error instanceof CliError) || error.status !== 412) throw error;
      if (attempt === 1) throw new CliError(`Task '${id}' changed twice while being updated. Retry the command.`);
    }
  }
}

async function deleteTask(ctx: Context, id: string): Promise<void> {
  const path = `/planner/tasks/${encodeURIComponent(id)}`;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const task = await graphRequest(ctx, 'GET', path) as Task;
    if (!task?.['@odata.etag']) throw new CliError(`Task '${id}' did not include an etag.`);
    try {
      await graphRequest(ctx, 'DELETE', path, { headers: { 'If-Match': task['@odata.etag'] } });
      return;
    } catch (error) {
      if (!(error instanceof CliError) || error.status !== 412) throw error;
      if (attempt === 1) throw new CliError(`Task '${id}' changed twice while being deleted. Retry the command.`);
    }
  }
}

async function writeDetails(ctx: Context, id: string, body: Record<string, unknown>): Promise<void> {
  const path = `/planner/tasks/${encodeURIComponent(id)}/details`;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const current = await graphRequest(ctx, 'GET', path) as { '@odata.etag'?: string };
    if (!current?.['@odata.etag']) throw new CliError(`Task Details for '${id}' did not include an etag.`);
    try {
      await graphRequest(ctx, 'PATCH', path, { body: JSON.stringify(body), headers: { 'If-Match': current['@odata.etag'] } });
      return;
    } catch (error) {
      if (!(error instanceof CliError) || error.status !== 412) throw error;
      if (attempt === 1) throw new CliError(`Task Details for '${id}' changed twice while being updated. Retry the command.`);
    }
  }
}

function parseOptions(command: 'update' | 'done' | 'delete', args: string[]): Options {
  const parsed: Options = { positionals: [], noAssign: false, yes: false };
  const valued: Record<string, keyof Options> = {
    '--bucket': 'bucket', '--assign': 'assign', '--priority': 'priority', '--due': 'due',
    '--start': 'start', '--description': 'description', '--checklist': 'checklist', '--percent': 'percent',
  };
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (arg === '--yes' && command === 'delete') { parsed.yes = true; continue; }
    if (arg === '--no-assign' && command === 'update') { parsed.noAssign = true; continue; }
    if (arg === '--json') continue;
    const [name, inline] = arg.includes('=') ? [arg.slice(0, arg.indexOf('=')), arg.slice(arg.indexOf('=') + 1)] : [arg, undefined];
    const key = valued[name];
    if (key && command === 'update') {
      const value = inline ?? args[++index];
      if (value === undefined || (inline === undefined && value.startsWith('--'))) usage(`Option '${name}' requires a value.`);
      (parsed[key] as string | undefined) = value;
      continue;
    }
    if (arg.startsWith('-')) usage(`Unknown ${command} option: ${arg}.`);
    parsed.positionals.push(arg);
  }
  if (parsed.noAssign && parsed.assign !== undefined) usage("Use either '--assign' or '--no-assign', not both.");
  return parsed;
}

function parsePriority(value: string): number {
  const named = PRIORITIES[value.toLocaleLowerCase()];
  if (named !== undefined) return named;
  if (/^(?:10|[0-9])$/.test(value)) return Number(value);
  usage('Priority must be urgent, important, medium, low, or a number from 0 to 10.');
}

function parseDate(value: string, flag: string): string {
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const date = new Date(`${value}T00:00:00.000Z`);
    if (date.toISOString().slice(0, 10) === value) return date.toISOString();
  } else if (/^\d{4}-\d\d-\d\dT/.test(value) && Number.isFinite(Date.parse(value))) return value;
  usage(`Option '${flag}' must be YYYY-MM-DD or a full ISO date and time.`);
}

function parsePercent(value: string): number {
  if (/^(?:100|[0-9]{1,2})$/.test(value)) return Number(value);
  usage("Option '--percent' must be a whole number from 0 to 100.");
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
