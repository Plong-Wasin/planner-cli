import { HELP } from './help.js';
import { VERSION } from './utils/version.js';
import { runAuth, selectedContext, type GlobalOptions } from './commands/auth-commands.js';
import { runConfig } from './commands/config-commands.js';
import { runReadCommand } from './commands/read-commands.js';
import { runTaskReadCommand } from './commands/task-read-commands.js';
import { runApi } from './commands/api-commands.js';
import { runAdd } from './commands/add-command.js';
import { CliError, usage } from './utils/errors.js';

const USAGE_ERROR = 2;

function globals(args: string[]): { options: GlobalOptions; rest: string[] } {
  const options: GlobalOptions = {};
  const rest: string[] = [];
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (arg === '-p' || arg === '--profile' || arg === '--client-id' || arg === '--plan') {
      const value = args[index + 1];
      if (!value || value.startsWith('-')) usage(`Option '${arg}' requires a value.`);
      if (arg === '--client-id') options.clientId = value;
      else if (arg === '--plan') options.plan = value;
      else options.profile = value;
      index += 1;
    } else if (arg.startsWith('--profile=')) options.profile = arg.slice('--profile='.length);
    else if (arg.startsWith('--client-id=')) options.clientId = arg.slice('--client-id='.length);
    else if (arg.startsWith('--plan=')) options.plan = arg.slice('--plan='.length);
    else rest.push(arg);
  }
  return { options, rest };
}

async function main(args: string[]): Promise<void> {
  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    process.stdout.write(`${HELP}\n`);
    return;
  }

  if (args.length === 1 && (args[0] === '--version' || args[0] === '-v')) {
    process.stdout.write(`${VERSION}\n`);
    return;
  }

  const { options, rest } = globals(args);
  const [command, subcommand, ...commandArgs] = rest;
  const ctx = selectedContext(process.env, options);
  if (command === 'auth') return runAuth(ctx, subcommand);
  if (command === 'config') return runConfig(ctx, [subcommand, ...commandArgs].filter((value): value is string => value !== undefined));
  if (command === 'api') return runApi(ctx, [subcommand, ...commandArgs].filter((value): value is string => value !== undefined));
  if (command === 'add') return runAdd(ctx, [subcommand, ...commandArgs].filter((value): value is string => value !== undefined));
  if (command === 'me' || command === 'plans' || command === 'buckets' || command === 'members') return runReadCommand(ctx, command, [subcommand, ...commandArgs].filter((value): value is string => value !== undefined));
  if (command === 'tasks') return runTaskReadCommand(ctx, 'tasks', [subcommand, ...commandArgs].filter((value): value is string => value !== undefined));
  if (command === 'task') return runTaskReadCommand(ctx, 'task', [subcommand, ...commandArgs].filter((value): value is string => value !== undefined));
  if (!command) usage("Run 'planner --help' for usage.");
  throw new CliError(`Unknown command: ${command}. Run 'planner --help' for usage.`, USAGE_ERROR);
}

main(process.argv.slice(2)).catch((error: unknown) => {
  const code = error instanceof CliError ? error.exitCode : 1;
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`planner: ${message}\n`);
  process.exitCode = code;
});
