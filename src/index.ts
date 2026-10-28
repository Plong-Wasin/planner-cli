import { HELP } from './help.js';
import { VERSION } from './utils/version.js';

const USAGE_ERROR = 2;

function main(args: string[]): void {
  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    process.stdout.write(`${HELP}\n`);
    return;
  }

  if (args.length === 1 && (args[0] === '--version' || args[0] === '-v')) {
    process.stdout.write(`${VERSION}\n`);
    return;
  }

  const command = args[0];
  process.stderr.write(`planner: Unknown command: ${command}. Run 'planner --help' for usage.\n`);
  process.exitCode = USAGE_ERROR;
}

main(process.argv.slice(2));
