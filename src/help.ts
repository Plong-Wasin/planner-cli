export const HELP = `Usage: planner <command> [options]
       planner --help
       planner --version

Commands:
  auth       Sign in and manage Profiles
  config     Set Profile defaults
  me         Show the signed-in user
  plans      List Plans
  buckets    List Buckets
  members    List Plan members
  tasks      List Tasks
  task       Show one Task
  add        Create Tasks
  update     Update Tasks
  done       Complete Tasks
  delete     Delete Tasks
  api        Send a raw Microsoft Graph request

Options:
  -h, --help     Show help
  -v, --version  Show version

Exit codes:
  0  Success
  1  Server or authentication error
  2  Usage error`;
