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
  add        Create a Task or bulk create from NDJSON/JSON with --data @file|-
  update     Update Tasks
  done       Complete Tasks
  delete     Delete Tasks
  api        Send a raw Microsoft Graph request

Options:
  --plan <name|id> Select a Plan (overrides PLANNER_PLAN and Profile default)
  --bucket <name|id> Select the Bucket for a new Task
  --json           Print the raw Graph response
  --all            Follow pages up to the safety cap
  --cursor <token> Continue from a previous page
  tasks --fields <names> Select Task columns (id,title,bucket,percent,due,assignees)
  tasks --details     Include Task description and checklist
  -h, --help     Show help
  -v, --version  Show version

Exit codes:
  0  Success
  1  Server or authentication error
  2  Usage error`;
