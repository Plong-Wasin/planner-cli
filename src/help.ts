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
  add        Create a Task, or bulk create from NDJSON/JSON with --data @file|-
  update     Update Tasks (bucket, assign, priority, dates, Details, percent)
  done       Complete Tasks
  delete     Delete Tasks (requires --yes)
  api        Send a raw Microsoft Graph request

Global options:
  -p, --profile <name> Select a Profile (or set PLANNER_PROFILE)
  --client-id <id>    Override the OAuth client id (or set PLANNER_CLIENT_ID)
  --plan <name|id>    Select a Plan (overrides PLANNER_PLAN and Profile default)

Commands and options:
  auth login | check | logout
  config show | set plan|bucket <name|id>
  me | plans [--json] [--all] [--cursor <token>]
  buckets | members [--plan <name|id>] [--json] [--all] [--cursor <token>]
  tasks [--plan <name|id>] [--bucket <name|id>] [--mine] [--open]
        [--created-after <date>] [--completed-after <date>]
        [--fields <id,title,bucket,percent,due,assignees,created,completed,priority,start>] [--details]
        [--json] [--all] [--cursor <token>]
  task <id|url> [--json]
  add <title> [--bucket <name|id>] [--assign <me|email|name|id>] [--no-assign]
      [--priority <urgent|important|medium|low|0-10>] [--due <date>] [--start <date>]
      [--description <text|@file|->] [--checklist <text|@file|->]
  add --data <@file|->  Bulk create from NDJSON or a JSON array
  update <id...> [--bucket <name|id>] [--assign <me|email|name|id>|--no-assign]
         [--priority <urgent|important|medium|low|0-10>] [--due <date>] [--start <date>]
         [--description <text|@file|->] [--checklist <text|@file|->]
         [--percent <0-100>]
  done <id...>
  delete <id...> --yes
  api <GET|POST|PUT|PATCH|DELETE> <path> [key=value...] [--data <JSON|@file|->] [--header 'Name: value'] [--all]

Common options:
  --json           Print the raw Graph response (read commands only)
  --all            Follow pages up to the safety cap
  --cursor <token> Continue from a previous page
  -h, --help     Show help
  -v, --version  Show version

Environment:
  PLANNER_PROFILE, PLANNER_TOKEN, PLANNER_CLIENT_ID, PLANNER_CONFIG, PLANNER_PLAN
  PLANNER_GRAPH_URL and PLANNER_LOGIN_URL (for alternate services and tests)

Exit codes:
  0  Success
  1  Server or authentication error
  2  Usage error`;
