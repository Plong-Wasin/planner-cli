# Planner CLI

`planner` is a command-line client for Microsoft Planner. It talks to Microsoft Graph directly, uses device-code sign-in, stores credentials per Profile, and prints compact tables by default. The CLI is bundled into one self-contained script and requires Node.js 20 or later.

## Install

Install the CI-built script from the `release` branch with npm:

```sh
npm install -g --install-links git+https://github.com/Plong-Wasin/planner-cli.git#release
```

The `main` branch workflow runs typecheck, tests, and build before force-publishing the source and built `dist/planner.js` to `release`. The branch is created by the first successful workflow run. Re-running the install command updates the CLI.

## Sign in and configure a Profile

```sh
planner auth login
planner auth check
planner plans
planner config set plan "Example Plan"
planner buckets
planner config set bucket "To do"
planner config show
```

`auth login` prints a device code and verification URL. Complete sign-in in a browser; the CLI stores the refresh token in its private config file. `auth check` verifies the current Profile, `auth logout` removes its saved credentials, and `me` displays the signed-in identity.

A Profile contains credentials, a Default Plan, and a Default Bucket. Select one with `-p` or `--profile`, or set `PLANNER_PROFILE`. Profiles are useful for separate accounts or tenants:

```sh
planner -p personal auth login
planner -p work auth login
planner -p personal plans
```

For Plan and Bucket selection, command flags override environment variables, which override Profile defaults. `--client-id` overrides `PLANNER_CLIENT_ID`. `PLANNER_TOKEN` supplies a bearer token for headless use without a Profile. `PLANNER_CONFIG` selects another config file. `PLANNER_GRAPH_URL` and `PLANNER_LOGIN_URL` can point to alternate endpoints, such as local test servers.

## Commands

Run `planner --help` for the same command and option reference.

| Command | Purpose and common options |
|---|---|
| `planner auth login\|check\|logout` | Sign in, verify the selected Profile, or remove its credentials. |
| `planner config show` | Show Profiles and defaults with secrets masked. |
| `planner config set plan <name\|id>` | Set the Profile's Default Plan. |
| `planner config set bucket <name\|id>` | Set the Default Bucket in the Default Plan. |
| `planner me` | Show the signed-in user. |
| `planner plans` | List reachable Plans. |
| `planner buckets [--plan <name\|id>]` | List Buckets in the selected Plan. |
| `planner members [--plan <name\|id>]` | List members of the selected Plan's owning group. |
| `planner tasks` | List compact Task rows. Supports `--plan`, `--bucket`, `--mine`, `--open`, `--created-after`, `--completed-after`, `--fields`, `--details`, `--cursor`, `--all`, and `--json`. |
| `planner task <id\|url>` | Show one Task with Details. |
| `planner add <title>` | Create a Task in the Default Bucket, assigned to you at `important` Priority by default. Supports `--bucket`, `--assign`, `--no-assign`, `--priority`, `--due`, `--start`, `--description`, `--checklist`, and `--json`. |
| `planner add --data <@file\|->` | Bulk create Tasks from NDJSON or a JSON array. Input fields are `title`, `bucket`, `assign`, `priority`, `due`, `start`, `description`, and `checklist`. |
| `planner update <id...>` | Update one or more Tasks. Supports `--bucket`, `--assign`, `--no-assign`, `--priority`, `--due`, `--start`, `--description`, `--checklist`, and `--percent`. |
| `planner done <id...>` | Mark each Task complete. |
| `planner delete <id...> --yes` | Delete Tasks. Requires the explicit `--yes` flag. |
| `planner api <METHOD> <path>` | Send a raw Graph request with optional `key=value` query pairs, `--data <JSON\|@file\|->`, `--header 'Name: value'` (PATCH/DELETE under `/planner/` fetch the etag automatically), and GET-only `--all`. |

Names can be used for Plans, Buckets, and Assignees; ids and pasted Planner Task URLs are also accepted. Ambiguous or unknown names are rejected. Dates accept `YYYY-MM-DD` or a full ISO timestamp. `--description` and `--checklist` accept literal text, `@file`, or `-` for stdin; checklist text uses one item per line. `--json` prints the raw Graph response. Lists show a continuation cursor on stderr when another page is available; use `--cursor` to continue or `--all` to follow pages up to the safety cap.

## Why these shortcuts exist

The shortcuts follow observed usage: 415 calls across eight Planner operations. Task creation accounted for 212 calls, listing for 114, and reading one Task for 58; update appeared 15 times. The single-user and single-Plan pattern also supports assigning to `me` and storing Plan and Bucket defaults. `delete` was not observed and remains guarded by `--yes`.

## Agent plugin

The Claude Code and Codex marketplace manifests are generated from `plugins/planner-cli/plugin.json`; both plugins load the same `planner` skill.

Claude Code:

```text
/plugin marketplace add Plong-Wasin/planner-cli
/plugin install planner-cli@planner-cli
```

Codex:

```sh
codex plugin marketplace add https://github.com/Plong-Wasin/planner-cli.git
codex plugin add planner-cli
```

After installing the plugin, install the CLI using the command above. The skill checks `planner --version` and the selected Profile before use.

## Exit codes

- `0`: command succeeded.
- `1`: server or authentication error.
- `2`: usage error, including refusing a delete without `--yes`.

Data is written to stdout and progress notes and errors to stderr. Errors start with `planner:` and include a correction hint where available. Graph errors for 401, 403, 404, and 429 identify the method and path.

## Development

```sh
pnpm install --frozen-lockfile
pnpm typecheck
pnpm test
pnpm build
pnpm plugin:sync
```

Tests run the built script as a child process with an isolated home directory and in-process fake Graph and login servers. They do not contact Microsoft services or use a real config file. `pnpm plugin:sync` regenerates the Claude Code and Codex marketplace/plugin manifests from their shared source.
