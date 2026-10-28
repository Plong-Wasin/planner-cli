#!/usr/bin/env node

// src/help.ts
var HELP = `Usage: planner <command> [options]
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

// package.json
var package_default = {
  name: "planner-cli",
  version: "0.1.0",
  description: "Command-line client for Microsoft Planner",
  type: "module",
  bin: {
    planner: "./dist/planner.js"
  },
  scripts: {
    typecheck: "tsc --noEmit && tsc --noEmit -p tsconfig.test.json",
    build: 'pnpm typecheck && esbuild src/index.ts --bundle --platform=node --target=node20 --format=esm --banner:js="#!/usr/bin/env node" --outfile=dist/planner.js && chmod +x dist/planner.js',
    test: 'pnpm build && node --test "test/*.test.ts"',
    "plugin:sync": "node scripts/sync-plugin-manifests.mjs"
  },
  engines: {
    node: ">=20"
  },
  devDependencies: {
    "@types/node": "^20.19.0",
    esbuild: "^0.28.2",
    typescript: "^7.0.2"
  },
  private: true,
  license: "MIT",
  packageManager: "pnpm@12.3.4"
};

// src/utils/version.ts
var VERSION = package_default.version;

// src/utils/errors.ts
var CliError = class extends Error {
  exitCode;
  status;
  constructor(message, exitCode = 1, status) {
    super(message);
    this.exitCode = exitCode;
    this.status = status;
  }
};
function usage(message) {
  throw new CliError(message, 2);
}

// src/utils/config.ts
import { chmodSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { homedir } from "node:os";
function configPath(env) {
  return env.PLANNER_CONFIG || `${homedir()}/.config/planner/config.json`;
}
function loadConfig(path) {
  try {
    const value = JSON.parse(readFileSync(path, "utf8"));
    if (!value.profiles || typeof value.profiles !== "object") return { profiles: {} };
    return { defaultProfile: value.defaultProfile, profiles: value.profiles };
  } catch (error) {
    if (error.code === "ENOENT") return { profiles: {} };
    throw new CliError(`Cannot read config file ${path}: ${error.message}`);
  }
}
function saveConfig(path, config) {
  mkdirSync(dirname(path), { recursive: true, mode: 448 });
  writeFileSync(path, `${JSON.stringify(config, null, 2)}
`, { mode: 384 });
  chmodSync(path, 384);
}
function maskSecrets(config) {
  return {
    ...config,
    profiles: Object.fromEntries(Object.entries(config.profiles).map(([name, profile]) => [name, {
      ...profile,
      ...profile.accessToken ? { accessToken: "********" } : {},
      ...profile.refreshToken ? { refreshToken: "********" } : {}
    }]))
  };
}

// src/commands/auth-commands.ts
var DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
var SCOPES = "Group.Read.All Tasks.ReadWrite GroupMember.Read.All Directory.Read.All offline_access";
function context(env, options) {
  const path = configPath(env);
  const config = loadConfig(path);
  const profileName = options.profile || env.PLANNER_PROFILE || config.defaultProfile || "default";
  config.profiles[profileName] ??= {};
  return { env, options, config, path, profileName, profile: config.profiles[profileName] };
}
function persist(ctx) {
  ctx.config.profiles[ctx.profileName] = ctx.profile;
  ctx.config.defaultProfile ??= ctx.profileName;
  saveConfig(ctx.path, ctx.config);
}
function clientIdFor(ctx) {
  return ctx.options.clientId || ctx.env.PLANNER_CLIENT_ID || ctx.profile.clientId || DEFAULT_CLIENT_ID;
}
function tokenEndpoint(ctx) {
  return `${(ctx.env.PLANNER_LOGIN_URL || "https://login.microsoftonline.com").replace(/\/$/, "")}/common/oauth2/v2.0/token`;
}
async function getAccessToken(ctx) {
  if (ctx.env.PLANNER_TOKEN) return ctx.env.PLANNER_TOKEN;
  if (!ctx.profile.accessToken) throw new CliError(`Profile '${ctx.profileName}' is not signed in. Run 'planner auth login'.`);
  if (ctx.profile.expiresAt && ctx.profile.expiresAt <= Date.now() + 6e4) {
    await refresh(ctx);
  }
  return ctx.profile.accessToken;
}
async function refresh(ctx) {
  if (!ctx.profile.refreshToken) throw new CliError(`Sign-in expired. Run 'planner auth login' to sign in again.`);
  const body = new URLSearchParams({
    client_id: clientIdFor(ctx),
    grant_type: "refresh_token",
    refresh_token: ctx.profile.refreshToken,
    scope: SCOPES
  });
  let response;
  try {
    response = await fetch(tokenEndpoint(ctx), { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body });
  } catch {
    throw new CliError(`Could not refresh sign-in. Run 'planner auth login' to sign in again.`);
  }
  const data = await response.json();
  if (!response.ok || !data.access_token) throw new CliError(`Sign-in expired. Run 'planner auth login' to sign in again.`);
  ctx.profile.accessToken = data.access_token;
  ctx.profile.refreshToken = data.refresh_token || ctx.profile.refreshToken;
  ctx.profile.expiresAt = Date.now() + (data.expires_in ?? 3600) * 1e3;
  persist(ctx);
}
var MAX_THROTTLE_RETRIES = 2;
var MAX_RETRY_WAIT_SECONDS = 30;
async function graph(ctx, method, path, options = {}, retried = false, throttled = 0) {
  const base = (ctx.env.PLANNER_GRAPH_URL || "https://graph.microsoft.com/v1.0").replace(/\/$/, "");
  const url = new URL(options.rawUrl || `${base}${path.startsWith("/") ? path : `/${path}`}`);
  options.query?.forEach((value, key) => url.searchParams.append(key, value));
  const token = await getAccessToken(ctx);
  let response;
  try {
    response = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        ...options.headers,
        ...options.body === void 0 ? {} : { "Content-Type": "application/json" }
      },
      ...options.body === void 0 ? {} : { body: options.body }
    });
  } catch {
    throw new CliError(`Cannot reach ${url} for ${method} ${path}. Check the network and Graph URL.`);
  }
  if (response.status === 401 && !retried && !ctx.env.PLANNER_TOKEN && ctx.profile.refreshToken) {
    await refresh(ctx);
    return graph(ctx, method, path, options, true, throttled);
  }
  const retryable = response.status === 429 || response.status === 503 && method !== "POST";
  if (retryable && throttled < MAX_THROTTLE_RETRIES) {
    const header = response.headers.get("retry-after");
    const wait = header === null ? Number.NaN : Number(header);
    const seconds = Number.isFinite(wait) && wait >= 0 ? Math.min(wait, MAX_RETRY_WAIT_SECONDS) : 2 ** throttled;
    await new Promise((resolve) => setTimeout(resolve, seconds * 1e3));
    return graph(ctx, method, path, options, retried, throttled + 1);
  }
  if (!response.ok) {
    const detail = await graphErrorDetail(response);
    const messages = {
      401: `Authentication failed for ${method} ${path}. Run 'planner auth login'.`,
      403: `Access denied for ${method} ${path}. Check that your account has the required Planner permissions.`,
      404: `Not found: ${method} ${path}. Check the id and path.`,
      429: `Rate limited for ${method} ${path}. Wait, then try again.`
    };
    throw new CliError(`${messages[response.status] || `HTTP ${response.status} for ${method} ${path}.`}${detail ? ` Graph said: ${detail}` : ""}`, 1, response.status);
  }
  const text = await response.text();
  return text ? JSON.parse(text) : void 0;
}
async function graphErrorDetail(response) {
  try {
    const body = JSON.parse(await response.text());
    if (typeof body.error === "string") return body.error;
    return [body.error?.code, body.error?.message].filter(Boolean).join(": ");
  } catch {
    return "";
  }
}
async function runAuth(ctx, subcommand) {
  if (subcommand === "login") return login(ctx);
  if (subcommand === "check") {
    const me = await graph(ctx, "GET", "/me");
    process.stdout.write(`Signed in as ${me.displayName || me.userPrincipalName || me.mail || "current user"}
`);
    return;
  }
  if (subcommand === "logout") {
    delete ctx.profile.accessToken;
    delete ctx.profile.refreshToken;
    delete ctx.profile.expiresAt;
    persist(ctx);
    process.stdout.write(`Signed out of Profile '${ctx.profileName}'.
`);
    return;
  }
  usage("Use 'planner auth login', 'planner auth check', or 'planner auth logout'.");
}
async function login(ctx) {
  const clientId = clientIdFor(ctx);
  const base = (ctx.env.PLANNER_LOGIN_URL || "https://login.microsoftonline.com").replace(/\/$/, "");
  const deviceUrl = `${base}/common/oauth2/v2.0/devicecode`;
  let deviceResponse;
  try {
    deviceResponse = await fetch(deviceUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: clientId, scope: SCOPES })
    });
  } catch {
    throw new CliError(`Cannot reach ${deviceUrl}. Check the login URL and network.`);
  }
  const device = await deviceResponse.json();
  if (!deviceResponse.ok || !device.device_code) throw new CliError("Could not start device-code sign-in. Check the client id and try again.");
  process.stdout.write(`To sign in, visit ${device.verification_uri} and enter code ${device.user_code}.
`);
  const interval = Math.max(1, device.interval || 5) * 1e3;
  const expiresAt = Date.now() + (device.expires_in || 900) * 1e3;
  while (Date.now() < expiresAt) {
    const response = await fetch(tokenEndpoint(ctx), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: clientId, grant_type: "urn:ietf:params:oauth:grant-type:device_code", device_code: device.device_code })
    });
    const data = await response.json();
    if (response.ok && data.access_token) {
      ctx.profile.accessToken = data.access_token;
      ctx.profile.refreshToken = data.refresh_token;
      ctx.profile.clientId = clientId;
      ctx.profile.expiresAt = Date.now() + (data.expires_in ?? 3600) * 1e3;
      persist(ctx);
      process.stdout.write(`Signed in to Profile '${ctx.profileName}'.
`);
      return;
    }
    if (data.error === "authorization_pending") {
      await new Promise((resolve) => setTimeout(resolve, interval));
      continue;
    }
    if (data.error === "slow_down") {
      await new Promise((resolve) => setTimeout(resolve, interval + 5e3));
      continue;
    }
    if (data.error === "authorization_declined" || data.error === "expired_token") throw new CliError(`Device-code sign-in ${data.error === "expired_token" ? "expired" : "was declined"}. Run 'planner auth login' to try again.`);
    throw new CliError(`Device-code sign-in failed. Run 'planner auth login' to try again.`);
  }
  throw new CliError(`Device code expired. Run 'planner auth login' to try again.`);
}
async function graphRequest(ctx, method, path, options) {
  return graph(ctx, method, path, options);
}
function selectedContext(env, options) {
  return context(env, options);
}

// src/utils/name-resolution.ts
var PAGE_CAP = 100;
var lookups = /* @__PURE__ */ new WeakMap();
function cached(ctx, key, load) {
  let entries = lookups.get(ctx);
  if (!entries) lookups.set(ctx, entries = /* @__PURE__ */ new Map());
  let entry = entries.get(key);
  if (!entry) {
    entry = load();
    entries.set(key, entry);
    entry.catch(() => entries.delete(key));
  }
  return entry;
}
function isGuid(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}
function looksLikePlannerId(value) {
  return /^[A-Za-z0-9_-]{28}$/.test(value);
}
async function resolvePlanId(ctx, value) {
  if (isGuid(value)) return value;
  const plans = await collection(ctx, "/me/planner/plans");
  const matches = plans.filter((plan) => plan.id === value || equal(plan.title ?? plan.name, value));
  if (matches.length === 1 && matches[0]?.id) return matches[0].id;
  if (!matches.length && looksLikePlannerId(value)) return value;
  const choices = (matches.length ? matches : plans).map((plan) => `${plan.title ?? plan.name ?? "Untitled"} (${plan.id ?? "unknown id"})`);
  fail("Plan", value, matches.length > 1, choices);
}
async function resolveBucketId(ctx, plan, value) {
  const planId = await resolvePlanId(ctx, plan);
  if (isGuid(value)) return value;
  const buckets = await collection(ctx, `/planner/plans/${encodeURIComponent(planId)}/buckets`);
  const matches = buckets.filter((bucket) => bucket.id === value || equal(bucket.name, value));
  if (matches.length === 1 && matches[0]?.id) return matches[0].id;
  if (!matches.length && looksLikePlannerId(value)) return value;
  fail("Bucket", value, matches.length > 1, (matches.length ? matches : buckets).map((bucket) => `${bucket.name ?? "Untitled"} (${bucket.id ?? "unknown id"})`));
}
async function resolveAssigneeId(ctx, plan, value) {
  if (equal(value, "me")) {
    const me = await cached(ctx, "me", () => graphRequest(ctx, "GET", "/me"));
    if (me.id) return me.id;
    usage("Could not identify the current user. Check the signed-in Profile.");
  }
  if (isGuid(value)) return value;
  const planId = await resolvePlanId(ctx, plan);
  const selectedPlan = await cached(ctx, `plan:${planId}`, () => graphRequest(ctx, "GET", `/planner/plans/${encodeURIComponent(planId)}`));
  if (!selectedPlan.owner) usage(`Could not find the owning group for Plan '${plan}'. Check the Plan id.`);
  const members = await planMembers(ctx, selectedPlan);
  const matches = members.filter((member) => member.id === value || equal(member.displayName, value) || equal(member.mail, value) || equal(member.userPrincipalName, value));
  if (matches.length === 1 && matches[0]?.id) return matches[0].id;
  fail("Assignee", value, matches.length > 1, (matches.length ? matches : members).map((member) => `${member.displayName ?? "Unknown"} <${member.mail ?? member.userPrincipalName ?? "no email"}> (${member.id ?? "unknown id"})`));
}
function isPersonalPlan(plan) {
  return /\/users\/[^/]+\/?$/i.test(plan.container?.url ?? "");
}
function personalPlanMember(ctx, owner) {
  return cached(ctx, `user:${owner}`, () => graphRequest(ctx, "GET", `/users/${encodeURIComponent(owner)}?$select=id,displayName,mail,userPrincipalName`));
}
async function planMembers(ctx, plan) {
  if (!plan.owner) return [];
  if (isPersonalPlan(plan)) return [await personalPlanMember(ctx, plan.owner)];
  try {
    return await collection(ctx, `/groups/${encodeURIComponent(plan.owner)}/members`);
  } catch (error) {
    if (error instanceof CliError && error.status === 404) return [await personalPlanMember(ctx, plan.owner)];
    throw error;
  }
}
function collection(ctx, path) {
  return cached(ctx, `collection:${path}`, () => loadCollection(ctx, path));
}
async function loadCollection(ctx, path) {
  const values = [];
  let next = path;
  for (let page = 0; page < PAGE_CAP; page += 1) {
    const response = await graphPage(ctx, next);
    values.push(...response.value ?? []);
    if (!response["@odata.nextLink"]) return values;
    next = response["@odata.nextLink"];
    validateNext(next, ctx);
  }
  return values;
}
function graphPage(ctx, url) {
  return /^https?:\/\//i.test(url) ? graphRequest(ctx, "GET", url, { rawUrl: url }) : graphRequest(ctx, "GET", url);
}
function validateNext(value, ctx) {
  let link;
  try {
    link = new URL(value, ctx.env.PLANNER_GRAPH_URL || "https://graph.microsoft.com/v1.0");
  } catch {
    usage("Graph returned an invalid continuation URL while resolving a name.");
  }
  const base = new URL(ctx.env.PLANNER_GRAPH_URL || "https://graph.microsoft.com/v1.0");
  const basePath = base.pathname.replace(/\/$/, "");
  const inBasePath = !basePath || link.pathname === basePath || link.pathname.startsWith(`${basePath}/`);
  if (link.origin !== base.origin || !inBasePath) usage("Graph returned a continuation URL outside the configured Graph URL.");
}
function equal(left, right) {
  return (left ?? "").trim().toLocaleLowerCase() === right.trim().toLocaleLowerCase();
}
function fail(subject, value, ambiguous, candidates) {
  usage(`${ambiguous ? "Ambiguous" : "Unknown"} ${subject} '${value}'. Candidates: ${candidates.join(", ") || "none"}.`);
}

// src/commands/config-commands.ts
async function runConfig(ctx, args) {
  const [subcommand, key, ...values] = args;
  if (subcommand === "show" && args.length === 1) {
    process.stdout.write(`${JSON.stringify(maskSecrets(ctx.config), null, 2)}
`);
    return;
  }
  if (subcommand !== "set" || args.length < 3) usage("Use 'planner config set plan|bucket <name|id>' or 'planner config show'.");
  if (key !== "plan" && key !== "bucket") usage("Config key must be 'plan' or 'bucket'.");
  const value = values.join(" ").trim();
  if (!value) throw new CliError("A Plan or Bucket name/id is required.", 2);
  if (key === "plan") ctx.profile.defaultPlan = await resolvePlanId(ctx, value);
  else {
    const selectedPlan = ctx.options.plan ?? ctx.env.PLANNER_PLAN ?? ctx.profile.defaultPlan;
    if (!selectedPlan) throw new CliError("A Default Plan is required to set a Default Bucket. Run 'planner config set plan <name|id>' first.", 2);
    const planId = await resolvePlanId(ctx, selectedPlan);
    ctx.profile.defaultPlan = planId;
    ctx.profile.defaultBucket = await resolveBucketId(ctx, planId, value);
  }
  ctx.config.profiles[ctx.profileName] = ctx.profile;
  ctx.config.defaultProfile ??= ctx.profileName;
  saveConfig(ctx.path, ctx.config);
  process.stdout.write(`Default ${key === "plan" ? "Plan" : "Bucket"} set to ${value} for Profile '${ctx.profileName}'.
`);
}

// src/commands/read-commands.ts
var ALL_PAGE_CAP = 100;
function table(headers, rows) {
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values) => values.map((value, index) => value.padEnd(widths[index])).join("  ").trimEnd();
  return `${line(headers)}
${rows.map(line).join("\n")}${rows.length ? "\n" : ""}`;
}
async function runReadCommand(ctx, command, args) {
  const requestedPlan = command === "buckets" || command === "members" ? option(args, "--plan") ?? ctx.options.plan ?? ctx.env.PLANNER_PLAN ?? ctx.profile.defaultPlan : void 0;
  if (command === "buckets" || command === "members") {
    if (!requestedPlan) usage("A Default Plan is required. Set one with 'planner config set plan <name|id>' or pass --plan <name|id>.");
  }
  const selectedPlan = requestedPlan && !isGuid(requestedPlan) ? await resolvePlanId(ctx, requestedPlan) : requestedPlan;
  let response;
  let rows;
  let nextLink;
  if (command === "me") response = await graphRequest(ctx, "GET", "/me");
  else if (command === "buckets") {
    const result = await readCollection(ctx, `/planner/plans/${encodeURIComponent(selectedPlan)}/buckets`, args);
    response = result.response;
    rows = result.items;
    nextLink = result.nextLink;
  } else if (command === "members") {
    const plan = await graphRequest(ctx, "GET", `/planner/plans/${encodeURIComponent(selectedPlan)}`);
    if (!plan.owner) usage(`Could not find the owning group for Plan '${selectedPlan}'. Check the Plan id.`);
    const owner = plan.owner;
    const ownerOnly = async () => {
      const member = await personalPlanMember(ctx, owner);
      return { response: { value: [member] }, items: [member] };
    };
    let result;
    if (isPersonalPlan(plan)) result = await ownerOnly();
    else {
      try {
        result = await readCollection(ctx, `/groups/${encodeURIComponent(owner)}/members`, args);
      } catch (error) {
        if (!(error instanceof CliError && error.status === 404) || option(args, "--cursor")) throw error;
        result = await ownerOnly();
      }
    }
    response = result.response;
    rows = result.items;
    nextLink = result.nextLink;
  } else if (command === "plans") {
    const result = await readCollection(ctx, "/me/planner/plans", args);
    response = result.response;
    rows = result.items;
    nextLink = result.nextLink;
  } else usage(`Unknown read command: ${command}`);
  if (args.includes("--json")) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}
`);
    printContinuation(nextLink);
    return;
  }
  if (command === "me") {
    const user = response;
    process.stdout.write(table(["NAME", "EMAIL"], [[user.displayName ?? "Current user", user.mail ?? user.userPrincipalName ?? ""]]));
    return;
  }
  if (command === "members") {
    const members = rows ?? response.value ?? [];
    process.stdout.write(table(["ID", "NAME", "EMAIL"], members.map((member) => [member.id ?? "", member.displayName ?? "", member.mail ?? member.userPrincipalName ?? ""])));
    printContinuation(nextLink);
    return;
  }
  if (command === "buckets") {
    const buckets = rows ?? response.value ?? [];
    process.stdout.write(table(["ID", "NAME"], buckets.map((bucket) => [bucket.id ?? "", bucket.name ?? ""])));
    printContinuation(nextLink);
    return;
  }
  if (command !== "plans") throw new Error(`Unknown read command: ${command}`);
  const plans = rows ?? response.value ?? [];
  process.stdout.write(table(["ID", "NAME"], plans.map((plan) => [plan.id ?? "", plan.title ?? plan.name ?? ""])));
  printContinuation(nextLink);
}
async function readCollection(ctx, path, args) {
  const token = option(args, "--cursor");
  let url = token ? decodeCursor(token, ctx) : path;
  const all = args.includes("--all");
  const pages = [];
  let nextLink;
  for (let page = 0; page < (all ? ALL_PAGE_CAP : 1); page += 1) {
    const result = await graphPage2(ctx, url);
    pages.push(result);
    nextLink = result["@odata.nextLink"];
    if (!all || !nextLink) break;
    validateContinuation(nextLink, ctx);
    url = nextLink;
  }
  const first = pages[0] ?? { value: [] };
  const items = pages.flatMap((page) => page.value ?? []);
  const response = all ? { ...first, value: items, ...nextLink ? { "@odata.nextLink": nextLink } : {} } : first;
  if (all && !nextLink) delete response["@odata.nextLink"];
  return { response, items, nextLink };
}
function graphPage2(ctx, url) {
  return /^https?:\/\//i.test(url) ? graphRequest(ctx, "GET", url, { rawUrl: url }) : graphRequest(ctx, "GET", url);
}
function graphOrigin(ctx) {
  return new URL(ctx.env.PLANNER_GRAPH_URL || "https://graph.microsoft.com/v1.0").origin;
}
function validateContinuation(link, ctx) {
  let url;
  try {
    url = new URL(link, ctx.env.PLANNER_GRAPH_URL || "https://graph.microsoft.com/v1.0");
  } catch {
    usage("Graph returned an invalid continuation URL.");
  }
  const basePath = new URL(ctx.env.PLANNER_GRAPH_URL || "https://graph.microsoft.com/v1.0").pathname.replace(/\/$/, "");
  const inBasePath = !basePath || url.pathname === basePath || url.pathname.startsWith(`${basePath}/`);
  if (url.origin !== graphOrigin(ctx) || !inBasePath) usage("Graph returned a continuation URL outside the configured Graph URL.");
}
function decodeCursor(token, ctx) {
  let link;
  try {
    link = Buffer.from(token, "base64url").toString("utf8");
  } catch {
    usage("Invalid cursor. Copy the cursor from the previous command output.");
  }
  validateContinuation(link, ctx);
  return link;
}
function printContinuation(nextLink) {
  if (nextLink) process.stderr.write(`more available: --cursor ${Buffer.from(nextLink).toString("base64url")} (or --all)
`);
}
function option(args, name) {
  const index = args.indexOf(name);
  if (index < 0) return void 0;
  const value = args[index + 1];
  if (!value || value.startsWith("-")) usage(`Option '${name}' requires a value.`);
  return value;
}

// src/commands/task-read-commands.ts
var DEFAULT_PAGE_SIZE = 50;
var PAGE_CAP2 = 100;
var TITLE_WIDTH = 48;
var FIELD_MAP = {
  id: "id",
  title: "title",
  bucket: "bucketId",
  percent: "percentComplete",
  due: "dueDateTime",
  assignees: "assignments",
  created: "createdDateTime",
  completed: "completedDateTime",
  priority: "priority",
  start: "startDateTime"
};
var DEFAULT_FIELDS = ["id", "title", "bucket", "percent", "due", "assignees"];
async function runTaskReadCommand(ctx, command, args) {
  if (command === "task") return showTask(ctx, args);
  return listTasks(ctx, args);
}
async function listTasks(ctx, args) {
  const planInput = option2(args, "--plan") ?? ctx.options.plan ?? ctx.env.PLANNER_PLAN ?? ctx.profile.defaultPlan;
  const planId = planInput ? await resolvePlanId(ctx, planInput) : void 0;
  const bucketInput = option2(args, "--bucket");
  const bucketId = bucketInput ? isGuid(bucketInput) ? bucketInput : planId ? await resolveBucketId(ctx, planId, bucketInput) : usage("A Default Plan is required to resolve a Bucket name; pass --plan <name|id>.") : void 0;
  const fieldsArg = option2(args, "--fields");
  const fields = fieldsArg ? fieldsArg.split(",").map((field) => field.trim()).filter(Boolean) : DEFAULT_FIELDS;
  if (!fields.length || fields.some((field) => !Object.hasOwn(FIELD_MAP, field))) {
    usage(`Unknown field in --fields. Choose: ${Object.keys(FIELD_MAP).join(", ")}.`);
  }
  const selected = new Set(fields.map((field) => FIELD_MAP[field]));
  if (bucketId) selected.add("bucketId");
  if (args.includes("--mine")) selected.add("assignments");
  if (args.includes("--open")) selected.add("percentComplete");
  if (option2(args, "--created-after")) selected.add("createdDateTime");
  if (option2(args, "--completed-after")) selected.add("completedDateTime");
  const query = new URLSearchParams();
  query.set("$top", String(DEFAULT_PAGE_SIZE));
  query.set("$select", [...selected].join(","));
  if (args.includes("--details")) query.set("$expand", "details");
  const basePath = planId ? `/planner/plans/${encodeURIComponent(planId)}/tasks` : "/me/planner/tasks";
  const cursor = option2(args, "--cursor");
  let url = cursor ? decodeCursor2(cursor, ctx) : `${basePath}?${query}`;
  const all = args.includes("--all");
  const createdAfter = option2(args, "--created-after");
  const createdBound = createdAfter ? dateBound(createdAfter, "--created-after") : void 0;
  const completedAfter = option2(args, "--completed-after");
  const completedBound = completedAfter ? dateBound(completedAfter, "--completed-after") : void 0;
  let meId;
  if (args.includes("--mine")) {
    const me = await graphRequest(ctx, "GET", "/me");
    if (!me.id) usage("Could not identify the current user for --mine.");
    meId = me.id;
  }
  const keep = (task) => (meId === void 0 || Object.hasOwn(task.assignments ?? {}, meId)) && (!args.includes("--open") || (task.percentComplete ?? 0) < 100) && (!bucketId || task.bucketId === bucketId) && (createdBound === void 0 || after(task.createdDateTime, createdBound)) && (completedBound === void 0 || after(task.completedDateTime, completedBound));
  const filtered = meId !== void 0 || args.includes("--open") || bucketId !== void 0 || createdBound !== void 0 || completedBound !== void 0;
  const pages = [];
  const tasks = [];
  let nextLink;
  for (let page = 0; page < (all || filtered ? PAGE_CAP2 : 1); page += 1) {
    const result = await graphPage3(ctx, url);
    pages.push(result);
    tasks.push(...(result.value ?? []).filter(keep));
    nextLink = result["@odata.nextLink"];
    if (!nextLink || !all && tasks.length >= DEFAULT_PAGE_SIZE) break;
    if (!all && !filtered) break;
    validateContinuation2(nextLink, ctx);
    url = nextLink;
  }
  const first = pages[0] ?? { value: [] };
  const response = all ? { ...first, value: tasks, ...nextLink ? { "@odata.nextLink": nextLink } : {} } : { ...first, value: tasks };
  if (all && !nextLink) delete response["@odata.nextLink"];
  if (args.includes("--json")) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}
`);
    printContinuation2(nextLink);
    return;
  }
  const bucketNames = /* @__PURE__ */ new Map();
  const assigneeNames = /* @__PURE__ */ new Map();
  if (planId && fields.includes("bucket") && tasks.some((task) => task.bucketId)) {
    for (const bucket of await collection(ctx, `/planner/plans/${encodeURIComponent(planId)}/buckets`)) if (bucket.id) bucketNames.set(bucket.id, bucket.name ?? bucket.id);
  }
  if (planId && fields.includes("assignees") && tasks.some((task) => Object.keys(task.assignments ?? {}).length)) {
    const plan = await graphRequest(ctx, "GET", `/planner/plans/${encodeURIComponent(planId)}`);
    for (const member of await planMembers(ctx, plan)) if (member.id) assigneeNames.set(member.id, member.displayName ?? member.mail ?? member.userPrincipalName ?? member.id);
  }
  const detailFields = args.includes("--details") ? ["description", "checklist"] : [];
  const displayFields = [...fields, ...detailFields];
  const labels = displayFields.map((field) => field.toUpperCase());
  const rows = tasks.map((task) => displayFields.map((field) => formatField(task, field, bucketNames, assigneeNames)));
  process.stdout.write(table2(labels, rows));
  printContinuation2(nextLink);
}
async function showTask(ctx, args) {
  const idInput = args.find((arg) => !arg.startsWith("-"));
  if (!idInput) usage("Usage: planner task <id|url>.");
  const id = parseTaskId(idInput);
  const task = await graphRequest(ctx, "GET", `/planner/tasks/${encodeURIComponent(id)}?$expand=details`);
  if (args.includes("--json")) {
    process.stdout.write(`${JSON.stringify(task, null, 2)}
`);
    return;
  }
  const planId = typeof task.planId === "string" ? task.planId : void 0;
  const bucketNames = /* @__PURE__ */ new Map();
  const assigneeNames = /* @__PURE__ */ new Map();
  if (planId && task.bucketId) {
    for (const bucket of await collection(ctx, `/planner/plans/${encodeURIComponent(planId)}/buckets`)) if (bucket.id) bucketNames.set(bucket.id, bucket.name ?? bucket.id);
  }
  if (planId && Object.keys(task.assignments ?? {}).length) {
    const plan = await graphRequest(ctx, "GET", `/planner/plans/${encodeURIComponent(planId)}`);
    for (const member of await planMembers(ctx, plan)) {
      if (member.id) assigneeNames.set(member.id, member.displayName ?? member.mail ?? member.userPrincipalName ?? member.id);
    }
  }
  const lines = ["id", "title", "bucket", "percent", "priority", "start", "due", "assignees", "description", "checklist"].map((field) => [field, formatField(task, field, bucketNames, assigneeNames, false)]).filter(([, value]) => value !== "").map(([field, value]) => `${field}: ${value}`);
  process.stdout.write(`${lines.join("\n")}
`);
}
function isTaskId(value) {
  return /^[A-Za-z0-9_-]{28}$/.test(value) || isGuid(value);
}
function parseTaskId(value) {
  if (isTaskId(value)) return value;
  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    usage(`Invalid Task id '${value}'. Pass a Planner Task id or Planner Task URL.`);
  }
  const parts = parsed.pathname.split("/").filter(Boolean);
  const candidate = parts.at(-1) ?? "";
  if (!/^https?:$/.test(parsed.protocol) || !isTaskId(candidate)) usage(`Invalid Task id '${value}'. Pass a Planner Task id or Planner Task URL.`);
  return candidate;
}
function formatField(task, field, buckets, members, truncate = true) {
  if (field === "title") {
    const title = task.title ?? "";
    return truncate && title.length > TITLE_WIDTH ? `${title.slice(0, TITLE_WIDTH - 1)}\u2026` : title;
  }
  if (field === "bucket") return buckets.get(task.bucketId ?? "") ?? task.bucketId ?? "";
  if (field === "percent") return task.percentComplete === void 0 ? "" : `${task.percentComplete}%`;
  if (field === "due") return task.dueDateTime?.slice(0, 10) ?? "";
  if (field === "start") return task.startDateTime?.slice(0, 10) ?? "";
  if (field === "assignees") return Object.keys(task.assignments ?? {}).map((id) => members.get(id) ?? id).join(", ");
  if (field === "description") return task.details?.description ?? "";
  if (field === "checklist") return Object.values(task.details?.checklist ?? {}).map((item) => `${item.isChecked ? "[x]" : "[ ]"} ${item.title ?? ""}`).join("; ");
  const value = task[FIELD_MAP[field]];
  return value === void 0 || value === null ? "" : String(value);
}
function table2(headers, rows) {
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values) => values.map((value, index) => value.padEnd(widths[index])).join("  ").trimEnd();
  return `${line(headers)}
${rows.map(line).join("\n")}${rows.length ? "\n" : ""}`;
}
function dateBound(date, flag) {
  const match = /^(\d{4}-\d{2}-\d{2})$/.exec(date);
  const target = match ? Date.parse(`${date}T00:00:00Z`) : Number.NaN;
  if (!Number.isFinite(target) || new Date(target).toISOString().slice(0, 10) !== date) usage(`Option '${flag}' requires a valid YYYY-MM-DD date.`);
  return target;
}
function after(value, target) {
  if (!value) return false;
  const timestamp = Date.parse(value);
  if (!Number.isFinite(timestamp)) return false;
  return timestamp > target;
}
function option2(args, name) {
  const index = args.indexOf(name);
  if (index < 0) return void 0;
  const value = args[index + 1];
  if (!value || value.startsWith("-")) usage(`Option '${name}' requires a value.`);
  return value;
}
function graphPage3(ctx, url) {
  return /^https?:\/\//i.test(url) ? graphRequest(ctx, "GET", url, { rawUrl: url }) : graphRequest(ctx, "GET", url);
}
function decodeCursor2(token, ctx) {
  let link;
  try {
    link = Buffer.from(token, "base64url").toString("utf8");
  } catch {
    usage("Invalid cursor. Copy the cursor from the previous command output.");
  }
  validateContinuation2(link, ctx);
  return link;
}
function validateContinuation2(link, ctx) {
  let url;
  try {
    url = new URL(link, ctx.env.PLANNER_GRAPH_URL || "https://graph.microsoft.com/v1.0");
  } catch {
    usage("Graph returned an invalid continuation URL.");
  }
  const base = new URL(ctx.env.PLANNER_GRAPH_URL || "https://graph.microsoft.com/v1.0");
  const basePath = base.pathname.replace(/\/$/, "");
  if (url.origin !== base.origin || basePath && url.pathname !== basePath && !url.pathname.startsWith(`${basePath}/`)) {
    usage("Graph returned a continuation URL outside the configured Graph URL.");
  }
}
function printContinuation2(nextLink) {
  if (nextLink) process.stderr.write(`more available: --cursor ${Buffer.from(nextLink).toString("base64url")} (or --all)
`);
}

// src/commands/api-commands.ts
import { readFile } from "node:fs/promises";
async function runApi(ctx, args) {
  const method = args[0]?.toUpperCase();
  const path = args[1];
  if (!method || !path) usage("Use 'planner api <METHOD> <path> [key=value ...]'.");
  if (!["GET", "POST", "PUT", "PATCH", "DELETE"].includes(method)) usage(`Unsupported HTTP method: ${args[0]}.`);
  const query = new URLSearchParams();
  let body;
  let all = false;
  const headers = {};
  for (let index = 2; index < args.length; index += 1) {
    const pair = args[index];
    if (pair === "--all") {
      all = true;
      continue;
    }
    if (pair === "--header" || pair.startsWith("--header=")) {
      const header = pair === "--header" ? args[++index] : pair.slice("--header=".length);
      const colon = header?.indexOf(":") ?? -1;
      if (!header || colon <= 0) usage("Option '--header' requires 'Name: value'.");
      headers[header.slice(0, colon).trim()] = header.slice(colon + 1).trim();
      continue;
    }
    if (pair === "--data" || pair.startsWith("--data=")) {
      if (body !== void 0) usage("Option '--data' may only be provided once.");
      const source = pair === "--data" ? args[++index] : pair.slice("--data=".length);
      if (source === void 0) usage("Option '--data' requires JSON, @file, or '-'.");
      const text = source === "-" ? await readStdin() : source.startsWith("@") ? await readFile(source.slice(1), "utf8") : source;
      try {
        body = JSON.stringify(JSON.parse(text));
      } catch {
        usage("Option '--data' must contain valid JSON.");
      }
      continue;
    }
    const separator = pair.indexOf("=");
    if (separator <= 0) usage(`Expected a query pair in key=value form: ${pair}.`);
    query.append(pair.slice(0, separator), pair.slice(separator + 1));
  }
  if (all && method !== "GET") usage("Option '--all' is only supported with GET requests.");
  const hasIfMatch = Object.keys(headers).some((name) => name.toLowerCase() === "if-match");
  if ((method === "PATCH" || method === "DELETE") && path.startsWith("/planner/") && !hasIfMatch) {
    const current = await graphRequest(ctx, "GET", path);
    if (current?.["@odata.etag"]) headers["If-Match"] = current["@odata.etag"];
  }
  let data = await graphRequest(ctx, method, path, { query, body, headers });
  if (all) data = await collectPages(ctx, path, query, data);
  if (data !== void 0) process.stdout.write(`${JSON.stringify(data)}
`);
}
async function collectPages(ctx, path, query, initial) {
  if (!isRecord(initial) || !Array.isArray(initial.value)) return initial;
  const values = [...initial.value];
  const result = { ...initial, value: values };
  delete result["@odata.nextLink"];
  let nextLink = initial["@odata.nextLink"];
  for (let page = 1; typeof nextLink === "string" && page < 50; page += 1) {
    const nextUrl = new URL(nextLink);
    const base = new URL(ctx.env.PLANNER_GRAPH_URL || "https://graph.microsoft.com/v1.0");
    const basePath = base.pathname.replace(/\/$/, "");
    const inBasePath = !basePath || nextUrl.pathname === basePath || nextUrl.pathname.startsWith(`${basePath}/`);
    if (nextUrl.origin !== base.origin || !inBasePath) {
      usage("Graph returned a continuation URL outside the configured Graph URL.");
    }
    const data = await graphRequest(ctx, "GET", path, { rawUrl: nextUrl.href });
    if (!isRecord(data)) break;
    if (Array.isArray(data.value)) values.push(...data.value);
    nextLink = data["@odata.nextLink"];
  }
  return result;
}
function isRecord(value) {
  return typeof value === "object" && value !== null && Array.isArray(value.value);
}
async function readStdin() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString("utf8");
}

// src/commands/add-command.ts
import { randomUUID } from "node:crypto";
import { readFileSync as readFileSync2 } from "node:fs";
var PRIORITIES = { urgent: 1, important: 3, medium: 5, low: 9 };
async function runAdd(ctx, args) {
  const options = parseOptions(args);
  if (options.data !== void 0) {
    if (options.positionals.length || options.bucket !== void 0 || options.assign !== void 0 || options.noAssign || options.priority !== void 0 || options.due !== void 0 || options.start !== void 0 || options.description !== void 0 || options.checklist !== void 0) {
      usage("Use '--data' with fields on each input record, not Task flags or a title.");
    }
    const entries = bulkEntries(loadText(options.data, "--data"));
    let failed = false;
    for (const entry of entries) {
      try {
        if (entry.error !== void 0) throw new CliError(entry.error, 2);
        const record = entry.record;
        if (!record || Array.isArray(record) || typeof record !== "object") usage("Each Task record must be a JSON object.");
        const task = bulkOptions(record);
        await createTask(ctx, task.title, task);
      } catch (error) {
        failed = true;
        const message = error instanceof Error ? error.message : String(error);
        process.stderr.write(`line ${entry.line}: ${message}
`);
      }
    }
    if (failed) process.exitCode = 1;
    return;
  }
  const title = options.positionals[0];
  if (!title) usage("Usage: planner add <title> [--bucket <name|id>] [--assign <me|email|name|id>] [options]");
  if (options.positionals.length > 1) usage("Task titles with spaces must be quoted as one argument.");
  await createTask(ctx, title, options);
}
async function createTask(ctx, title, options) {
  if (!title) usage("Each Task record needs a non-empty 'title'.");
  if (options.noAssign && options.assign !== void 0) usage("Use either 'assign' or no assignment, not both.");
  const requestedPlan = ctx.options.plan || ctx.env.PLANNER_PLAN || ctx.profile.defaultPlan;
  if (!requestedPlan) usage("A Default Plan is required. Set one with 'planner config set plan <name|id>' or pass --plan <name|id>.");
  const planId = await resolvePlanId(ctx, requestedPlan);
  const requestedBucket = options.bucket ?? ctx.profile.defaultBucket;
  if (!requestedBucket) usage("A Default Bucket is required. Set one with 'planner config set bucket <name|id>' or pass --bucket <name|id>.");
  const bucketId = await resolveBucketId(ctx, planId, requestedBucket);
  const payload = {
    planId,
    bucketId,
    title,
    priority: options.priority === void 0 ? 3 : parsePriority(options.priority)
  };
  if (options.assign !== void 0 || !options.noAssign) {
    const assignee = await resolveAssigneeId(ctx, planId, options.assign ?? "me");
    payload.assignments = {
      [assignee]: { "@odata.type": "microsoft.graph.plannerAssignment", orderHint: " !" }
    };
  }
  if (options.due !== void 0) payload.dueDateTime = parseDate(options.due, "--due");
  if (options.start !== void 0) payload.startDateTime = parseDate(options.start, "--start");
  const task = await graphRequest(ctx, "POST", "/planner/tasks", { body: JSON.stringify(payload) });
  if (!task?.id) throw new CliError("Graph created the Task without returning its id.");
  const details = {};
  if (options.description !== void 0) details.description = options.literal ? options.description : loadText(options.description, "--description");
  if (options.checklist !== void 0) details.checklist = checklist(options.literal ? options.checklist : loadText(options.checklist, "--checklist"));
  process.stdout.write(`${task.id}
`);
  if (Object.keys(details).length) {
    try {
      await writeDetails(ctx, task.id, details);
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      throw new CliError(`Task ${task.id} was created but its Details were not saved: ${reason}`, error instanceof CliError ? error.exitCode : 1, error instanceof CliError ? error.status : void 0);
    }
  }
}
async function readNewDetails(ctx, path) {
  for (let attempt = 0; ; attempt += 1) {
    try {
      return await graphRequest(ctx, "GET", path);
    } catch (error) {
      if (!(error instanceof CliError) || error.status !== 404 || attempt >= 4) throw error;
      await new Promise((resolve) => setTimeout(resolve, 500));
    }
  }
}
async function writeDetails(ctx, taskId, body) {
  const path = `/planner/tasks/${encodeURIComponent(taskId)}/details`;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const current = await readNewDetails(ctx, path);
    if (!current?.["@odata.etag"]) throw new CliError(`Task Details for '${taskId}' did not include an etag.`);
    try {
      await graphRequest(ctx, "PATCH", path, {
        body: JSON.stringify(body),
        headers: { "If-Match": current["@odata.etag"] }
      });
      return;
    } catch (error) {
      if (!(error instanceof CliError) || error.status !== 412) throw error;
      if (attempt === 1) throw new CliError(`Task Details for '${taskId}' changed twice while being updated. Retry the command.`);
    }
  }
}
function parseOptions(args) {
  const parsed = { positionals: [], noAssign: false };
  const valued = {
    "--bucket": "bucket",
    "--assign": "assign",
    "--priority": "priority",
    "--due": "due",
    "--start": "start",
    "--description": "description",
    "--checklist": "checklist",
    "--data": "data"
  };
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === "--no-assign") {
      parsed.noAssign = true;
      continue;
    }
    if (arg === "--json") usage("Option '--json' is only supported by read commands.");
    const [name, inline] = arg.includes("=") ? [arg.slice(0, arg.indexOf("=")), arg.slice(arg.indexOf("=") + 1)] : [arg, void 0];
    const key = valued[name];
    if (key) {
      const value = inline ?? args[++index];
      if (value === void 0 || inline === void 0 && value.startsWith("--")) usage(`Option '${name}' requires a value.`);
      parsed[key] = value;
      continue;
    }
    if (arg.startsWith("-")) usage(`Unknown add option: ${arg}.`);
    parsed.positionals.push(arg);
  }
  if (parsed.noAssign && parsed.assign !== void 0) usage("Use either '--assign' or '--no-assign', not both.");
  return parsed;
}
function bulkEntries(input) {
  if (input.trimStart().startsWith("[")) {
    try {
      const value = JSON.parse(input);
      if (Array.isArray(value)) return value.map((record, index) => ({ line: index + 1, record }));
      return [{ line: 1, record: value }];
    } catch {
    }
  }
  const entries = [];
  input.split(/\r?\n/).forEach((line, index) => {
    if (!line.trim()) return;
    try {
      entries.push({ line: index + 1, record: JSON.parse(line) });
    } catch {
      entries.push({ line: index + 1, error: "Invalid JSON." });
    }
  });
  return entries;
}
function bulkOptions(record) {
  const allowed = /* @__PURE__ */ new Set(["title", "bucket", "assign", "priority", "due", "start", "description", "checklist"]);
  for (const key of Object.keys(record)) {
    if (!allowed.has(key)) usage(`Unknown Task field '${key}'.`);
  }
  const options = { literal: true };
  for (const key of allowed) {
    const value = record[key];
    if (value === void 0) continue;
    if (key === "checklist" && Array.isArray(value) && value.every((item) => typeof item === "string")) {
      options.checklist = value.join("\n");
      continue;
    }
    if (key === "priority" && typeof value === "number") {
      options.priority = String(value);
      continue;
    }
    if (typeof value !== "string") usage(`Task field '${key}' must be text${key === "checklist" ? " or a list of text" : key === "priority" ? " or a number" : ""}.`);
    if (key === "title") options.title = value;
    else if (key === "bucket") options.bucket = value;
    else if (key === "assign") options.assign = value;
    else if (key === "priority") options.priority = value;
    else if (key === "due") options.due = value;
    else if (key === "start") options.start = value;
    else if (key === "description") options.description = value;
    else if (key === "checklist") options.checklist = value;
  }
  return options;
}
function parsePriority(value) {
  const named = PRIORITIES[value.toLocaleLowerCase()];
  if (named !== void 0) return named;
  if (/^(?:10|[0-9])$/.test(value)) return Number(value);
  usage("Priority must be urgent, important, medium, low, or a number from 0 to 10.");
}
function parseDate(value, flag) {
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const date = /* @__PURE__ */ new Date(`${value}T00:00:00.000Z`);
    if (date.toISOString().slice(0, 10) === value) return date.toISOString();
  } else if (/^\d{4}-\d\d-\d\dT/.test(value) && Number.isFinite(Date.parse(value))) {
    return value;
  }
  usage(`Option '${flag}' must be YYYY-MM-DD or a full ISO date and time.`);
}
function loadText(value, flag) {
  if (value === "-") return readFileSync2(0, "utf8");
  if (value.startsWith("@")) {
    const path = value.slice(1);
    if (!path) usage(`Option '${flag}' needs a file path after '@'.`);
    try {
      return readFileSync2(path, "utf8");
    } catch {
      usage(`Cannot read input file for '${flag}': ${path}.`);
    }
  }
  return value;
}
function checklist(input) {
  return Object.fromEntries(input.split(/\r?\n/).filter((line) => line.length > 0).map((title) => [randomUUID(), {
    "@odata.type": "microsoft.graph.plannerChecklistItem",
    title,
    isChecked: false
  }]));
}

// src/commands/bulk-write-commands.ts
import { readFileSync as readFileSync3 } from "node:fs";
import { randomUUID as randomUUID2 } from "node:crypto";
var PRIORITIES2 = { urgent: 1, important: 3, medium: 5, low: 9 };
async function runBulkWrite(ctx, command, args) {
  const options = parseOptions2(command, args);
  const ids = options.positionals.map(parseTaskId);
  if (!ids.length) usage(`Usage: planner ${command} <id\u2026>${command === "delete" ? " --yes" : ""} [options]`);
  if (command === "delete" && !options.yes) usage("Refusing to delete Tasks without '--yes'.");
  const update = command === "update" ? prepareUpdate(options) : void 0;
  const failures = [];
  for (const id of ids) {
    try {
      if (command === "delete") await deleteTask(ctx, id);
      else if (command === "done") await updateTask(ctx, id, { percentComplete: 100 });
      else {
        const change = await updateBody(ctx, id, update);
        await updateTask(ctx, id, change.fields, change.details);
      }
      process.stdout.write(`${command === "done" ? "Completed" : command === "delete" ? "Deleted" : "Updated"} #${id}
`);
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      failures.push(id);
      process.stderr.write(`#${id}: ${reason}
`);
    }
  }
  if (failures.length) {
    process.stderr.write(`${failures.length} Task${failures.length === 1 ? "" : "s"} failed.
`);
    process.exitCode = 1;
  }
}
function prepareUpdate(options) {
  const fields = {};
  if (options.priority !== void 0) fields.priority = parsePriority2(options.priority);
  if (options.due !== void 0) fields.dueDateTime = parseDate2(options.due, "--due");
  if (options.start !== void 0) fields.startDateTime = parseDate2(options.start, "--start");
  if (options.percent !== void 0) fields.percentComplete = parsePercent(options.percent);
  const description = options.description === void 0 ? void 0 : loadText2(options.description, "--description");
  const checklistText = options.checklist === void 0 ? void 0 : loadText2(options.checklist, "--checklist");
  if (!Object.keys(fields).length && options.bucket === void 0 && options.assign === void 0 && !options.noAssign && description === void 0 && checklistText === void 0) usage("Provide at least one field to update.");
  return { fields, description, checklist: checklistText, bucket: options.bucket, assign: options.assign, noAssign: options.noAssign };
}
async function updateBody(ctx, id, update) {
  const fields = { ...update.fields };
  let task;
  const currentTask = async () => {
    task ??= await graphRequest(ctx, "GET", `/planner/tasks/${encodeURIComponent(id)}`);
    return task;
  };
  if (update.assign !== void 0 || update.noAssign || update.bucket !== void 0) {
    const current = await currentTask();
    if (update.assign !== void 0) {
      const planId = current.planId || ctx.options.plan || ctx.env.PLANNER_PLAN || ctx.profile.defaultPlan;
      if (!planId) usage("A Default Plan is required to resolve an Assignee. Set one with 'planner config set plan <name|id>'.");
      const assigneeId = await resolveAssigneeId(ctx, planId, update.assign);
      fields.assignments = { [assigneeId]: { "@odata.type": "microsoft.graph.plannerAssignment", orderHint: " !" } };
    } else if (update.noAssign) {
      fields.assignments = Object.fromEntries(Object.keys(current.assignments ?? {}).map((assigneeId) => [assigneeId, null]));
    }
    if (update.bucket !== void 0) {
      const planId = current.planId || ctx.options.plan || ctx.env.PLANNER_PLAN || ctx.profile.defaultPlan;
      if (!planId) usage("A Default Plan is required to move a Task. Set one with 'planner config set plan <name|id>'.");
      fields.bucketId = await resolveBucketId(ctx, planId, update.bucket);
    }
  }
  const details = {};
  if (update.description !== void 0) details.description = update.description;
  if (update.checklist !== void 0) details.checklist = checklist2(update.checklist);
  return { fields, ...Object.keys(details).length ? { details } : {} };
}
async function updateTask(ctx, id, fields, details) {
  if (Object.keys(fields).length) await writeTask(ctx, id, fields);
  if (details) await writeDetails2(ctx, id, details);
}
async function writeTask(ctx, id, body) {
  const path = `/planner/tasks/${encodeURIComponent(id)}`;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const task = await graphRequest(ctx, "GET", path);
    if (!task?.["@odata.etag"]) throw new CliError(`Task '${id}' did not include an etag.`);
    try {
      await graphRequest(ctx, "PATCH", path, {
        body: JSON.stringify(body),
        headers: { "If-Match": task["@odata.etag"] }
      });
      return;
    } catch (error) {
      if (!(error instanceof CliError) || error.status !== 412) throw error;
      if (attempt === 1) throw new CliError(`Task '${id}' changed twice while being updated. Retry the command.`);
    }
  }
}
async function deleteTask(ctx, id) {
  const path = `/planner/tasks/${encodeURIComponent(id)}`;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const task = await graphRequest(ctx, "GET", path);
    if (!task?.["@odata.etag"]) throw new CliError(`Task '${id}' did not include an etag.`);
    try {
      await graphRequest(ctx, "DELETE", path, { headers: { "If-Match": task["@odata.etag"] } });
      return;
    } catch (error) {
      if (!(error instanceof CliError) || error.status !== 412) throw error;
      if (attempt === 1) throw new CliError(`Task '${id}' changed twice while being deleted. Retry the command.`);
    }
  }
}
async function writeDetails2(ctx, id, body) {
  const path = `/planner/tasks/${encodeURIComponent(id)}/details`;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const current = await graphRequest(ctx, "GET", path);
    if (!current?.["@odata.etag"]) throw new CliError(`Task Details for '${id}' did not include an etag.`);
    try {
      await graphRequest(ctx, "PATCH", path, { body: JSON.stringify(body), headers: { "If-Match": current["@odata.etag"] } });
      return;
    } catch (error) {
      if (!(error instanceof CliError) || error.status !== 412) throw error;
      if (attempt === 1) throw new CliError(`Task Details for '${id}' changed twice while being updated. Retry the command.`);
    }
  }
}
function parseOptions2(command, args) {
  const parsed = { positionals: [], noAssign: false, yes: false };
  const valued = {
    "--bucket": "bucket",
    "--assign": "assign",
    "--priority": "priority",
    "--due": "due",
    "--start": "start",
    "--description": "description",
    "--checklist": "checklist",
    "--percent": "percent"
  };
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === "--yes" && command === "delete") {
      parsed.yes = true;
      continue;
    }
    if (arg === "--no-assign" && command === "update") {
      parsed.noAssign = true;
      continue;
    }
    if (arg === "--json") usage("Option '--json' is only supported by read commands.");
    const [name, inline] = arg.includes("=") ? [arg.slice(0, arg.indexOf("=")), arg.slice(arg.indexOf("=") + 1)] : [arg, void 0];
    const key = valued[name];
    if (key && command === "update") {
      const value = inline ?? args[++index];
      if (value === void 0 || inline === void 0 && value.startsWith("--")) usage(`Option '${name}' requires a value.`);
      parsed[key] = value;
      continue;
    }
    if (arg.startsWith("-")) usage(`Unknown ${command} option: ${arg}.`);
    parsed.positionals.push(arg);
  }
  if (parsed.noAssign && parsed.assign !== void 0) usage("Use either '--assign' or '--no-assign', not both.");
  return parsed;
}
function parsePriority2(value) {
  const named = PRIORITIES2[value.toLocaleLowerCase()];
  if (named !== void 0) return named;
  if (/^(?:10|[0-9])$/.test(value)) return Number(value);
  usage("Priority must be urgent, important, medium, low, or a number from 0 to 10.");
}
function parseDate2(value, flag) {
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const date = /* @__PURE__ */ new Date(`${value}T00:00:00.000Z`);
    if (date.toISOString().slice(0, 10) === value) return date.toISOString();
  } else if (/^\d{4}-\d\d-\d\dT/.test(value) && Number.isFinite(Date.parse(value))) return value;
  usage(`Option '${flag}' must be YYYY-MM-DD or a full ISO date and time.`);
}
function parsePercent(value) {
  if (/^(?:100|[0-9]{1,2})$/.test(value)) return Number(value);
  usage("Option '--percent' must be a whole number from 0 to 100.");
}
function loadText2(value, flag) {
  if (value === "-") return readFileSync3(0, "utf8");
  if (value.startsWith("@")) {
    const path = value.slice(1);
    if (!path) usage(`Option '${flag}' needs a file path after '@'.`);
    try {
      return readFileSync3(path, "utf8");
    } catch {
      usage(`Cannot read input file for '${flag}': ${path}.`);
    }
  }
  return value;
}
function checklist2(input) {
  return Object.fromEntries(input.split(/\r?\n/).filter((line) => line.length > 0).map((title) => [randomUUID2(), {
    "@odata.type": "microsoft.graph.plannerChecklistItem",
    title,
    isChecked: false
  }]));
}

// src/index.ts
var USAGE_ERROR = 2;
function globals(args) {
  const options = {};
  const rest = [];
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === "-p" || arg === "--profile" || arg === "--client-id" || arg === "--plan") {
      const value = args[index + 1];
      if (!value || value.startsWith("-")) usage(`Option '${arg}' requires a value.`);
      if (arg === "--client-id") options.clientId = value;
      else if (arg === "--plan") options.plan = value;
      else options.profile = value;
      index += 1;
    } else if (arg.startsWith("--profile=")) options.profile = arg.slice("--profile=".length);
    else if (arg.startsWith("--client-id=")) options.clientId = arg.slice("--client-id=".length);
    else if (arg.startsWith("--plan=")) options.plan = arg.slice("--plan=".length);
    else rest.push(arg);
  }
  return { options, rest };
}
async function main(args) {
  if (args.length === 0 || args.includes("--help") || args.includes("-h")) {
    process.stdout.write(`${HELP}
`);
    return;
  }
  if (args.length === 1 && (args[0] === "--version" || args[0] === "-v")) {
    process.stdout.write(`${VERSION}
`);
    return;
  }
  const { options, rest } = globals(args);
  const [command, subcommand, ...commandArgs] = rest;
  const ctx = selectedContext(process.env, options);
  if (command === "auth") return runAuth(ctx, subcommand);
  if (command === "config") return runConfig(ctx, [subcommand, ...commandArgs].filter((value) => value !== void 0));
  if (command === "api") return runApi(ctx, [subcommand, ...commandArgs].filter((value) => value !== void 0));
  if (command === "add") return runAdd(ctx, [subcommand, ...commandArgs].filter((value) => value !== void 0));
  if (command === "update" || command === "done" || command === "delete") return runBulkWrite(ctx, command, [subcommand, ...commandArgs].filter((value) => value !== void 0));
  if (command === "me" || command === "plans" || command === "buckets" || command === "members") return runReadCommand(ctx, command, [subcommand, ...commandArgs].filter((value) => value !== void 0));
  if (command === "tasks") return runTaskReadCommand(ctx, "tasks", [subcommand, ...commandArgs].filter((value) => value !== void 0));
  if (command === "task") return runTaskReadCommand(ctx, "task", [subcommand, ...commandArgs].filter((value) => value !== void 0));
  if (!command) usage("Run 'planner --help' for usage.");
  throw new CliError(`Unknown command: ${command}. Run 'planner --help' for usage.`, USAGE_ERROR);
}
main(process.argv.slice(2)).catch((error) => {
  const code = error instanceof CliError ? error.exitCode : 1;
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`planner: ${message}
`);
  process.exitCode = code;
});
