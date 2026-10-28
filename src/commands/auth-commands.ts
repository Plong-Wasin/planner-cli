import { CliError, usage } from '../utils/errors.js';
import { configPath, loadConfig, saveConfig, type Config, type Profile } from '../utils/config.js';

const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const SCOPES = 'Group.Read.All Tasks.ReadWrite GroupMember.Read.All Directory.Read.All offline_access';

export interface GlobalOptions {
  profile?: string;
  clientId?: string;
}

export interface Context {
  env: NodeJS.ProcessEnv;
  options: GlobalOptions;
  config: Config;
  path: string;
  profileName: string;
  profile: Profile;
}

export function context(env: NodeJS.ProcessEnv, options: GlobalOptions): Context {
  const path = configPath(env);
  const config = loadConfig(path);
  const profileName = options.profile || env.PLANNER_PROFILE || config.defaultProfile || 'default';
  config.profiles[profileName] ??= {};
  return { env, options, config, path, profileName, profile: config.profiles[profileName]! };
}

export function persist(ctx: Context): void {
  ctx.config.profiles[ctx.profileName] = ctx.profile;
  ctx.config.defaultProfile ??= ctx.profileName;
  saveConfig(ctx.path, ctx.config);
}

export function tokenEndpoint(ctx: Context): string {
  return `${(ctx.env.PLANNER_LOGIN_URL || 'https://login.microsoftonline.com').replace(/\/$/, '')}/common/oauth2/v2.0/token`;
}

export async function getAccessToken(ctx: Context): Promise<string> {
  if (ctx.env.PLANNER_TOKEN) return ctx.env.PLANNER_TOKEN;
  if (!ctx.profile.accessToken) throw new CliError(`Profile '${ctx.profileName}' is not signed in. Run 'planner auth login'.`);
  if (ctx.profile.expiresAt && ctx.profile.expiresAt <= Date.now() + 60_000) {
    await refresh(ctx);
  }
  return ctx.profile.accessToken!;
}

async function refresh(ctx: Context): Promise<void> {
  if (!ctx.profile.refreshToken) throw new CliError(`Sign-in expired. Run 'planner auth login' to sign in again.`);
  const body = new URLSearchParams({
    client_id: ctx.options.clientId || ctx.env.PLANNER_CLIENT_ID || DEFAULT_CLIENT_ID,
    grant_type: 'refresh_token',
    refresh_token: ctx.profile.refreshToken,
    scope: SCOPES,
  });
  let response: Response;
  try {
    response = await fetch(tokenEndpoint(ctx), { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  } catch {
    throw new CliError(`Could not refresh sign-in. Run 'planner auth login' to sign in again.`);
  }
  const data = await response.json() as TokenResponse;
  if (!response.ok || !data.access_token) throw new CliError(`Sign-in expired. Run 'planner auth login' to sign in again.`);
  ctx.profile.accessToken = data.access_token;
  ctx.profile.refreshToken = data.refresh_token || ctx.profile.refreshToken;
  ctx.profile.expiresAt = Date.now() + (data.expires_in ?? 3600) * 1000;
  persist(ctx);
}

interface TokenResponse {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
}

export interface GraphRequestOptions {
  query?: URLSearchParams;
  body?: string;
  rawUrl?: string;
}

async function graph(ctx: Context, method: string, path: string, options: GraphRequestOptions = {}, retried = false): Promise<unknown> {
  const base = (ctx.env.PLANNER_GRAPH_URL || 'https://graph.microsoft.com/v1.0').replace(/\/$/, '');
  const url = new URL(options.rawUrl || `${base}${path.startsWith('/') ? path : `/${path}`}`);
  options.query?.forEach((value, key) => url.searchParams.append(key, value));
  const token = await getAccessToken(ctx);
  let response: Response;
  try {
    response = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        ...(options.body === undefined ? {} : { 'Content-Type': 'application/json' }),
      },
      ...(options.body === undefined ? {} : { body: options.body }),
    });
  } catch {
    throw new CliError(`Cannot reach ${url} for ${method} ${path}. Check the network and Graph URL.`);
  }
  if (response.status === 401 && !retried && !ctx.env.PLANNER_TOKEN && ctx.profile.refreshToken) {
    await refresh(ctx);
    return graph(ctx, method, path, options, true);
  }
  if (!response.ok) {
    const messages: Record<number, string> = {
      401: `Authentication failed for ${method} ${path}. Run 'planner auth login'.`,
      403: `Access denied for ${method} ${path}. Check that your account has the required Planner permissions.`,
      404: `Not found: ${method} ${path}. Check the id and path.`,
      429: `Rate limited for ${method} ${path}. Wait, then try again.`,
    };
    throw new CliError(messages[response.status] || `HTTP ${response.status} for ${method} ${path}.`);
  }
  const text = await response.text();
  return text ? JSON.parse(text) as unknown : undefined;
}

export async function runAuth(ctx: Context, subcommand: string | undefined): Promise<void> {
  if (subcommand === 'login') return login(ctx);
  if (subcommand === 'check') {
    const me = await graph(ctx, 'GET', '/me') as { displayName?: string; userPrincipalName?: string; mail?: string };
    process.stdout.write(`Signed in as ${me.displayName || me.userPrincipalName || me.mail || 'current user'}\n`);
    return;
  }
  if (subcommand === 'logout') {
    delete ctx.profile.accessToken;
    delete ctx.profile.refreshToken;
    delete ctx.profile.expiresAt;
    persist(ctx);
    process.stdout.write(`Signed out of Profile '${ctx.profileName}'.\n`);
    return;
  }
  usage("Use 'planner auth login', 'planner auth check', or 'planner auth logout'.");
}

async function login(ctx: Context): Promise<void> {
  const clientId = ctx.options.clientId || ctx.env.PLANNER_CLIENT_ID || DEFAULT_CLIENT_ID;
  const base = (ctx.env.PLANNER_LOGIN_URL || 'https://login.microsoftonline.com').replace(/\/$/, '');
  const deviceUrl = `${base}/common/oauth2/v2.0/devicecode`;
  let deviceResponse: Response;
  try {
    deviceResponse = await fetch(deviceUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: clientId, scope: SCOPES }),
    });
  } catch {
    throw new CliError(`Cannot reach ${deviceUrl}. Check the login URL and network.`);
  }
  const device = await deviceResponse.json() as DeviceCode;
  if (!deviceResponse.ok || !device.device_code) throw new CliError('Could not start device-code sign-in. Check the client id and try again.');
  process.stdout.write(`To sign in, visit ${device.verification_uri} and enter code ${device.user_code}.\n`);
  const interval = Math.max(1, device.interval || 5) * 1000;
  const expiresAt = Date.now() + (device.expires_in || 900) * 1000;
  while (Date.now() < expiresAt) {
    const response = await fetch(tokenEndpoint(ctx), {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: clientId, grant_type: 'urn:ietf:params:oauth:grant-type:device_code', device_code: device.device_code }),
    });
    const data = await response.json() as TokenResponse;
    if (response.ok && data.access_token) {
      ctx.profile.accessToken = data.access_token;
      ctx.profile.refreshToken = data.refresh_token;
      ctx.profile.expiresAt = Date.now() + (data.expires_in ?? 3600) * 1000;
      persist(ctx);
      process.stdout.write(`Signed in to Profile '${ctx.profileName}'.\n`);
      return;
    }
    if (data.error === 'authorization_pending') {
      await new Promise((resolve) => setTimeout(resolve, interval));
      continue;
    }
    if (data.error === 'slow_down') {
      await new Promise((resolve) => setTimeout(resolve, interval + 5000));
      continue;
    }
    if (data.error === 'authorization_declined' || data.error === 'expired_token') throw new CliError(`Device-code sign-in ${data.error === 'expired_token' ? 'expired' : 'was declined'}. Run 'planner auth login' to try again.`);
    throw new CliError(`Device-code sign-in failed. Run 'planner auth login' to try again.`);
  }
  throw new CliError(`Device code expired. Run 'planner auth login' to try again.`);
}

interface DeviceCode {
  device_code?: string;
  user_code?: string;
  verification_uri?: string;
  interval?: number;
  expires_in?: number;
}

export async function graphRequest(ctx: Context, method: string, path: string, options?: GraphRequestOptions): Promise<unknown> {
  return graph(ctx, method, path, options);
}

export function selectedContext(env: NodeJS.ProcessEnv, options: GlobalOptions): Context {
  return context(env, options);
}

export function configLocation(ctx: Context): string {
  return configPath(ctx.env);
}
