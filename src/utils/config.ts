import { chmodSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { homedir } from 'node:os';
import { CliError } from './errors.js';

export interface Profile {
  accessToken?: string;
  refreshToken?: string;
  expiresAt?: number;
  defaultPlan?: string;
  defaultBucket?: string;
}

export interface Config {
  defaultProfile?: string;
  profiles: Record<string, Profile>;
}

export function configPath(env: NodeJS.ProcessEnv): string {
  return env.PLANNER_CONFIG || `${homedir()}/.config/planner/config.json`;
}

export function loadConfig(path: string): Config {
  try {
    const value = JSON.parse(readFileSync(path, 'utf8')) as Partial<Config>;
    if (!value.profiles || typeof value.profiles !== 'object') return { profiles: {} };
    return { defaultProfile: value.defaultProfile, profiles: value.profiles };
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return { profiles: {} };
    throw new CliError(`Cannot read config file ${path}: ${(error as Error).message}`);
  }
}

export function saveConfig(path: string, config: Config): void {
  mkdirSync(dirname(path), { recursive: true, mode: 0o700 });
  writeFileSync(path, `${JSON.stringify(config, null, 2)}\n`, { mode: 0o600 });
  chmodSync(path, 0o600);
}

export function maskSecrets(config: Config): Config {
  return {
    ...config,
    profiles: Object.fromEntries(Object.entries(config.profiles).map(([name, profile]) => [name, {
      ...profile,
      ...(profile.accessToken ? { accessToken: '********' } : {}),
      ...(profile.refreshToken ? { refreshToken: '********' } : {}),
    }])),
  };
}
