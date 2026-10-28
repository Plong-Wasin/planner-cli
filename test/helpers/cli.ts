import { spawn } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const BIN = new URL('../../dist/planner.js', import.meta.url).pathname;

export interface RunResult {
  code: number | null;
  stdout: string;
  stderr: string;
}

export interface Cli {
  home: string;
  run(args: string[], env?: Record<string, string>, stdin?: string): Promise<RunResult>;
  cleanup(): void;
}

export function createCli(env: Record<string, string> = {}): Cli {
  const home = mkdtempSync(join(tmpdir(), 'planner-cli-test-'));
  const configPath = join(home, 'config.json');

  return {
    home,
    run: (args, extraEnv = {}, stdin) => new Promise((resolve, reject) => {
      const child = spawn(process.execPath, [BIN, ...args], {
        // Intentionally do not spread process.env: developer settings must not leak in.
        env: {
          PATH: process.env.PATH ?? '',
          HOME: home,
          PLANNER_CONFIG: configPath,
          ...env,
          ...extraEnv,
        },
        stdio: [stdin === undefined ? 'ignore' : 'pipe', 'pipe', 'pipe'],
      });

      let stdout = '';
      let stderr = '';
      child.stdout?.setEncoding('utf8').on('data', (chunk: string) => { stdout += chunk; });
      child.stderr?.setEncoding('utf8').on('data', (chunk: string) => { stderr += chunk; });
      child.on('error', reject);
      child.on('close', (code) => resolve({ code, stdout, stderr }));
      if (stdin !== undefined) {
        child.stdin?.write(stdin);
        child.stdin?.end();
      }
    }),
    cleanup: () => rmSync(home, { recursive: true, force: true }),
  };
}
