export class CliError extends Error {
  readonly exitCode: number;
  readonly status?: number;

  constructor(message: string, exitCode = 1, status?: number) {
    super(message);
    this.exitCode = exitCode;
    this.status = status;
  }
}

export function usage(message: string): never {
  throw new CliError(message, 2);
}
