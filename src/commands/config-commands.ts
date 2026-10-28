import { CliError, usage } from '../utils/errors.js';
import { maskSecrets, saveConfig } from '../utils/config.js';
import { resolveBucketId, resolvePlanId } from '../utils/name-resolution.js';
import type { Context } from './auth-commands.js';

export async function runConfig(ctx: Context, args: string[]): Promise<void> {
  const [subcommand, key, ...values] = args;
  if (subcommand === 'show' && args.length === 1) {
    process.stdout.write(`${JSON.stringify(maskSecrets(ctx.config), null, 2)}\n`);
    return;
  }
  if (subcommand !== 'set' || args.length < 3) usage("Use 'planner config set plan|bucket <name|id>' or 'planner config show'.");
  if (key !== 'plan' && key !== 'bucket') usage("Config key must be 'plan' or 'bucket'.");
  const value = values.join(' ').trim();
  if (!value) throw new CliError('A Plan or Bucket name/id is required.', 2);
  if (key === 'plan') ctx.profile.defaultPlan = await resolvePlanId(ctx, value);
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
  process.stdout.write(`Default ${key === 'plan' ? 'Plan' : 'Bucket'} set to ${value} for Profile '${ctx.profileName}'.\n`);
}
