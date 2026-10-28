## Agent skills

### Issue tracker

Issues live as local markdown under `.scratch/<feature>/` (gitignored). See `docs/agents/issue-tracker.md`.

### Triage labels

Default five-role vocabulary (`needs-triage`, `needs-info`, `ready-for-agent`, `ready-for-human`, `wontfix`). See `docs/agents/triage-labels.md`.

### Domain docs

Single-context: one `GLOSSARY.md` + `docs/adr/` at repo root. See `docs/agents/domain.md`.

### CLI documentation

When a command or flag changes, update `README.md` and `plugins/planner-cli/skills/planner/SKILL.md` in the same change. Keep the generated marketplace and plugin manifests in sync by running `pnpm plugin:sync`.
