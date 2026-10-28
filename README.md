# Planner CLI

A command-line client for Microsoft Planner. The CLI is bundled into one
self-contained script and uses Node.js built-in APIs at runtime.

## Development

```sh
pnpm install
pnpm typecheck
pnpm test
pnpm build
```

Tests run the built script in a child process with an isolated home directory
and local fake Graph and login servers. They do not contact Microsoft services.

See `GLOSSARY.md` for the project's domain terms.
