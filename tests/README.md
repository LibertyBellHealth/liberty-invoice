# Smoke tests

Fast, dependency-light checks on the critical data-integrity logic in `../app.js`
(DHS-1210 parsing, authorization time/date math, the client dirty-signature that
decides whether a save fires, localStorage round-trips, HTML escaping). They load
the **real** `app.js` in jsdom and call its actual functions — no copied logic.

## Run

```bash
cd tests
npm install   # first time only (installs jsdom, dev-only, gitignored)
npm test
```

## Why it lives here (not the repo root)

The live site is deployed by Azure Static Web Apps with `app_location: "/"` and no
build command, so a `package.json` at the **repo root** would make Azure treat the
static site as a Node app and try to build it. Keeping the test setup self-contained
in this folder leaves the root untouched, so the site deploys exactly as before.
`node_modules/` is gitignored; the folder is 404'd on the live site via
`staticwebapp.config.json`.

## Adding tests

New `*.test.js` files here are picked up automatically by `node --test`.
Get the app's functions via `const w = require('./harness').loadApp();` then call
`w.functionName(...)`. Note: values returned from app.js live in the jsdom realm —
compare primitives directly, and use `{...obj}` / `JSON.stringify` (or the
`jsonEqual` helper) for objects/arrays rather than `deepStrictEqual`.
