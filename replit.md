# Weave — Webpage to Workbook

Weave turns public webpages into structured, editable XLSX workbooks.

## Run & Operate

- `pnpm --filter @workspace/api-server run dev` — run the API server (port 5000)
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from the OpenAPI spec
- `pnpm --filter @workspace/db run push` — push DB schema changes (dev only)
- Required env: `DATABASE_URL` — Postgres connection string

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- API: Express 5
- DB: PostgreSQL + Drizzle ORM
- Validation: Zod (`zod/v4`), `drizzle-zod`
- API codegen: Orval (from OpenAPI spec)
- Build: esbuild (CJS bundle)

## Where things live

- `artifacts/weave-workbook/src/pages/workspace.tsx` — primary workbook workspace and client-side XLSX generation.
- `artifacts/weave-workbook/src/index.css` — editorial research-desk theme, typography, motion, and responsive styling.
- `artifacts/api-server/src/routes/extract.ts` — safe public page fetching and HTML/JSON-LD extraction heuristics.
- `lib/api-spec/openapi.yaml` — source of truth for the extraction API contract.
- `lib/api-client-react/src/generated/` and `lib/api-zod/src/generated/` — generated API hooks and server validators.

## Architecture decisions

- Extraction is server-side so the browser can read public pages without relying on cross-origin browser permissions.
- The extractor uses deterministic HTML tables, JSON-LD, repeated article cards, and page metadata rather than fabricating content.
- XLSX files are generated in the browser and local saves use browser storage; raw page contents are not persisted server-side.
- Public URL fetching validates redirects and blocks private or loopback addresses to reduce SSRF risk.

## Product

- Paste a public URL and choose a row limit.
- Review suggested sheets with confidence, source trail, and page metadata.
- Edit sheet names, headers, cells, and rows; append another page as additional tabs.
- Save the current workbook locally or download a real `.xlsx` file.
- Start a fresh workbook without leaving the workspace.

## User preferences

_Populate as you build — explicit user instructions worth remembering across sessions._

## Gotchas

_Populate as you build — sharp edges, "always run X before Y" rules._

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
