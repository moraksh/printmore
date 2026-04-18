# PrintMore Codebase Guide

This guide is for developers opening the project for the first time.

## 1) Project Layout

- `index.html`
  Main single-page app shell and all view containers.
- `css/styles.css`
  Global styles for login, home, designer, run, modals, live preview.
- `js/app.js`
  SPA controller: routing between views, event wiring, run/live/share flows.
- `js/designer.js`
  Layout designer engine: canvas tools, element editing, table editing, properties.
- `js/pdf.js`
  V2 PDF + live preview rendering pipeline (shared canonical page planning).
- `js/storage.js`
  Layout persistence layer (Supabase RPC + local fallback).
- `js/auth.js`
  Login/session/user management client via Supabase RPC.
- `js/config.js`
  Environment and database provider configuration.
- `js/pdf-config.js`
  V2 PDF profile configuration and limits.
- `api/send-mail.js`
  Vercel serverless endpoint for sending generated PDFs via email.
- `supabase/schema.sql`
  SQL schema + RPCs for users, sessions, layouts, smart rules.

## 2) Core Runtime Flow

1. App boot: `DOMContentLoaded` in `js/app.js`
2. Session check via `AuthStore`
3. Layout load via `LayoutStore`
4. Home view render
5. User opens Designer or Run
6. Run/Live preview renders layout via `renderLayoutPreview()` in `js/pdf.js`
7. PDF export calls `generatePDF()` in `js/pdf.js`

## 3) Rendering Architecture (Important)

- The app is **V2-only** for PDF generation.
- Live preview and PDF share the same planner:
  - `_buildCanonicalRenderPlan(...)`
- This planner decides:
  - page size/orientation
  - header/footer inclusion by page
  - detail-table slicing for multi-page output
  - per-page element list and overrides

Using one planner is the main parity mechanism.

## 4) Data Persistence Boundaries

Persisted:
- Users, layouts, layout metadata, smart UI rules, email layout settings.

Not persisted:
- Runtime pasted/manual row data used during run/live generation.

## 5) Where to Change What

- Add/modify toolbar or view actions:
  - `js/app.js`
- Add a new designer element type:
  - `js/designer.js` + `js/pdf.js`
- Adjust PDF quality/size behavior:
  - `js/pdf-config.js` (profiles) + minor hooks in `js/pdf.js`
- Change DB provider details:
  - `js/config.js`
- Change auth/storage backend calls:
  - `js/auth.js`, `js/storage.js`, `supabase/schema.sql`

## 6) Safe Change Tips

- Keep rendering math in mm, convert to px only at render edges.
- When touching table/barcode logic, test all 3 surfaces:
  - Designer
  - Live preview
  - PDF output
- For multi-page fixes, always test with large detail rows (100+).

