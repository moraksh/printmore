# PrintMore

A static browser app for designing printable layouts and generating PDFs from pasted or manually entered runtime data.

For first-time code walkthrough, start with:

- `CODEBASE_GUIDE.md`

## What Is Stored

Stored in Supabase:

- User ids and password hashes (not plaintext)
- Layout name
- Page settings
- Field definitions
- Designer elements
- Images/logos embedded in the layout

Not stored in Supabase:

- Values pasted from Excel while running a layout
- Manual values typed into the Run/Preview forms
- Generated PDF output

Runtime data is kept in browser memory only long enough to render the preview or PDF.

## Local Run

Open `index.html` directly in a browser, or serve the folder with any static file server.

Without Supabase settings, layouts are stored in browser `localStorage`.

## PDF Engine (V2)

PrintMore now runs on a single V2 PDF pipeline with profile-based output modes:

- `Draft (Small)`
- `Standard (Balanced)`
- `Print HD (Sharp)`

V2 rendering rules:

- text/lines/tables: vector drawing via jsPDF primitives
- images/logos: raster with profile-based JPEG control
- barcodes: high-contrast PNG rendering

Fallback safety:

- If `Print HD` exceeds size/time gates, it auto-falls back to `Standard` with user notice.

Email size guard:

- If generated PDF exceeds configured thresholds, email send is blocked with clear options.

Central config for all V2 PDF behavior:

- `js/pdf-config.js`

Telemetry helpers in browser console:

- `getLastPdfTelemetry()`
- `getPdfTelemetry()`
- `evaluatePdfReleaseGate()`

## Supabase Setup

1. Create a Supabase project.
2. Open the SQL editor.
3. Run `supabase/schema.sql`.
4. Copy your Project URL and anon/publishable key.
5. Put them in `js/config.js` (single place for server/database settings):

```js
window.PRINT_LAYOUT_CONFIG = {
  DATABASE: {
    PROVIDER: 'supabase',
    SUPABASE: {
      URL: 'https://your-project.supabase.co',
      ANON_KEY: 'your-anon-or-publishable-key',
      LAYOUTS_TABLE: 'layouts',
    },
  },
};
```

Never put a `service_role` key in `js/config.js`.

For future SQL Server migration, keep using `js/config.js` as the single control point:

- Change `DATABASE.PROVIDER` to `sqlserver`
- Set `DATABASE.SQLSERVER.API_BASE_URL` and related values
- Keep frontend unchanged; integrate backend API implementation separately

The SQL creates the super user:

- User ID: `moraksh`
- Password: `More400`

Only this super user sees the `Add User` button. From that modal, the super user can create users and reset any user's password. New user ids are unique, and each saved layout is tagged with the logged-in user's id so the UI shows only that user's layouts.

The included app-level login is suitable for a small controlled tool. For stricter public production security, move to Supabase Auth with owner-based RLS policies.

## Vercel Deploy

This is a static site, so Vercel can deploy it without a build step.

Recommended settings:

- Framework preset: `Other`
- Build command: leave empty
- Output directory: `.`

## Git

Initial commit example:

```powershell
git add .
git commit -m "Prepare PrintMore for Supabase and Vercel"
```

Then add your Git remote and push:

```powershell
git remote add origin https://github.com/YOUR_USER/YOUR_REPO.git
git push -u origin master
```

## V2 Quality Gate Script

Golden baseline and sample metrics are provided under `tests/`.

Run release gate:

```powershell
npm run pdf:v2:gate
```

Optional explicit paths:

```powershell
node scripts/pdf-v2-release-gate-check.js tests/pdf-v2-golden-baseline.json tests/pdf-v2-current-metrics.sample.json
```
