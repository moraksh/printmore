# PrintMore

A static browser app for designing printable layouts and generating PDFs from pasted or manually entered runtime data.

## What Is Stored

Stored in Supabase:

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

## Supabase Setup

1. Create a Supabase project.
2. Open the SQL editor.
3. Run `supabase/schema.sql`.
4. Copy your Project URL and anon/publishable key.
5. Put them in `js/config.js`:

```js
window.PRINT_LAYOUT_CONFIG = {
  SUPABASE_URL: 'https://your-project.supabase.co',
  SUPABASE_ANON_KEY: 'your-anon-or-publishable-key',
  SUPABASE_LAYOUTS_TABLE: 'layouts',
};
```

Never put a `service_role` key in `js/config.js`.

The included SQL allows public anonymous read/write access to layouts. That is convenient for a shared internal tool, but anyone with the deployed app URL can change layouts. Add Supabase Auth and owner-based RLS before using it for private or multi-user production work.

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
