create table if not exists public.layouts (
  id text primary key,
  name text not null,
  layout jsonb not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.layouts enable row level security;

create policy "Allow public layout reads"
on public.layouts
for select
to anon
using (true);

create policy "Allow public layout inserts"
on public.layouts
for insert
to anon
with check (true);

create policy "Allow public layout updates"
on public.layouts
for update
to anon
using (true)
with check (true);

create policy "Allow public layout deletes"
on public.layouts
for delete
to anon
using (true);

create index if not exists layouts_updated_at_idx
on public.layouts (updated_at desc);
