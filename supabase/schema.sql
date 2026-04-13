create extension if not exists pgcrypto;

create table if not exists public.printmore_users (
  id uuid primary key default gen_random_uuid(),
  username text not null unique,
  password text not null,
  role text not null default 'user' check (role in ('user', 'designer', 'super')),
  active boolean not null default true,
  is_super_user boolean not null default false,
  created_at timestamptz not null default now()
);

alter table public.printmore_users enable row level security;

alter table public.printmore_users
add column if not exists password text;

alter table public.printmore_users
add column if not exists role text not null default 'user';

alter table public.printmore_users
add column if not exists active boolean not null default true;

update public.printmore_users
set password = 'More400'
where lower(username) = 'moraksh';

update public.printmore_users
set username = upper(username);

update public.printmore_users
set password = 'changeme'
where password is null;

alter table public.printmore_users
alter column password set not null;

alter table public.printmore_users
drop column if exists password_hash;

insert into public.printmore_users (username, password, role, active, is_super_user)
values ('MORAKSH', 'More400', 'super', true, true)
on conflict (username) do update
set password = excluded.password,
    role = 'super',
    active = true,
    is_super_user = true;

create table if not exists public.layouts (
  id text primary key,
  user_id uuid references public.printmore_users(id) on delete cascade,
  name text not null,
  layout jsonb not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.layouts
add column if not exists user_id uuid references public.printmore_users(id) on delete cascade;

update public.layouts
set user_id = (select id from public.printmore_users where username = 'MORAKSH')
where user_id is null;

alter table public.layouts
alter column user_id set not null;

alter table public.layouts enable row level security;

drop policy if exists "Allow public layout reads" on public.layouts;
drop policy if exists "Allow public layout inserts" on public.layouts;
drop policy if exists "Allow public layout updates" on public.layouts;
drop policy if exists "Allow public layout deletes" on public.layouts;
drop policy if exists "Allow app layout reads" on public.layouts;
drop policy if exists "Allow app layout inserts" on public.layouts;
drop policy if exists "Allow app layout updates" on public.layouts;
drop policy if exists "Allow app layout deletes" on public.layouts;

create policy "Allow app layout reads"
on public.layouts
for select
to anon
using (true);

create policy "Allow app layout inserts"
on public.layouts
for insert
to anon
with check (true);

create policy "Allow app layout updates"
on public.layouts
for update
to anon
using (true)
with check (true);

create policy "Allow app layout deletes"
on public.layouts
for delete
to anon
using (true);

create index if not exists layouts_user_updated_at_idx
on public.layouts (user_id, updated_at desc);

create table if not exists public.app_settings (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.app_settings enable row level security;

drop policy if exists "Allow app settings reads" on public.app_settings;
drop policy if exists "Allow app settings writes" on public.app_settings;

create policy "Allow app settings reads"
on public.app_settings
for select
to anon
using (true);

create policy "Allow app settings writes"
on public.app_settings
for all
to anon
using (true)
with check (true);

drop function if exists public.authenticate_printmore_user(text, text);
drop function if exists public.create_printmore_user(text, text, text, text);
drop function if exists public.create_printmore_user(text, text, text, text, text);
drop function if exists public.reset_printmore_user_password(text, text, text, text);
drop function if exists public.set_printmore_user_active(text, text, text, boolean);

create or replace function public.authenticate_printmore_user(
  p_username text,
  p_password text
)
returns table(id uuid, username text, role text, is_super_user boolean)
language sql
security definer
set search_path = public
as $$
  select u.id, upper(u.username), u.role, u.is_super_user
  from public.printmore_users u
  where lower(u.username) = lower(trim(p_username))
    and u.password = p_password
    and u.active = true
  limit 1;
$$;

create or replace function public.create_printmore_user(
  p_admin_username text,
  p_admin_password text,
  p_username text,
  p_password text,
  p_role text default 'user'
)
returns table(id uuid, username text, role text, is_super_user boolean)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin public.printmore_users;
  v_user public.printmore_users;
begin
  select *
  into v_admin
  from public.printmore_users u
  where lower(u.username) = lower(trim(p_admin_username))
    and u.password = p_admin_password
    and u.is_super_user = true
  limit 1;

  if v_admin.id is null then
    raise exception 'Invalid super user password.';
  end if;

  if trim(coalesce(p_username, '')) = '' or coalesce(p_password, '') = '' then
    raise exception 'User id and password are required.';
  end if;

  insert into public.printmore_users (username, password, role, active, is_super_user)
  values (upper(trim(p_username)), p_password, coalesce(nullif(p_role, ''), 'user'), true, coalesce(nullif(p_role, ''), 'user') = 'super')
  returning * into v_user;

  return query select v_user.id, upper(v_user.username), v_user.role, v_user.is_super_user;
end;
$$;

create or replace function public.reset_printmore_user_password(
  p_admin_username text,
  p_admin_password text,
  p_username text,
  p_new_password text
)
returns table(id uuid, username text, role text, is_super_user boolean)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin public.printmore_users;
  v_user public.printmore_users;
begin
  select *
  into v_admin
  from public.printmore_users u
  where lower(u.username) = lower(trim(p_admin_username))
    and u.password = p_admin_password
    and u.is_super_user = true
  limit 1;

  if v_admin.id is null then
    raise exception 'Invalid super user password.';
  end if;

  if trim(coalesce(p_username, '')) = '' or coalesce(p_new_password, '') = '' then
    raise exception 'User id and new password are required.';
  end if;

  update public.printmore_users u
  set password = p_new_password
  where lower(u.username) = lower(trim(p_username))
  returning * into v_user;

  if v_user.id is null then
    raise exception 'User id not found.';
  end if;

  return query select v_user.id, upper(v_user.username), v_user.role, v_user.is_super_user;
end;
$$;

create or replace function public.set_printmore_user_active(
  p_admin_username text,
  p_admin_password text,
  p_username text,
  p_active boolean
)
returns table(id uuid, username text, role text, is_super_user boolean)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin public.printmore_users;
  v_user public.printmore_users;
begin
  select *
  into v_admin
  from public.printmore_users u
  where lower(u.username) = lower(trim(p_admin_username))
    and u.password = p_admin_password
    and u.is_super_user = true
  limit 1;

  if v_admin.id is null then
    raise exception 'Invalid super user password.';
  end if;

  update public.printmore_users u
  set active = p_active
  where lower(u.username) = lower(trim(p_username))
  returning * into v_user;

  if v_user.id is null then
    raise exception 'User id not found.';
  end if;

  return query select v_user.id, upper(v_user.username), v_user.role, v_user.is_super_user;
end;
$$;

grant execute on function public.authenticate_printmore_user(text, text) to anon;
grant execute on function public.create_printmore_user(text, text, text, text, text) to anon;
grant execute on function public.reset_printmore_user_password(text, text, text, text) to anon;
grant execute on function public.set_printmore_user_active(text, text, text, boolean) to anon;
