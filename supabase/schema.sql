create extension if not exists pgcrypto;

create table if not exists public.printmore_users (
  id uuid primary key default gen_random_uuid(),
  username text not null unique,
  full_name text not null default '',
  password text not null,
  role text not null default 'user' check (role in ('user', 'designer', 'super')),
  active boolean not null default true,
  is_super_user boolean not null default false,
  last_login_at timestamptz,
  created_at timestamptz not null default now()
);

alter table public.printmore_users enable row level security;

alter table public.printmore_users
add column if not exists password text;

alter table public.printmore_users
add column if not exists role text not null default 'user';

alter table public.printmore_users
add column if not exists active boolean not null default true;

alter table public.printmore_users
add column if not exists full_name text not null default '';

alter table public.printmore_users
add column if not exists last_login_at timestamptz;

update public.printmore_users
set password = 'More400'
where lower(username) = 'moraksh';

update public.printmore_users
set username = upper(username);

update public.printmore_users
set full_name = coalesce(nullif(trim(full_name), ''), username);

update public.printmore_users
set password = 'changeme'
where password is null;

alter table public.printmore_users
alter column password set not null;

alter table public.printmore_users
drop column if exists password_hash;

insert into public.printmore_users (username, full_name, password, role, active, is_super_user)
values ('MORAKSH', 'Akshay More', 'More400', 'super', true, true)
on conflict (username) do update
set password = excluded.password,
    full_name = 'Akshay More',
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
drop function if exists public.create_printmore_user(text, text, text, text, text, text);
drop function if exists public.reset_printmore_user_password(text, text, text, text);
drop function if exists public.set_printmore_user_active(text, text, text, boolean);
drop function if exists public.update_printmore_user(text, text, text, text, boolean);
drop function if exists public.find_printmore_user(text);
drop function if exists public.list_printmore_users(text, text);
drop function if exists public.delete_printmore_user(text, text, text);

create or replace function public.authenticate_printmore_user(
  p_username text,
  p_password text
)
returns table(id uuid, username text, full_name text, role text, is_super_user boolean)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user public.printmore_users;
begin
  update public.printmore_users u
  set last_login_at = now()
  where lower(u.username) = lower(trim(p_username))
    and u.password = p_password
    and u.active = true
  returning * into v_user;

  if v_user.id is null then
    return;
  end if;

  return query
  select v_user.id, upper(v_user.username), v_user.full_name, v_user.role, v_user.is_super_user;
end;
$$;

create or replace function public.create_printmore_user(
  p_admin_username text,
  p_admin_password text,
  p_username text,
  p_full_name text,
  p_password text,
  p_role text default 'user'
)
returns table(id uuid, username text, full_name text, role text, is_super_user boolean)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin public.printmore_users;
  v_user public.printmore_users;
  v_target_id uuid;
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

  insert into public.printmore_users (username, full_name, password, role, active, is_super_user)
  values (
    upper(trim(p_username)),
    coalesce(nullif(trim(p_full_name), ''), upper(trim(p_username))),
    p_password,
    coalesce(nullif(p_role, ''), 'user'),
    true,
    coalesce(nullif(p_role, ''), 'user') = 'super'
  )
  returning * into v_user;

  return query select v_user.id, upper(v_user.username), v_user.full_name, v_user.role, v_user.is_super_user;
end;
$$;

create or replace function public.reset_printmore_user_password(
  p_admin_username text,
  p_admin_password text,
  p_username text,
  p_new_password text
)
returns table(id uuid, username text, full_name text, role text, is_super_user boolean)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin public.printmore_users;
  v_user public.printmore_users;
  v_target_id uuid;
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

  select u.id
  into v_target_id
  from public.printmore_users u
  where lower(u.username) = lower(trim(p_username))
     or lower(u.full_name) = lower(trim(p_username))
  order by case when lower(u.username) = lower(trim(p_username)) then 0 else 1 end
  limit 1;

  update public.printmore_users u
  set password = p_new_password
  where u.id = v_target_id
  returning * into v_user;

  if v_user.id is null then
    raise exception 'User id not found.';
  end if;

  return query select v_user.id, upper(v_user.username), v_user.full_name, v_user.role, v_user.is_super_user;
end;
$$;

create or replace function public.set_printmore_user_active(
  p_admin_username text,
  p_admin_password text,
  p_username text,
  p_active boolean
)
returns table(id uuid, username text, full_name text, role text, is_super_user boolean)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin public.printmore_users;
  v_user public.printmore_users;
  v_target_id uuid;
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

  select u.id
  into v_target_id
  from public.printmore_users u
  where lower(u.username) = lower(trim(p_username))
     or lower(u.full_name) = lower(trim(p_username))
  order by case when lower(u.username) = lower(trim(p_username)) then 0 else 1 end
  limit 1;

  update public.printmore_users u
  set active = p_active
  where u.id = v_target_id
  returning * into v_user;

  if v_user.id is null then
    raise exception 'User id not found.';
  end if;

  return query select v_user.id, upper(v_user.username), v_user.full_name, v_user.role, v_user.is_super_user;
end;
$$;

create or replace function public.update_printmore_user(
  p_admin_username text,
  p_admin_password text,
  p_username text,
  p_role text,
  p_active boolean
)
returns table(id uuid, username text, full_name text, role text, is_super_user boolean)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin public.printmore_users;
  v_user public.printmore_users;
  v_role text;
  v_target_id uuid;
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

  v_role := coalesce(nullif(trim(p_role), ''), 'user');

  select u.id
  into v_target_id
  from public.printmore_users u
  where lower(u.username) = lower(trim(p_username))
     or lower(u.full_name) = lower(trim(p_username))
  order by case when lower(u.username) = lower(trim(p_username)) then 0 else 1 end
  limit 1;

  update public.printmore_users u
  set role = v_role,
      is_super_user = (v_role = 'super'),
      active = p_active
  where u.id = v_target_id
  returning * into v_user;

  if v_user.id is null then
    raise exception 'User id not found.';
  end if;

  return query select v_user.id, upper(v_user.username), v_user.full_name, v_user.role, v_user.is_super_user;
end;
$$;

create or replace function public.find_printmore_user(
  p_username text
)
returns table(id uuid, username text, full_name text, role text, is_super_user boolean)
language sql
security definer
set search_path = public
as $$
  select u.id, upper(u.username), u.full_name, u.role, u.is_super_user
  from public.printmore_users u
  where lower(u.username) = lower(trim(p_username))
    and u.active = true
  limit 1;
$$;

create or replace function public.list_printmore_users(
  p_admin_username text,
  p_admin_password text
)
returns table(id uuid, username text, full_name text, role text, is_super_user boolean, active boolean, last_login_at timestamptz)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin public.printmore_users;
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

  return query
  select
    u.id,
    upper(u.username),
    u.full_name,
    u.role,
    u.is_super_user,
    u.active,
    u.last_login_at
  from public.printmore_users u
  order by upper(u.username);
end;
$$;

create or replace function public.delete_printmore_user(
  p_admin_username text,
  p_admin_password text,
  p_username text
)
returns table(id uuid, username text, full_name text, role text, is_super_user boolean)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin public.printmore_users;
  v_user public.printmore_users;
  v_target_id uuid;
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

  select u.id
  into v_target_id
  from public.printmore_users u
  where lower(u.username) = lower(trim(p_username))
     or lower(u.full_name) = lower(trim(p_username))
  order by case when lower(u.username) = lower(trim(p_username)) then 0 else 1 end
  limit 1;

  if v_target_id is null then
    raise exception 'User id not found.';
  end if;

  if v_target_id = v_admin.id then
    raise exception 'You cannot delete your own user id.';
  end if;

  delete from public.printmore_users u
  where u.id = v_target_id
  returning * into v_user;

  return query
  select v_user.id, upper(v_user.username), v_user.full_name, v_user.role, v_user.is_super_user;
end;
$$;

grant execute on function public.authenticate_printmore_user(text, text) to anon;
grant execute on function public.create_printmore_user(text, text, text, text, text, text) to anon;
grant execute on function public.reset_printmore_user_password(text, text, text, text) to anon;
grant execute on function public.set_printmore_user_active(text, text, text, boolean) to anon;
grant execute on function public.update_printmore_user(text, text, text, text, boolean) to anon;
grant execute on function public.find_printmore_user(text) to anon;
grant execute on function public.list_printmore_users(text, text) to anon;
grant execute on function public.delete_printmore_user(text, text, text) to anon;
