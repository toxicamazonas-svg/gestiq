-- ════════════════════════════════════════════════════════════════════════
-- IPRECON · Cuentas con nombre de usuario
-- Aplicar en Supabase → SQL Editor (proyecto xdydorreyvkenbifefus).
-- Idempotente: se puede correr varias veces sin romper nada.
-- ════════════════════════════════════════════════════════════════════════

-- 1) Campos nuevos en profiles
alter table public.profiles add column if not exists apellido  text;
alter table public.profiles add column if not exists username  text;
alter table public.profiles add column if not exists celular   text;
alter table public.profiles add column if not exists tipo_doc  text;
alter table public.profiles add column if not exists documento text;

-- 2) Username único, insensible a mayúsculas (ignora vacíos/nulos)
create unique index if not exists profiles_username_uniq
  on public.profiles (lower(username))
  where username is not null and username <> '';

-- 3) ¿Username disponible?  (lo llama el registro, aún sin sesión)
create or replace function public.username_disponible(p_username text)
returns boolean
language sql
security definer
set search_path = public
as $$
  select not exists (
    select 1 from public.profiles
    where lower(username) = lower(trim(p_username))
  );
$$;
grant execute on function public.username_disponible(text) to anon, authenticated;

-- 4) Resolver username → correo  (lo usará el login por usuario, Fase 2)
--    NOTA: accesible sin sesión; permite a un atacante mapear usuario→correo
--    por sondeo. Aceptable para empezar. Si se quiere blindar, mover a Edge Function.
create or replace function public.email_de_usuario(p_username text)
returns text
language sql
security definer
set search_path = public, auth
as $$
  select u.email
  from public.profiles p
  join auth.users u on u.id = p.user_id
  where lower(p.username) = lower(trim(p_username))
  limit 1;
$$;
grant execute on function public.email_de_usuario(text) to anon, authenticated;
