-- ════════════════════════════════════════════════════════════════════════
-- Gestiq — Configuración del servidor de licencias (Supabase, plan gratis)
--
-- PASOS (una sola vez):
--  1. Crea cuenta y proyecto en https://supabase.com (gratis).
--  2. En el proyecto: SQL Editor → pega TODO este archivo → Run.
--  3. Authentication → Sign In / Up → desactiva "Allow new users to sign up"
--     (solo tú creas usuarios) y desactiva "Confirm email".
--  4. Settings → API → copia "Project URL" y "anon public key"
--     y pégalas en licencia.py (SUPABASE_URL y SUPABASE_ANON_KEY).
--
-- PARA CADA CLIENTE NUEVO:
--  a. Authentication → Users → Add user → email + contraseña.
--  b. Table Editor → licenses → Insert row:
--     user_id = (id del usuario recién creado), status = 'activa',
--     expires_at = fecha de fin de la suscripción, plan = 'mensual'.
--  Para suspender o renovar: edita esa fila (status / expires_at).
-- ════════════════════════════════════════════════════════════════════════

create table if not exists public.licenses (
  user_id    uuid primary key references auth.users(id) on delete cascade,
  status     text        not null default 'activa',      -- activa | suspendida
  plan       text        not null default 'mensual',     -- mensual | trimestral | anual
  expires_at timestamptz not null,
  machine_id text,                                       -- se fija solo en el primer uso
  updated_at timestamptz not null default now()
);

-- Nadie accede a la tabla directamente desde la app (solo vía la función)
alter table public.licenses enable row level security;

-- Equipos ilimitados para cuentas propietarias: si es true, check_license
-- NO ata ni compara machine_id (esa cuenta entra desde cualquier equipo).
alter table public.licenses
  add column if not exists unlimited_devices boolean not null default false;

-- Función que valida TODO en el servidor: licencia, estado, fecha (hora del
-- servidor) y equipo. La app solo recibe el veredicto.
create or replace function public.check_license(p_machine_id text)
returns json
language plpgsql
security definer
set search_path = public
as $$
declare
  lic licenses%rowtype;
begin
  select * into lic from licenses where user_id = auth.uid();

  if not found then
    return json_build_object('ok', false, 'reason', 'SIN_LICENCIA');
  end if;

  -- Equipos ilimitados (cuentas propietarias): no atar ni comparar machine_id
  if not coalesce(lic.unlimited_devices, false) then
    -- primer uso: ata la licencia a este equipo
    if lic.machine_id is null or lic.machine_id = '' then
      update licenses set machine_id = p_machine_id, updated_at = now()
        where user_id = auth.uid();
      lic.machine_id := p_machine_id;
    end if;

    if lic.machine_id <> p_machine_id then
      return json_build_object('ok', false, 'reason', 'OTRO_EQUIPO');
    end if;
  end if;

  if lic.status <> 'activa' then
    return json_build_object('ok', false, 'reason', 'SUSPENDIDA');
  end if;

  if lic.expires_at < now() then
    return json_build_object('ok', false, 'reason', 'VENCIDA');
  end if;

  return json_build_object(
    'ok', true,
    'plan', lic.plan,
    'expires_at', lic.expires_at,
    'server_time', now()
  );
end;
$$;

revoke execute on function public.check_license(text) from anon, public;
grant  execute on function public.check_license(text) to authenticated;

-- Para cambiar un cliente de equipo: vaciar su machine_id
--   update public.licenses set machine_id = null where user_id = '...';

-- ════════════════════════════════════════════════════════════════════════
-- Estado de licencia para la WEB (solo lectura, NO ata machine_id)
-- ════════════════════════════════════════════════════════════════════════
create or replace function public.license_status()
returns json
language plpgsql
security definer
set search_path = public
as $$
declare
  lic licenses%rowtype;
begin
  select * into lic from licenses where user_id = auth.uid();
  if not found then
    return json_build_object('ok', false, 'reason', 'SIN_LICENCIA');
  end if;
  return json_build_object(
    'ok', lic.status = 'activa' and lic.expires_at >= now(),
    'status', lic.status,
    'plan', lic.plan,
    'expires_at', lic.expires_at,
    'server_time', now()
  );
end;
$$;

revoke execute on function public.license_status() from anon, public;
grant  execute on function public.license_status() to authenticated;

-- ════════════════════════════════════════════════════════════════════════
-- Migración planes por módulo (jun-2026): imagine | guardian | completo
-- El plan indica QUÉ MÓDULOS puede usar el cliente; el periodo pagado
-- (1/6/12 meses) solo afecta expires_at.
-- ════════════════════════════════════════════════════════════════════════
alter table public.licenses alter column plan set default 'completo';
update public.licenses set plan = 'completo'
 where plan in ('mensual', 'trimestral', 'anual');

-- ════════════════════════════════════════════════════════════════════════
-- Cuenta propietaria con equipos ilimitados (jun-2026)
-- ════════════════════════════════════════════════════════════════════════
update public.licenses set unlimited_devices = true
 where user_id = (select id from auth.users
                   where email = 'camilovitola1205@gmail.com');

-- ════════════════════════════════════════════════════════════════════════
-- Perfil por cuenta (nombre, foto, tema, módulo) — guardado en el servidor
-- para que siga al usuario entre equipos. Cada quien solo ve/edita el suyo.
-- ════════════════════════════════════════════════════════════════════════
create table if not exists public.profiles (
  user_id    uuid primary key references auth.users(id) on delete cascade,
  nombre     text,
  foto       text,
  tema       text,
  modulo     text,
  updated_at timestamptz not null default now()
);
alter table public.profiles enable row level security;

drop policy if exists "perfil_select" on public.profiles;
create policy "perfil_select" on public.profiles
  for select using (auth.uid() = user_id);

drop policy if exists "perfil_insert" on public.profiles;
create policy "perfil_insert" on public.profiles
  for insert with check (auth.uid() = user_id);

drop policy if exists "perfil_update" on public.profiles;
create policy "perfil_update" on public.profiles
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);

-- Permiso de uso de la tabla para usuarios autenticados.
-- SIN esto, PostgREST devuelve 403 aunque las policies RLS existan.
grant select, insert, update on public.profiles to authenticated;
