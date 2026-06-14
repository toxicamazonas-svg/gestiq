-- ════════════════════════════════════════════════════════════════════════
-- IPRECON · Exigir 2FA en el servidor (anti "API directa")
-- Si la cuenta tiene un factor TOTP verificado, una sesión sin segundo factor
-- (aal1) NO puede validar licencia ni leer/escribir su perfil.
-- Aplicar en Supabase → SQL Editor. Requiere haber aplicado antes db_usuarios.sql.
-- ════════════════════════════════════════════════════════════════════════

-- 1) check_license (app): rechaza aal1 cuando hay 2FA activo
create or replace function public.check_license(p_machine_id text)
returns json
language plpgsql
security definer
set search_path = public
as $$
declare
  lic licenses%rowtype;
begin
  -- 2FA: si la cuenta tiene un segundo factor verificado, exige nivel aal2
  if exists (select 1 from auth.mfa_factors
             where user_id = auth.uid() and status = 'verified')
     and coalesce(auth.jwt() ->> 'aal', 'aal1') <> 'aal2' then
    return json_build_object('ok', false, 'reason', 'REQUIERE_2FA');
  end if;

  select * into lic from licenses where user_id = auth.uid();

  if not found then
    return json_build_object('ok', false, 'reason', 'SIN_LICENCIA');
  end if;

  if not coalesce(lic.unlimited_devices, false) then
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

-- 2) license_status (web): mismo chequeo de 2FA
create or replace function public.license_status()
returns json
language plpgsql
security definer
set search_path = public
as $$
declare
  lic licenses%rowtype;
begin
  if exists (select 1 from auth.mfa_factors
             where user_id = auth.uid() and status = 'verified')
     and coalesce(auth.jwt() ->> 'aal', 'aal1') <> 'aal2' then
    return json_build_object('ok', false, 'reason', 'REQUIERE_2FA');
  end if;

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

-- 3) profiles: protege los datos personales con aal2 cuando hay 2FA activo.
--    Política RESTRICTIVA (se suma con AND a las permisivas existentes).
--    Si la cuenta NO tiene factores verificados, acepta aal1 o aal2 (no estorba
--    el registro ni a quien no usa 2FA).
drop policy if exists "perfil_mfa" on public.profiles;
create policy "perfil_mfa" on public.profiles
  as restrictive to authenticated
  using (
    (array[coalesce(auth.jwt() ->> 'aal', 'aal1')]) <@ (
      select case when count(id) > 0 then array['aal2']
                  else array['aal1','aal2'] end
      from auth.mfa_factors
      where user_id = auth.uid() and status = 'verified'
    )
  )
  with check (
    (array[coalesce(auth.jwt() ->> 'aal', 'aal1')]) <@ (
      select case when count(id) > 0 then array['aal2']
                  else array['aal1','aal2'] end
      from auth.mfa_factors
      where user_id = auth.uid() and status = 'verified'
    )
  );
