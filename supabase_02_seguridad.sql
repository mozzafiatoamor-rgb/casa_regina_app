-- =====================================================================
-- CASA REGINA — Fase 3: SEGURIDAD (login de personal + reglas RLS)
-- =====================================================================
-- Qué hace:
--   1. Crea una tabla casa_regina.staff que liga cada cuenta de login (Supabase Auth)
--      con su rol en el hotel.
--   2. Crea la función is_staff() que responde "¿este usuario es personal del hotel?".
--   3. Da permisos al rol 'authenticated' (el de los usuarios con sesión) PERO con RLS:
--      solo el personal del hotel (is_staff) puede leer/escribir. Nadie más (ni anónimos,
--      ni usuarios de tus otras 2 apps) puede tocar estos datos.
--
-- Es SEGURO correrlo ahora: NO afecta tu app en vivo (hoy la app no lee de Supabase;
-- el backend usa la llave de servicio, que ignora RLS). Solo prepara la seguridad.
--
-- Cómo correrlo: SQL Editor → New query → pega TODO → Run. "Success. No rows returned".
-- =====================================================================

-- 1) Tabla de personal del hotel (una fila por empleado con cuenta de login)
create table if not exists casa_regina.staff (
  user_id uuid primary key references auth.users(id) on delete cascade,
  usuario text,
  nombre  text,
  rol     text,
  creado  timestamptz default now()
);
alter table casa_regina.staff enable row level security;

-- 2) ¿El usuario actual es personal del hotel?
create or replace function casa_regina.is_staff()
returns boolean
language sql
stable
security definer
set search_path = casa_regina, public
as $$
  select exists (select 1 from casa_regina.staff s where s.user_id = auth.uid());
$$;

-- 3) Permisos para usuarios CON sesión (authenticated). RLS filtra abajo por is_staff().
grant usage on schema casa_regina to authenticated;
grant select, insert, update, delete on all tables in schema casa_regina to authenticated;
grant usage, select on all sequences in schema casa_regina to authenticated;
alter default privileges in schema casa_regina grant select, insert, update, delete on tables to authenticated;
alter default privileges in schema casa_regina grant usage, select on sequences to authenticated;

-- 4) El personal puede leer SU propia fila de staff (necesario para el login / is_staff)
drop policy if exists staff_self on casa_regina.staff;
create policy staff_self on casa_regina.staff
  for select to authenticated using (user_id = auth.uid());

-- 5) Política "acceso total del personal" en cada tabla de datos del hotel.
--    for all = leer, insertar, actualizar, borrar — pero SOLO si is_staff() es verdadero.
do $$
declare t text;
begin
  foreach t in array array[
    'habitaciones','reservas','registro','accesos','housekeeping','bitacora',
    'usuarios','cobros','clientes','inventario','desayunos','cotizaciones','gastos','config'
  ] loop
    execute format('drop policy if exists staff_all on casa_regina.%I', t);
    execute format(
      'create policy staff_all on casa_regina.%I for all to authenticated using (casa_regina.is_staff()) with check (casa_regina.is_staff())',
      t);
  end loop;
end $$;

-- =====================================================================
-- Resultado: con esto, un usuario con sesión que esté en casa_regina.staff
-- tiene acceso completo; cualquier otro (anónimo o de otra app) NO ve nada.
-- El backend (service_role) sigue funcionando igual, sin cambios.
-- =====================================================================
