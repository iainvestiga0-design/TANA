-- TANA: historial persistente por usuario
-- Ejecutar una sola vez en Supabase > SQL Editor.

create table if not exists public.tana_historial (
    id bigint generated always as identity primary key,
    user_id uuid not null references auth.users(id) on delete cascade,
    title text not null,
    created_at timestamptz not null default now()
);

create index if not exists tana_historial_user_created_idx
    on public.tana_historial (user_id, created_at desc);

alter table public.tana_historial enable row level security;

drop policy if exists "tana_historial_select_own" on public.tana_historial;
create policy "tana_historial_select_own"
on public.tana_historial
for select
to authenticated
using ((select auth.uid()) = user_id);

drop policy if exists "tana_historial_insert_own" on public.tana_historial;
create policy "tana_historial_insert_own"
on public.tana_historial
for insert
to authenticated
with check ((select auth.uid()) = user_id);

-- Permisos necesarios para la Data API.
grant select, insert on public.tana_historial to authenticated;
grant all on public.tana_historial to service_role;

grant usage, select on sequence public.tana_historial_id_seq to authenticated;
grant usage, select on sequence public.tana_historial_id_seq to service_role;
