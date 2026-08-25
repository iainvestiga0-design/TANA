-- TANA V8: almacenamiento persistente de usuarios y actividad en Supabase
create table if not exists public.tana_users (
  email text primary key,
  role text not null default 'Estudiante',
  first_seen timestamptz not null,
  last_seen timestamptz not null,
  visits integer not null default 0,
  activity jsonb not null default '[]'::jsonb
);

-- TANA usa SUPABASE_SERVICE_ROLE_KEY en Streamlit, por lo que las políticas
-- RLS no son necesarias para las llamadas del servidor. Mantén esta clave
-- únicamente en Streamlit Secrets y nunca la publiques en GitHub.
