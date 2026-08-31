-- TANA — INTEGRACIÓN MACRODROID (YAPE) — v2
-- =========================================
-- Este script es 100% ADITIVO y SEGURO DE RE-EJECUTAR: usa CREATE OR REPLACE,
-- no borra tablas ni datos existentes. Reemplaza la función anterior
-- tana_confirm_payment_from_macrodroid por una versión que soluciona el
-- problema de "carrera contra el tiempo":
--
--   PROBLEMA ANTERIOR:
--   El estudiante solo puede ver su "código de seguridad" DESPUÉS de yapear.
--   Pero MacroDroid reacciona casi al instante (1-2 segundos) a la notificación
--   de Yape, así que casi siempre llegaba a Supabase ANTES de que el
--   estudiante terminara de escribir el código en la web. Como el código
--   todavía no estaba registrado, la función devolvía 'code_not_found' y
--   el pago se perdía para siempre (nadie reintentaba después).
--
--   SOLUCIÓN:
--   Si el código no está registrado todavía, la función ahora guarda el
--   pago como "huérfano" (email = NULL, status = 'confirmed') en vez de
--   descartarlo. Cuando el estudiante registra su código en la app (aunque
--   sea uno o dos minutos después), el código de la app (ver parche de
--   app.py más abajo) busca si ya existe un pago huérfano con ese mismo
--   código y, si lo encuentra, lo "reclama" asignándole su correo.
--
--   Devuelve una fila con: email, payment_code, amount, status
--   status puede ser:
--     'confirmed'          -> código ya estaba registrado, todo perfecto
--     'confirmed_orphan'   -> pago válido pero código aún no registrado;
--                             queda esperando a que el estudiante lo registre
--     'already_confirmed'  -> ese código ya se había confirmado antes (con dueño)
--     'already_orphan'     -> ese código ya está guardado como huérfano, no se duplica
--     'invalid_code'       -> el código venía vacío
--   (nunca lanza una excepción dura, para que la respuesta HTTP a MacroDroid
--   sea siempre 200 y el macro no falle).

-- Permite que el pago quede "sin dueño" temporalmente (email = NULL) mientras
-- se reclama. Si tu tabla tana_payments ya permitía NULL en email, esta línea
-- no hace nada (es segura de re-ejecutar).
alter table public.tana_payments alter column email drop not null;

create or replace function public.tana_confirm_payment_from_macrodroid(
    p_code text,
    p_amount numeric,
    p_raw_text text default null
)
returns table (
    email text,
    payment_code text,
    amount numeric,
    status text
)
language plpgsql
security definer
set search_path = public
as $$
declare
    v_email text;
    v_existing record;
    v_code text;
begin
    v_code := upper(trim(coalesce(p_code, '')));

    if v_code = '' then
        return query select null::text, v_code, p_amount, 'invalid_code'::text;
        return;
    end if;

    -- ¿Ya existe algún pago (confirmado o huérfano) registrado con este código?
    select tp.email, tp.payment_code, tp.amount, tp.status
    into v_existing
    from public.tana_payments tp
    where tp.payment_code = v_code
      and tp.status = 'confirmed'
    limit 1;

    if found then
        if v_existing.email is null then
            return query select null::text, v_existing.payment_code, v_existing.amount, 'already_orphan'::text;
        else
            return query select v_existing.email, v_existing.payment_code, v_existing.amount, 'already_confirmed'::text;
        end if;
        return;
    end if;

    -- ¿El estudiante ya registró este código en tana_payment_codes?
    select tpc.email into v_email
    from public.tana_payment_codes tpc
    where upper(tpc.code) = v_code
    order by tpc.created_at desc
    limit 1;

    if v_email is null then
        -- Todavía no lo registró: guardamos el pago como "huérfano" para
        -- que se reclame en cuanto el estudiante ingrese el código en la app.
        insert into public.tana_payments (email, payment_code, amount, status, paid_at)
        values (null, v_code, p_amount, 'confirmed', now());

        return query select null::text, v_code, p_amount, 'confirmed_orphan'::text;
        return;
    end if;

    insert into public.tana_payments (email, payment_code, amount, status, paid_at)
    values (v_email, v_code, p_amount, 'confirmed', now());

    return query select v_email, v_code, p_amount, 'confirmed'::text;
end;
$$;

grant execute on function public.tana_confirm_payment_from_macrodroid(text, numeric, text) to anon, authenticated;

-- Verificación (solo lectura, no cambia nada):
select routine_name
from information_schema.routines
where routine_schema = 'public'
  and routine_name = 'tana_confirm_payment_from_macrodroid';
