-- TANA — INTEGRACIÓN MACRODROID (YAPE)
-- =========================================
-- Este script es 100% ADITIVO: no borra ni reemplaza ninguna tabla,
-- función, constraint ni dato existente. Solo crea una función nueva.
--
-- Qué hace:
--   MacroDroid detecta la notificación de Yape en el celular, extrae el
--   código TANA-XXXXXX (que el estudiante escribe en el mensaje/nota del
--   Yape) y el monto, y llama a esta función vía la API REST de Supabase
--   (POST a /rest/v1/rpc/tana_confirm_payment_from_macrodroid).
--
--   La función:
--     1. Busca en tana_payment_codes a qué correo pertenece ese código.
--     2. Si ya existe un pago "confirmed" para ese código, no lo duplica.
--     3. Si no existe, inserta la fila en tana_payments con status='confirmed'.
--        Esa tabla ya la lee tu app (_payment_confirmed_for_code), así que
--        el botón "Verificar pago" del estudiante empieza a funcionar solo,
--        sin más cambios en el código de la app.
--
--   Devuelve una fila con: email, payment_code, amount, status
--   status puede ser: 'confirmed' | 'already_confirmed' | 'code_not_found' | 'invalid_code'
--   (nunca lanza una excepción dura, para que la respuesta HTTP a MacroDroid
--   sea siempre 200 y el macro no falle).

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

    select tpc.email into v_email
    from public.tana_payment_codes tpc
    where upper(tpc.code) = v_code
    order by tpc.created_at desc
    limit 1;

    if v_email is null then
        return query select null::text, v_code, p_amount, 'code_not_found'::text;
        return;
    end if;

    select tp.email, tp.payment_code, tp.amount, tp.status
    into v_existing
    from public.tana_payments tp
    where tp.payment_code = v_code
      and tp.status = 'confirmed'
    limit 1;

    if found then
        return query select v_existing.email, v_existing.payment_code, v_existing.amount, 'already_confirmed'::text;
        return;
    end if;

    insert into public.tana_payments (email, payment_code, amount, status, paid_at)
    values (v_email, v_code, p_amount, 'confirmed', now());

    return query select v_email, v_code, p_amount, 'confirmed'::text;
end;
$$;

-- Permite llamar a esta función con la anon key (recomendado para MacroDroid:
-- así el celular nunca necesita cargar la service_role key, que es mucho
-- más peligrosa si el teléfono se pierde o alguien clona el macro).
-- La función es SECURITY DEFINER, así que igual puede leer/escribir aunque
-- las tablas tengan RLS activado y bloqueen el acceso directo por anon key.
grant execute on function public.tana_confirm_payment_from_macrodroid(text, numeric, text) to anon, authenticated;

-- Verificación (solo lectura, no cambia nada):
select routine_name
from information_schema.routines
where routine_schema = 'public'
  and routine_name = 'tana_confirm_payment_from_macrodroid';
