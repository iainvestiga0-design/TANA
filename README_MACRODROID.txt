TANA — INTEGRACIÓN CON MACRODROID PARA CONFIRMAR YAPE
========================================================

Qué trae este ZIP
------------------
- app.py                              → tu app con UN panel nuevo, solo visible
                                         para el Creador: "🧪 Prueba de integración
                                         MacroDroid (Yape)".
- 02_macrodroid_confirmar_pago.sql    → función nueva en Supabase (aditiva).
- README_MACRODROID.txt               → este archivo.

No se tocó nada de lo que ya tenías: ni el motor contable, ni el flujo de
login, ni la tabla tana_access, ni las funciones anteriores. Solo se agregó
código nuevo.

Por qué se hizo así (y no "leyendo MacroDroid" directamente en Streamlit)
--------------------------------------------------------------------------
Streamlit no puede recibir peticiones HTTP de MacroDroid directamente (no es
un servidor de API). Lo que sí puede recibir peticiones HTTP es Supabase, que
ya usas. Por eso el celular con MacroDroid le habla directamente a Supabase
(no a tu app), y tu app —como ya hacía antes— sigue revisando la tabla
tana_payments para saber si el pago llegó. Es la misma arquitectura que ya
tenías, solo que ahora algo más además de ti puede escribir en esa tabla:
el celular, vía una función seguendia de Supabase.

Paso 1 — Aplicar el SQL
-------------------------
1. Supabase > SQL Editor > New query.
2. Pega TODO el contenido de 02_macrodroid_confirmar_pago.sql.
3. Run.
4. La consulta de verificación al final debe devolver una fila con
   "tana_confirm_payment_from_macrodroid".

Paso 2 — Subir el app.py nuevo
---------------------------------
1. En GitHub, reemplaza app.py por el de este ZIP.
2. Commit changes. Streamlit Cloud lo vuelve a desplegar solo.

Cómo cambió el flujo del estudiante
--------------------------------------
Ya NO se genera un código inventado por la app para que el estudiante lo
escriba en la nota de Yape. Ahora es al revés, usando el propio comprobante
de Yape:
  1. El estudiante yapea S/1.00 al número indicado.
  2. Yape le da un "Código de operación" (el número que aparece en su
     comprobante de pago).
  3. El estudiante pega ESE código en la app (pantalla del pase de 24h) y
     pulsa "Registrar código y verificar pago".
  4. MacroDroid, en el celular que RECIBE los pagos, detecta la notificación
     de Yape, extrae ese mismo código de operación y el monto, y llama a la
     función de Supabase para confirmarlo.
  5. Cuando el estudiante pulsa "Verificar pago", si ya está confirmado, se
     activa su pase de 24 horas.

Paso 3 — Probar SIN el celular todavía
-----------------------------------------
1. Entra a TANA con una cuenta de estudiante y en la pantalla del pase de
   24h escribe un código de operación de prueba (ej. "000482913") y pulsa
   "Registrar código y verificar pago".
2. Entra con tu cuenta de Creador y abre
   "🧪 Prueba de integración MacroDroid (Yape)".
3. Pega un texto de ejemplo con ESE MISMO código, por ejemplo:
     Yape: Te llegó un pago de Juan Pérez por S/ 1.00. Código de operación: 000482913
4. Pulsa "📲 Simular notificación de MacroDroid".
5. Si todo está bien, verás "✅ Pago confirmado para <correo>".
6. Vuelve a la cuenta de estudiante y pulsa "🔄 Verificar pago": debería
   activarse el pase de 24 horas normalmente, sin que tú toques nada más.

Este paso 3 ya te permite probar TODO el flujo de principio a fin, incluso
antes de configurar el celular real.

Paso 4 — Configurar el macro real en MacroDroid (cuando quieras probarlo
con el celular)
--------------------------------------------------------------------------
1. Disparador: "Notificación recibida" → selecciona la app de Yape, en el
   celular que RECIBE los pagos (no el del estudiante).
2. Acción "Analizar/Extraer texto" (o una variable con Regex) para separar
   del texto de la notificación:
     - el monto (ej. buscar patrón S/ seguido de números)
     - el código de operación (el número que Yape muestra como
       "Código de operación" o "N° de operación")
3. Acción "Solicitud HTTP (HTTP Request)":
     - Método: POST
     - URL: <TU_SUPABASE_URL>/rest/v1/rpc/tana_confirm_payment_from_macrodroid
     - Encabezados:
         apikey: <tu anon key de Supabase>
         Authorization: Bearer <tu anon key de Supabase>
         Content-Type: application/json
     - Cuerpo (JSON):
         {"p_code": "[código de operación extraído]", "p_amount": [variable monto], "p_raw_text": "[texto notificación]"}

Importante: usa la ANON key en el celular, no la service_role key. La
función ya quedó protegida (SECURITY DEFINER), así que la anon key alcanza
y es mucho más segura de tener guardada en un macro del teléfono.

Qué NO se cambió
------------------
- La lógica contable de TANA: intacta.
- El flujo de login y roles: intacto.
- Las funciones y tablas anteriores (tana_activate_24h, tana_access, etc.):
  intactas, no se borró ni reemplazó nada.
- El botón "Verificar pago" del estudiante: sigue funcionando exactamente
  igual que antes; ahora simplemente encuentra el pago más rápido porque
  algo más (tú simulando, o el celular real después) puede insertarlo.
