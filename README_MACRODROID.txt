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

Paso 3 — Probar SIN el celular todavía
-----------------------------------------
1. Entra a TANA con tu cuenta de Creador.
2. Genera (como si fueras estudiante, en otra pestaña/cuenta) un código de
   pago TANA-XXXXXX.
3. En tu cuenta de Creador, abre "🧪 Prueba de integración MacroDroid (Yape)".
4. Pega un texto de ejemplo como:
     Te llegó un Yape de Juan Pérez por S/ 1.00. Mensaje: TANA-482913
   (usa el código real que generaste).
5. Pulsa "📲 Simular notificación de MacroDroid".
6. Si todo está bien, verás "✅ Pago confirmado para <correo>".
7. Vuelve a la cuenta de estudiante y pulsa "🔄 Verificar pago": debería
   activarse el pase de 24 horas normalmente, sin que tú toques nada más.

Este paso 3 ya te permite probar TODO el flujo de principio a fin, incluso
antes de configurar el celular real.

Paso 4 — Configurar el macro real en MacroDroid (cuando quieras probarlo
con el celular)
--------------------------------------------------------------------------
1. Disparador: "Notificación recibida" → selecciona la app de Yape.
2. Acción "Analizar/Extraer texto" (o una variable con Regex) para separar
   del texto de la notificación:
     - el monto (ej. buscar patrón S/ seguido de números)
     - el código (ej. buscar patrón TANA- seguido de números)
3. Acción "Solicitud HTTP (HTTP Request)":
     - Método: POST
     - URL: <TU_SUPABASE_URL>/rest/v1/rpc/tana_confirm_payment_from_macrodroid
     - Encabezados:
         apikey: <tu anon key de Supabase>
         Authorization: Bearer <tu anon key de Supabase>
         Content-Type: application/json
     - Cuerpo (JSON):
         {"p_code": "[variable código]", "p_amount": [variable monto], "p_raw_text": "[texto notificación]"}

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
