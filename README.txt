TANA V10 CORREGIDA — LECTURA DE PRÁCTICAS + VOZ

Esta versión parte de la V10 de actividad persistente en Supabase y corrige dos problemas observados en la prueba:

1. PRÁCTICAS WORD / DOC
- Mantiene PDF, DOC, DOCX, XLS, XLSX, JPG, JPEG y PNG.
- Para Word antiguo (.doc) se envía el MIME application/msword cuando el SDK lo permite.
- Si la extracción estructurada devuelve 0 operaciones o JSON vacío, TANA hace una segunda lectura: primero recupera el texto completo del documento y después vuelve a estructurarlo.
- TANA ya no continúa silenciosamente con una práctica vacía.
- Si Gemini reconoce la práctica pero devuelve 0 asientos, TANA no genera ni ofrece un Excel vacío.

2. VOZ
- El procesamiento por voz ya no depende de que existan asientos previamente generados.
- El audio se sube con MIME audio/wav cuando el SDK lo permite.
- El audio solo se marca como procesado después de obtener una respuesta.
- Si Gemini falla, el error queda visible en el historial y el mismo audio no se marca como procesado.
- TANA puede responder por voz incluso si todavía no hay una monografía cargada.

SUPABASE
No se modifica el esquema de Supabase. Se conserva:
- tana_users
- tana_activity

STREAMLIT SECRETS
Mantener los Secrets actuales. No incluir claves secretas en este ZIP ni en GitHub.

IMPORTANTE
Esta versión no cambia la lógica contable de PCGE, HT, ERN, ERF, ESF ni el sistema de actividad persistente que ya quedó funcionando.
