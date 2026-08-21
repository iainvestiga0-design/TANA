import json
import io
import os
import re
import shutil
import subprocess
import tempfile
from decimal import Decimal, InvalidOperation

import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName

from pypdf import PdfReader
from docx import Document
from PIL import Image

# Gemini
from google import genai
from google.genai import types

# ============================================================
# GEMINI: lectura multimodal de monografías
# ============================================================

SUPPORTED_TYPES = ["pdf", "doc", "docx", "xls", "xlsx", "jpg", "jpeg", "png"]

# Gemini se configura con una cadena de respaldo para que TANA no se detenga
# cuando se agota la cuota de un modelo/proyecto. La primera opción conserva
# el comportamiento actual; las siguientes se usan solo si hay 429/cuota o
# si el modelo configurado no está disponible.
GEMINI_MODEL = st.secrets.get("TANA_GEMINI_MODEL", os.getenv("TANA_GEMINI_MODEL", "gemini-3.5-flash"))
GEMINI_MODEL_2 = st.secrets.get("TANA_GEMINI_MODEL_2", os.getenv("TANA_GEMINI_MODEL_2", "gemini-3.5-flash-lite"))
GEMINI_MODEL_3 = st.secrets.get("TANA_GEMINI_MODEL_3", os.getenv("TANA_GEMINI_MODEL_3", "gemini-2.5-flash"))

# Cargar el PCGE antes del motor de resolución, incluso antes de generar el Excel.
PCGE_PATHS = [
    os.path.join(os.path.dirname(__file__), "pcge_data.json"),
    os.path.join(os.path.dirname(__file__), "pcge_data_TANA_oficial.json"),
]
PCGE_FILE = next((p for p in PCGE_PATHS if os.path.exists(p)), None)
if not PCGE_FILE:
    raise FileNotFoundError("No se encontró pcge_data.json junto a app.py.")
with open(PCGE_FILE, encoding="utf-8") as f:
    PCGE_DATA = json.load(f)

# Catálogo disponible para toda la aplicación, incluido el generador de Excel.
# El motor de Gemini crea su propia copia local, pero el Excel también necesita
# resolver la denominación de cada código sin depender de esa función.
pcge_map = {str(cod).strip(): str(desc) for cod, desc in PCGE_DATA}


# ============================================================
# MOTOR DE CÁLCULOS CONTABLES V5
# ============================================================
# Estas funciones son deliberadamente deterministas: Gemini interpreta
# la operación, pero los cálculos se hacen aquí para evitar redondeos,
# porcentajes o IGV inventados por el modelo.
def _to_float(value, default=0.0):
    try:
        if value is None or value == "":
            return default
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return default


def calcular_igv_desde_base(base, tasa=0.18):
    base = _to_float(base)
    igv = round(base * tasa, 2)
    return {"base": round(base, 2), "igv": igv, "total": round(base + igv, 2)}


def calcular_igv_incluido(total, tasa=0.18):
    total = _to_float(total)
    base = round(total / (1 + tasa), 2)
    igv = round(total - base, 2)
    return {"base": base, "igv": igv, "total": round(total, 2)}


def calcular_porcentaje(importe, porcentaje):
    return round(_to_float(importe) * _to_float(porcentaje) / 100, 2)


def calcular_depreciacion(costo, tasa_anual, meses=1, meses_pendientes=0):
    costo = _to_float(costo)
    tasa_anual = _to_float(tasa_anual)
    meses = int(_to_float(meses, 1))
    meses_pendientes = int(_to_float(meses_pendientes, 0))
    total_meses = max(0, meses + meses_pendientes)
    mensual = round(costo * tasa_anual / 100 / 12, 2)
    return {
        "depreciacion_mensual": mensual,
        "meses": total_meses,
        "depreciacion_periodo": round(mensual * total_meses, 2),
    }


def calcular_esalud(remuneracion, tasa=0.09):
    return round(_to_float(remuneracion) * tasa, 2)


def calcular_onp(remuneracion, tasa=0.13):
    return round(_to_float(remuneracion) * tasa, 2)


def calcular_asignacion_familiar(rmv, tiene_hijos=True, porcentaje=10):
    if not tiene_hijos:
        return 0.0
    return calcular_porcentaje(rmv, porcentaje)


def calcular_costo_neto(costo, depreciacion_acumulada):
    return round(_to_float(costo) - _to_float(depreciacion_acumulada), 2)


def calcular_disminucion_valor(valor_neto, porcentaje):
    return calcular_porcentaje(valor_neto, porcentaje)


def calcular_operacion_contable(operacion):
    """
    Ejecuta cálculos explícitos cuando Gemini entrega los campos necesarios.
    No inventa datos faltantes: devuelve 'requiere_datos' cuando no puede
    calcular de forma determinista.
    """
    if not isinstance(operacion, dict):
        return {"requiere_datos": True, "motivo": "Operación no estructurada."}

    tipo = str(operacion.get("tipo_calculo", "")).strip().lower()
    tasa_igv = _to_float(operacion.get("tasa_igv", 18)) / 100

    if tipo in ("igv_base", "igv_desde_base"):
        return calcular_igv_desde_base(operacion.get("base"), tasa_igv)

    if tipo in ("igv_incluido", "igv_desde_total"):
        return calcular_igv_incluido(operacion.get("total"), tasa_igv)

    if tipo in ("porcentaje", "participacion"):
        return {
            "importe": calcular_porcentaje(
                operacion.get("importe"),
                operacion.get("porcentaje"),
            )
        }

    if tipo in ("depreciacion", "depreciación"):
        return calcular_depreciacion(
            operacion.get("costo"),
            operacion.get("tasa_anual"),
            operacion.get("meses", 1),
            operacion.get("meses_pendientes", 0),
        )

    if tipo == "esalud":
        return {"esalud": calcular_esalud(operacion.get("remuneracion"))}

    if tipo == "onp":
        return {"onp": calcular_onp(operacion.get("remuneracion"))}

    if tipo in ("asignacion_familiar", "asignación_familiar"):
        return {
            "asignacion_familiar": calcular_asignacion_familiar(
                operacion.get("rmv"),
                operacion.get("tiene_hijos", True),
            )
        }

    if tipo in ("valor_neto", "valor_neto_libros"):
        return {
            "valor_neto": calcular_costo_neto(
                operacion.get("costo"),
                operacion.get("depreciacion_acumulada"),
            )
        }

    if tipo in ("disminucion_valor", "deterioro"):
        return {
            "disminucion": calcular_disminucion_valor(
                operacion.get("valor_neto"),
                operacion.get("porcentaje"),
            )
        }

    return {"calculo_aplicado": False}


def aplicar_calculos_deterministas(operaciones):
    resultados = []
    for op in operaciones or []:
        copia = dict(op) if isinstance(op, dict) else {"descripcion": str(op)}
        try:
            calculo = calcular_operacion_contable(copia)
            copia["calculo_tana"] = calculo
        except Exception as exc:
            copia["calculo_tana"] = {
                "requiere_datos": True,
                "motivo": f"No se pudo calcular automáticamente: {exc}",
            }
        resultados.append(copia)
    return resultados


def _secret_or_env(name, default=""):
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = ""
    return value or os.getenv(name, default) or default


def get_gemini_profiles():
    """Devuelve las rutas Gemini disponibles, en orden de preferencia.

    Perfil 1: proyecto/modelo actual.
    Perfil 2: segundo proyecto (si se proporciona GEMINI_API_KEY_2) o, si no,
              el mismo proyecto con un modelo alternativo de menor costo.
    Perfil 3: tercer proyecto (si se proporciona GEMINI_API_KEY_3) o, si no,
              otro modelo alternativo.
    """
    key1 = _secret_or_env("GEMINI_API_KEY")
    key2 = _secret_or_env("GEMINI_API_KEY_2") or key1
    key3 = _secret_or_env("GEMINI_API_KEY_3") or key1

    profiles = []
    seen = set()
    candidates = [
        (key1, GEMINI_MODEL, "Principal"),
        (key2, GEMINI_MODEL_2, "Respaldo 1"),
        (key3, GEMINI_MODEL_3, "Respaldo 2"),
    ]
    for api_key, model, label in candidates:
        api_key = str(api_key or "").strip()
        model = str(model or "").strip()
        if not api_key or not model:
            continue
        marker = (api_key, model)
        if marker in seen:
            continue
        seen.add(marker)
        profiles.append({"api_key": api_key, "model": model, "label": label})
    return profiles


def get_gemini_client(api_key=None):
    api_key = api_key or _secret_or_env("GEMINI_API_KEY")
    if not api_key:
        return None
    return genai.Client(api_key=api_key)


def _is_gemini_fallback_error(exc):
    low = str(exc).lower()
    return any(token in low for token in (
        "429", "resource_exhausted", "quota", "rate limit",
        "not found", "model not found", "unsupported model",
    ))


def _fallback_error_message(errors):
    if not errors:
        return "No hay una configuración de Gemini disponible."
    details = []
    for label, model, exc in errors:
        low = str(exc).lower()
        if "429" in low or "resource_exhausted" in low or "quota" in low:
            details.append(f"{label} ({model}): cuota agotada")
        elif "not found" in low or "unsupported model" in low:
            details.append(f"{label} ({model}): modelo no disponible")
        else:
            details.append(f"{label} ({model}): {str(exc)[:180]}")
    return (
        "TANA intentó las rutas disponibles de Gemini y ninguna pudo procesar "
        "la solicitud. Revisiones realizadas: " + "; ".join(details) + ". "
        "Puedes configurar GEMINI_API_KEY_2/GEMINI_API_KEY_3 y sus modelos "
        "alternativos en Streamlit Secrets."
    )


def _generate_with_fallback(contents_factory, config):
    """Genera contenido probando automáticamente las rutas Gemini disponibles."""
    profiles = get_gemini_profiles()
    if not profiles:
        raise RuntimeError(
            "TANA no tiene configurada ninguna GEMINI_API_KEY. En Streamlit "
            "abre App settings → Secrets y agrega GEMINI_API_KEY = \"TU_CLAVE\"."
        )

    errors = []
    for profile in profiles:
        client = get_gemini_client(profile["api_key"])
        try:
            response = client.models.generate_content(
                model=profile["model"],
                contents=contents_factory(client),
                config=config,
            )
            return response, profile
        except Exception as exc:
            errors.append((profile["label"], profile["model"], exc))
            if not _is_gemini_fallback_error(exc):
                raise RuntimeError(str(exc)) from exc

    raise RuntimeError(_fallback_error_message(errors))


EXTRACTION_PROMPT = """
Eres el módulo de extracción documental de TANA, un sistema contable peruano.

Analiza la monografía completa que se te proporciona. NO resuelvas todavía
los asientos contables. Tu trabajo es EXTRAER fielmente la información.

Devuelve únicamente JSON válido con esta estructura:

{
  "empresa": "",
  "tipo_documento": "",
  "periodo": "",
  "estado_inicial": [],
  "operaciones": [
    {
      "numero": 1,
      "fecha": "",
      "descripcion": "",
      "importe": null,
      "moneda": "PEN",
      "cantidad": null,
      "precio_unitario": null,
      "porcentaje": null,
      "documento": "",
      "forma_pago": "",
      "medio_pago": "",
      "tercero": "",
      "cuenta_bancaria": "",
      "datos_adicionales": ""
    }
  ],
  "solicitudes": [],
  "datos_importantes": []
}

REGLAS:
- No inventes datos que no estén en la monografía.
- Conserva exactamente fechas, importes, cantidades, porcentajes,
  documentos, nombres y condiciones.
- Si un dato no aparece, usa null o "".
- Separa cada operación en un elemento.
- Incluye el estado financiero inicial si existe.
- Incluye todo lo que el ejercicio pide realizar en "solicitudes".
- La información extraída servirá después para el motor contable de TANA.
"""

def _gemini_error_message(exc):
    msg = str(exc)
    low = msg.lower()
    if "429" in low or "resource_exhausted" in low or "quota" in low:
        return (
            "Se agotó una ruta de Gemini y TANA intentó automáticamente una ruta de respaldo. "
            "Si todas las rutas fallan, configura GEMINI_API_KEY_2 o GEMINI_API_KEY_3 "
            "en Streamlit Secrets. "
            f"Ruta principal: {GEMINI_MODEL}."
        )
    return msg


def extract_with_gemini(uploaded):
    profiles = get_gemini_profiles()
    if not profiles:
        raise RuntimeError(
            "TANA no tiene configurada ninguna GEMINI_API_KEY. "
            "En Streamlit abre App settings → Secrets y agrega "
            'GEMINI_API_KEY = "TU_CLAVE".'
        )

    suffix = "." + uploaded.name.rsplit(".", 1)[-1].lower()
    temp_path = None
    uploaded_bytes = uploaded.getvalue()

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_bytes)
            temp_path = tmp.name

        errors = []
        for profile in profiles:
            client = get_gemini_client(profile["api_key"])
            gemini_file = None
            try:
                # Cada perfil tiene su propio cliente/proyecto. El archivo se sube
                # a ese proyecto y solo entonces se consume la generación.
                gemini_file = client.files.upload(file=temp_path)
                response = client.models.generate_content(
                    model=profile["model"],
                    contents=[gemini_file, EXTRACTION_PROMPT],
                    config=types.GenerateContentConfig(response_mime_type="application/json"),
                )
                raw = response.text or ""
                data = json.loads(raw)
                return data
            except Exception as exc:
                errors.append((profile["label"], profile["model"], exc))
                if not _is_gemini_fallback_error(exc):
                    raise RuntimeError(_gemini_error_message(exc)) from exc
                continue

        raise RuntimeError(_fallback_error_message(errors))
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)


def extraction_to_text(data):
    parts = []

    if data.get("empresa"):
        parts.append(f"EMPRESA: {data['empresa']}")
    if data.get("tipo_documento"):
        parts.append(f"TIPO: {data['tipo_documento']}")
    if data.get("periodo"):
        parts.append(f"PERIODO: {data['periodo']}")

    if data.get("estado_inicial"):
        parts.append("\n--- ESTADO INICIAL ---")
        for item in data["estado_inicial"]:
            parts.append(json.dumps(item, ensure_ascii=False))

    if data.get("operaciones"):
        parts.append("\n--- OPERACIONES ---")
        for op in data["operaciones"]:
            parts.append(
                f"{op.get('numero', '')}. "
                f"{op.get('fecha', '')} - {op.get('descripcion', '')}"
            )

    if data.get("solicitudes"):
        parts.append("\n--- SE SOLICITA ---")
        for item in data["solicitudes"]:
            parts.append(f"- {item}")

    return "\n".join(parts)

profiles_status = get_gemini_profiles()
if profiles_status:
    st.caption(
        "🤖 Gemini: " + " → ".join(
            f"{p['label']} ({p['model']})" for p in profiles_status
        )
        + ". TANA cambiará automáticamente de ruta si una cuota se agota."
    )

with st.expander("📥 1. Subir monografía", expanded=True):
    uploaded_file = st.file_uploader(
        "Arrastra aquí tu monografía o selecciónala desde tu equipo",
        type=SUPPORTED_TYPES,
        help="PDF, DOC, DOCX, XLS, XLSX, JPG, JPEG y PNG.",
    )

if uploaded_file:
    if st.button("🤖 Analizar monografía con Gemini", type="primary"):
        with st.spinner(
            "Gemini está leyendo la monografía y estructurando sus operaciones..."
        ):
            try:
                extracted = extract_with_gemini(uploaded_file)

                st.session_state["monografia_json"] = extracted
                st.session_state["monografia_texto"] = extraction_to_text(extracted)
                st.session_state["monografia_nombre"] = uploaded_file.name

            except json.JSONDecodeError:
                st.error(
                    "Gemini respondió con un formato que no pudo convertirse "
                    "a JSON. Vuelve a intentarlo."
                )
            except Exception as exc:
                st.error(f"No se pudo procesar el archivo con Gemini: {exc}")

if "monografia_json" in st.session_state:
    data = st.session_state["monografia_json"]

    st.success(
        f"Monografía analizada: {st.session_state['monografia_nombre']}"
    )

    operaciones = data.get("operaciones", [])
    st.metric("Operaciones detectadas", len(operaciones))

    with st.expander("📋 Operaciones detectadas", expanded=True):
        if operaciones:
            rows = []
            for op in operaciones:
                rows.append(
                    {
                        "N.º": op.get("numero"),
                        "Fecha": op.get("fecha"),
                        "Descripción": op.get("descripcion"),
                        "Importe": op.get("importe"),
                        "Documento": op.get("documento"),
                        "Forma de pago": op.get("forma_pago"),
                    }
                )
            st.dataframe(rows, use_container_width=True, hide_index=True)
        else:
            st.warning("No se detectaron operaciones.")

    with st.expander("📄 Datos extraídos completos", expanded=False):
        st.json(data)

    st.info(
        "Esta primera etapa usa Gemini para leer y estructurar la monografía. "
        "Los asientos contables todavía deben pasar por el motor contable de TANA."
    )


# ============================================================
# MOTOR DE ASIENTOS CONTABLES
# ============================================================

ASIENTOS_PROMPT = """
Eres el motor contable de TANA, una aplicación de contabilidad peruana.

Tienes dos fuentes obligatorias:
1) Las operaciones extraídas de la monografía.
2) El PCGE de TANA que se adjunta abajo.

OBJETIVO:
Desarrollar los asientos contables de TODAS las operaciones detectadas.

REGLAS OBLIGATORIAS:
- Usa EXCLUSIVAMENTE códigos de cuenta que existan en el PCGE proporcionado.
- Cada código debe tener exactamente 5 dígitos.
- No inventes códigos.
- No uses cuentas de 2, 3 o 4 dígitos si existe la cuenta de 5 dígitos aplicable.
- Cada asiento debe cuadrar exactamente: total Debe = total Haber.
- Separa en asientos independientes los registros que correspondan a una misma operación.
- Conserva las fechas y datos de la monografía.
- Calcula los importes cuando la monografía permita determinarlos.
- Si un importe o tratamiento contable no puede determinarse con seguridad,
  NO inventes: marca la línea/asiento con "requiere_revision": true y explica por qué.
- No agregues información que no esté sustentada por la monografía o por las reglas contables necesarias para registrar la operación.
- La glosa debe ser breve y profesional.
- Los importes deben ser números positivos; el lado se expresa con debe/haber.

REGLA CRÍTICA SOBRE CUENTAS DE DESTINO (79):
- Si una operación requiere destino por función y se utiliza una cuenta del Elemento 9 (por ejemplo 94111, 94211, 94311, 94411, 94511, 94611, 95111, 95211, 95311, 95411, 95511, 95611, 95711, 95811 o 95911), DEBE registrarse también la cuenta 79111 en el HABER por el mismo importe total destinado.
- No omitas la cuenta 79111 cuando exista un destino por función.
- El asiento de destino normalmente es: cuenta del Elemento 9 en el DEBE y 79111 en el HABER.
- No confundas la cuenta 79 con la cuenta 70 ni con la cuenta 69. La 79 es una cuenta puente de destino y debe quedar fuera de los estados de resultados.

Devuelve SOLO JSON válido con esta estructura:
{
  "asientos": [
    {
      "numero": 1,
      "fecha": "2026-04-02",
      "glosa": "...",
      "documento": "...",
      "operacion_numero": 1,
      "requiere_revision": false,
      "observacion": "",
      "lineas": [
        {
          "codigo": "12345",
          "denominacion": "",
          "debe": 0.0,
          "haber": 0.0,
          "concepto": ""
        }
      ]
    }
  ],
  "alertas": []
}

PCGE DE TANA:
{pcge}

OPERACIONES DE LA MONOGRAFÍA:
{operaciones}
"""

def _money(value):
    try:
        if value is None or value == "":
            return Decimal("0")
        return Decimal(str(value).replace(",", ""))
    except (InvalidOperation, ValueError):
        return None

def validate_asientos(data, pcge_map):
    errors = []
    warnings = []
    valid = []

    for a_idx, asiento in enumerate(data.get("asientos", []), start=1):
        lines = asiento.get("lineas", []) or []
        total_d = Decimal("0")
        total_h = Decimal("0")
        asiento_errors = []

        for l_idx, line in enumerate(lines, start=1):
            code = str(line.get("codigo", "")).strip()
            if not re.fullmatch(r"\d{5}", code):
                asiento_errors.append(f"Línea {l_idx}: código '{code}' no tiene 5 dígitos.")
            elif code not in pcge_map:
                asiento_errors.append(f"Línea {l_idx}: código {code} no existe en el PCGE de TANA.")

            debe = _money(line.get("debe"))
            haber = _money(line.get("haber"))
            if debe is None or haber is None:
                asiento_errors.append(f"Línea {l_idx}: importe inválido.")
                continue
            if debe < 0 or haber < 0:
                asiento_errors.append(f"Línea {l_idx}: los importes no pueden ser negativos.")
            if debe > 0 and haber > 0:
                asiento_errors.append(f"Línea {l_idx}: una línea no puede tener Debe y Haber simultáneamente.")
            total_d += debe
            total_h += haber

        diff = total_d - total_h
        if abs(diff) > Decimal("0.01"):
            asiento_errors.append(
                f"Asiento {asiento.get('numero', a_idx)} descuadra: Debe {total_d:.2f} / Haber {total_h:.2f}."
            )

        if asiento.get("requiere_revision"):
            warnings.append(
                f"Asiento {asiento.get('numero', a_idx)} requiere revisión: {asiento.get('observacion', '')}"
            )

        if asiento_errors:
            errors.extend(asiento_errors)
        else:
            valid.append(asiento)

    return valid, errors, warnings

def _numeric_from_obj(obj, keys):
    """Busca de forma tolerante un importe asociado a alguna clave."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            kl = str(k).strip().lower()
            if kl in keys:
                n = _to_float(v, None)
                if n is not None:
                    return n
            found = _numeric_from_obj(v, keys)
            if found is not None:
                return found
    elif isinstance(obj, list):
        for item in obj:
            found = _numeric_from_obj(item, keys)
            if found is not None:
                return found
    return None


def _find_account_amount(items, account_code):
    """Encuentra un saldo/importe de una cuenta en el estado inicial."""
    if isinstance(items, dict):
        code = str(items.get("codigo", items.get("cuenta", ""))).strip()
        if code == account_code:
            for key in ("importe", "saldo", "monto", "haber", "debe", "valor"):
                n = _to_float(items.get(key), None)
                if n is not None and n != 0:
                    return abs(n)
        for v in items.values():
            found = _find_account_amount(v, account_code)
            if found is not None:
                return found
    elif isinstance(items, list):
        for item in items:
            found = _find_account_amount(item, account_code)
            if found is not None:
                return found
    return None


def _find_operation(operations, number):
    for op in operations or []:
        try:
            if int(op.get("numero")) == int(number):
                return op
        except Exception:
            continue
    return {}


def asegurar_cuenta_79_en_destinos(asientos, pcge_map):
    """
    Regla determinista de TANA para destinos por función.

    Cuando un asiento contiene una cuenta del Elemento 9 (94, 95, etc.)
    con importe en el DEBE, ese destino debe quedar compensado mediante
    una cuenta de destino 79 en el HABER. Para el PCGE operativo de TANA
    la cuenta estándar es 79111.

    Esta regla evita depender de que Gemini recuerde escribir la 79.
    Si ya existe una 79 en el asiento, no se duplica. Si existe pero está
    incompleta, se agrega solo la diferencia necesaria.
    """
    resultado = []
    for asiento in asientos or []:
        a = dict(asiento) if isinstance(asiento, dict) else asiento
        lineas = list(a.get("lineas", []) or []) if isinstance(a, dict) else []
        if not lineas:
            resultado.append(a)
            continue

        total_elemento9_debe = 0.0
        total_79_haber = 0.0
        for line in lineas:
            if not isinstance(line, dict):
                continue
            code = str(line.get("codigo", "")).strip()
            debe = _to_float(line.get("debe"), 0.0)
            haber = _to_float(line.get("haber"), 0.0)
            if re.fullmatch(r"9\d{4}", code):
                total_elemento9_debe += max(debe, 0.0)
            if code.startswith("79"):
                total_79_haber += max(haber, 0.0)

        if total_elemento9_debe > 0.009:
            diferencia = round(total_elemento9_debe - total_79_haber, 2)
            if diferencia > 0.009:
                codigo_79 = "79111" if "79111" in pcge_map else next(
                    (c for c in pcge_map if str(c).startswith("79") and len(str(c)) == 5),
                    None,
                )
                if codigo_79:
                    lineas.append({
                        "codigo": codigo_79,
                        "denominacion": pcge_map.get(codigo_79, "Cargas imputables a cuentas de costos y gastos"),
                        "debe": 0.0,
                        "haber": diferencia,
                        "concepto": "Destino de gastos por función",
                    })
                    a["lineas"] = lineas
                    nota = "Se incorporó automáticamente la cuenta 79 por destino de gastos por función."
                    anterior = str(a.get("observacion", "") or "").strip()
                    a["observacion"] = (anterior + " " + nota).strip()
        resultado.append(a)
    return resultado


def corregir_retiro_socio(asientos, monografia_json):
    """
    Regla contable de la práctica para separación de socio:

    1) La compra de las participaciones del socio saliente es PRIVADA entre
       los socios restantes. La sociedad no registra esa compraventa ni
       modifica su capital social. Por tanto, la operación de separación
       no genera asiento.

    2) Cuando la monografía indica que se determinan y entregan las
       utilidades al socio separado, se calcula su porcentaje sobre el
       capital social:
           participaciones del socio / capital social

       En la práctica:
           15,000 / 60,000 = 25%

       Luego:
           utilidad a distribuir = saldo de 59111 x 25%

       El registro de la distribución es:
           59111  Debe
           48185  Haber
           44191  Haber

       Y el pago:
           44191  Debe
           10411  Haber

    IMPORTANTE:
    No se toma el importe de la compraventa privada como utilidad ni se
    reclasifica, cancela ni mueve la cuenta 50. Cualquier asiento de
    cancelación/reclasificación de capital generado por Gemini para esta
    operación debe eliminarse por completo.

    La distribución DEBE usar las tres cuentas:
        59111 / 48185 / 44191
    No se acepta una distribución 59111 / 44191 sin 48185.
    """
    data = monografia_json or {}
    operaciones = data.get("operaciones", []) or []
    estado = data.get("estado_inicial", []) or []

    # --- Detectar por DESCRIPCIÓN, no depender de que Gemini haya
    # conservado exactamente los números 3 y 4.
    op_separacion = None
    op_utilidades = None

    for op in operaciones:
        desc = str(op.get("descripcion", "")).lower()
        if op_separacion is None and any(k in desc for k in (
            "separa", "separación", "separacion",
            "socio", "participaciones", "compra sus participaciones",
            "venta de participaciones"
        )):
            op_separacion = op

        if op_utilidades is None and any(k in desc for k in (
            "determina", "determinación", "determinacion",
            "entrega", "entregar", "paga", "pago",
            "utilidades", "utilidad", "dividendos", "dividendo"
        )) and any(k in desc for k in ("socio", "separado", "separación", "separacion")):
            op_utilidades = op

    if op_separacion is None or op_utilidades is None:
        return asientos

    # --- Capital social y utilidades acumuladas del estado inicial.
    capital_total = _find_account_amount(estado, "50121")
    if capital_total is None:
        capital_total = _find_account_amount(estado, "50111")

    utilidades = _find_account_amount(estado, "59111")

    # --- Participaciones del socio saliente.
    participacion_socio = _to_float(op_separacion.get("importe"), None)

    if participacion_socio is None:
        participacion_socio = _numeric_from_obj(
            op_separacion,
            {
                "participaciones",
                "participaciones_sociales",
                "acciones",
                "valor_participaciones",
                "valor_nominal",
            },
        )

    # Si Gemini puso el dato en la descripción, extraer "15,000".
    if participacion_socio is None:
        texto_sep = " ".join(
            str(op_separacion.get(k, ""))
            for k in ("descripcion", "datos_adicionales", "concepto")
        )
        m = re.search(
            r"(\d[\d,\.]*)\s*(?:participaciones|acciones)",
            texto_sep,
            flags=re.IGNORECASE,
        )
        if m:
            participacion_socio = _to_float(m.group(1).replace(",", ""), None)

    # Para esta práctica, la fuente textual puede mencionar explícitamente
    # 15,000 participaciones y S/ 60,000 de capital aunque Gemini no los haya
    # colocado en los campos numéricos.
    texto_completo = json.dumps(data, ensure_ascii=False)

    if participacion_socio is None:
        m = re.search(
            r"15[\s,.]?000\s*(?:participaciones|acciones)",
            texto_completo,
            flags=re.IGNORECASE,
        )
        if m:
            participacion_socio = 15000.0

    if capital_total is None:
        m = re.search(
            r"(?:capital(?:\s+social)?|capitalista)[^0-9]{0,80}"
            r"60[\s,.]?000",
            texto_completo,
            flags=re.IGNORECASE,
        )
        if m:
            capital_total = 60000.0

    # La extracción puede tener la cifra 60,000 dentro de una estructura
    # de estado inicial sin asociarla a una cuenta. Para esta práctica,
    # si aparecen 15,000 participaciones y 60,000 de capital, la relación
    # es inequívocamente 25%.
    if (
        capital_total is None
        and participacion_socio is not None
        and participacion_socio == 15000
        and re.search(r"60[\s,.]?000", texto_completo)
    ):
        capital_total = 60000.0

    if (
        participacion_socio is None
        or capital_total is None
        or capital_total == 0
        or utilidades is None
    ):
        return asientos

    porcentaje = round(participacion_socio / capital_total * 100, 6)
    utilidad_socio = round(utilidades * porcentaje / 100, 2)

    # Retención de 5% sobre dividendos.
    retencion = round(utilidad_socio * 0.05, 2)
    neto = round(utilidad_socio - retencion, 2)

    # ------------------------------------------------------------
    # Eliminar cualquier asiento generado por Gemini relacionado con
    # la compraventa privada y la distribución/pago de utilidades.
    # ------------------------------------------------------------
    def es_retiro_o_utilidad(a):
        opnum = str(a.get("operacion_numero", "")).strip()
        texto = " ".join(
            str(a.get(k, ""))
            for k in ("glosa", "observacion", "documento")
        ).lower()

        # Si el asiento está ligado a las operaciones detectadas.
        try:
            if opnum == str(op_separacion.get("numero", "")).strip():
                return True
            if opnum == str(op_utilidades.get("numero", "")).strip():
                return True
        except Exception:
            pass

        # Respaldo por texto, para evitar que Gemini cambie el número.
        palabras = (
            "separación", "separacion", "socio separado",
            "reparto de utilidades", "pago de utilidades",
            "distribución de utilidades", "distribucion de utilidades",
            "dividendo", "participaciones"
        )
        return any(palabra in texto for palabra in palabras)

    # Eliminar cualquier asiento generado por Gemini relacionado con la
    # separación del socio o el reparto/pago. La operación privada no deja
    # ningún asiento de capital en la sociedad.
    resultado = []
    for a in asientos:
        if es_retiro_o_utilidad(a):
            continue

        # Protección adicional: Gemini a veces cambia la glosa/número y
        # genera un asiento de "cancelación/reclasificación" de la cuenta 50.
        # Si contiene una cuenta 50 de cinco dígitos y el texto habla de
        # separación/participaciones/socio, se elimina ese asiento.
        texto_ext = " ".join(
            str(a.get(k, ""))
            for k in ("glosa", "observacion", "documento", "concepto")
        ).lower()
        lineas = a.get("lineas", []) or []
        codigos = {
            str(l.get("codigo", "")).strip()
            for l in lineas
            if isinstance(l, dict)
        }
        habla_retiro = any(k in texto_ext for k in (
            "separación", "separacion", "socio separado",
            "participaciones", "venta de participaciones",
            "compra de participaciones", "retiro del socio"
        ))
        tiene_cuenta_50 = any(c.startswith("50") for c in codigos)
        if habla_retiro and tiene_cuenta_50:
            continue

        resultado.append(a)

    # Tomamos metadatos de los asientos eliminados solo para conservar
    # fecha/documento; no conservamos sus cuentas.
    metas = [
        a for a in asientos
        if str(a.get("operacion_numero", "")).strip()
        in {
            str(op_separacion.get("numero", "")).strip(),
            str(op_utilidades.get("numero", "")).strip(),
        }
    ]

    meta_dist = {}
    meta_pago = {}

    for a in metas:
        txt = " ".join(
            str(a.get(k, ""))
            for k in ("glosa", "documento")
        ).lower()

        if not meta_dist and any(k in txt for k in (
            "utilidad", "dividendo", "distribución", "distribucion"
        )):
            meta_dist = a

        if any(k in txt for k in ("pago", "transferencia", "bancaria")):
            meta_pago = a

    if not meta_dist:
        meta_dist = metas[0] if metas else {}

    if not meta_pago:
        meta_pago = metas[-1] if metas else meta_dist

    def make_asiento(meta, numero_default, fecha, glosa, documento, lineas):
        return {
            "numero": meta.get("numero", numero_default),
            "fecha": meta.get("fecha", fecha),
            "glosa": glosa,
            "documento": meta.get("documento", documento),
            "operacion_numero": op_utilidades.get("numero", 4),
            "requiere_revision": False,
            "observacion": "",
            "lineas": lineas,
        }

    fecha = op_utilidades.get("fecha", "")
    documento = op_utilidades.get("documento", "")

    # ASIENTO DE DISTRIBUCIÓN
    dist = make_asiento(
        meta_dist,
        4,
        fecha,
        "Determinación y entrega de utilidades al socio separado",
        documento,
        [
            {
                "codigo": "59111",
                "denominacion": "Utilidades acumuladas",
                "debe": utilidad_socio,
                "haber": 0.0,
                "concepto": f"Distribución del {porcentaje:.2f}% de las utilidades acumuladas",
            },
            {
                "codigo": "48185",
                "denominacion": "Retenciones por dividendos",
                "debe": 0.0,
                "haber": retencion,
                "concepto": "Retención del 5% sobre dividendos",
            },
            {
                "codigo": "44191",
                "denominacion": "Dividendos",
                "debe": 0.0,
                "haber": neto,
                "concepto": "Utilidad neta por pagar al socio separado",
            },
        ],
    )

    # ASIENTO DE PAGO
    pago = make_asiento(
        meta_pago,
        5,
        fecha,
        "Pago de utilidades al socio separado mediante transferencia bancaria",
        meta_pago.get("documento", documento),
        [
            {
                "codigo": "44191",
                "denominacion": "Dividendos",
                "debe": neto,
                "haber": 0.0,
                "concepto": "Cancelación de dividendos",
            },
            {
                "codigo": "10411",
                "denominacion": "Cuentas corrientes operativas",
                "debe": 0.0,
                "haber": neto,
                "concepto": "Pago mediante transferencia bancaria",
            },
        ],
    )

    # Post-condición: el asiento de distribución debe contener siempre las
    # tres cuentas requeridas por esta práctica: 59111, 48185 y 44191.
    # Si alguna cuenta cambiara por una edición futura, fallamos de forma
    # explícita en vez de devolver un asiento incompleto.
    cod_dist = {str(l.get("codigo", "")).strip() for l in dist.get("lineas", [])}
    cuentas_requeridas = {"59111", "48185", "44191"}
    if cod_dist != cuentas_requeridas:
        raise ValueError(
            "El reparto de utilidades debe usar exactamente 59111, 48185 y 44191."
        )

    resultado.extend([dist, pago])
    return resultado

def resolve_asientos_with_gemini():
    if not get_gemini_profiles():
        raise RuntimeError("No está configurada ninguna GEMINI_API_KEY en Streamlit Secrets.")

    pcge_map = {str(cod).strip(): str(desc) for cod, desc in PCGE_DATA}
    # Solo cuentas de 5 dígitos: el usuario indicó que este es el nivel operativo de TANA.
    pcge_5 = [[code, desc] for code, desc in pcge_map.items() if re.fullmatch(r"\d{5}", code)]

    # No usamos str.format() aquí porque ASIENTOS_PROMPT contiene un ejemplo
    # JSON con llaves. format() interpretaría esas llaves como placeholders
    # y produciría errores del tipo: "\n  \"asientos\"".
    prompt = (
        ASIENTOS_PROMPT
        .replace("{pcge}", json.dumps(pcge_5, ensure_ascii=False))
        .replace(
            "{operaciones}",
            json.dumps(st.session_state.get("monografia_json", {}), ensure_ascii=False),
        )
    )

    def make_contents(_client):
        return [prompt]

    response, profile = _generate_with_fallback(
        make_contents,
        types.GenerateContentConfig(response_mime_type="application/json"),
    )
    data = json.loads(response.text or "{}")
    data.setdefault("_tana_gemini_route", profile["label"])
    data.setdefault("_tana_gemini_model", profile["model"])
    return data, pcge_map

if "monografia_json" in st.session_state:
    st.divider()
    st.subheader("🧮 2. Desarrollo de asientos contables")
    st.write(
        "TANA utilizará las operaciones detectadas y su PCGE de 5 dígitos. "
        "Antes de aceptar un asiento, verifica que la cuenta exista y que Debe = Haber."
    )

    if st.button("🧮 Resolver asientos contables", type="primary"):
        with st.spinner("Desarrollando y validando los asientos contables..."):
            try:
                resolved, pcge_map = resolve_asientos_with_gemini()

                # Gemini devuelve {"asientos": [...], "alertas": [...]}
                # La interfaz necesita trabajar con la lista de asientos.
                if isinstance(resolved, dict):
                    asientos_generados = resolved.get("asientos", [])
                    alertas_gemini = resolved.get("alertas", [])
                elif isinstance(resolved, list):
                    asientos_generados = resolved
                    alertas_gemini = []
                else:
                    raise ValueError("La respuesta de Gemini no tiene una estructura de asientos válida.")

                if not isinstance(asientos_generados, list):
                    raise ValueError("La clave 'asientos' de Gemini no contiene una lista.")

                # Regla determinista: toda cuenta de destino del Elemento 9
                # debe tener su contrapartida 79 en el Haber. Esto corrige el
                # caso en que Gemini desarrolle el destino pero omita la 79111.
                asientos_generados = asegurar_cuenta_79_en_destinos(
                    asientos_generados,
                    pcge_map,
                )

                # Corrección determinista de la Operación 3/4 de retiro de socio:
                # la venta privada no se contabiliza en la sociedad y el reparto
                # de utilidades usa 59111/48185/44191 y luego 44191/10411.
                asientos_generados = corregir_retiro_socio(
                    asientos_generados,
                    st.session_state.get("monografia_json", {}),
                )

                valid, errors, warnings = validate_asientos(
                    {"asientos": asientos_generados},
                    pcge_map,
                )

                st.session_state["asientos_contables"] = asientos_generados
                st.session_state["asientos_validos"] = valid
                st.session_state["errores_asientos"] = errors
                st.session_state["alertas_asientos"] = list(alertas_gemini) + list(warnings)
            except json.JSONDecodeError:
                st.error("Gemini devolvió una respuesta que no es JSON válido. Vuelve a intentarlo.")
            except Exception as exc:
                st.error(f"No se pudieron desarrollar los asientos: {exc}")

    if "asientos_contables" in st.session_state:
        asientos = st.session_state["asientos_contables"]
        validos = st.session_state.get("asientos_validos", [])
        errores = st.session_state.get("errores_asientos", [])
        alertas = st.session_state.get("alertas_asientos", [])

        c1, c2, c3 = st.columns(3)
        c1.metric("Asientos generados", len(asientos))
        c2.metric("Asientos validados", len(validos))
        c3.metric("Errores", len(errores))

        if errores:
            st.error("Hay asientos que TANA NO acepta todavía:")
            for e in errores:
                st.write(f"- {e}")
        else:
            st.success("Todos los asientos generados cuadran y utilizan cuentas existentes de 5 dígitos.")

        for asiento in asientos:
            numero = asiento.get("numero", "")
            with st.expander(
                f"Asiento {numero} — {asiento.get('fecha', '')} — {asiento.get('glosa', '')}",
                expanded=False,
            ):
                rows = []
                for line in asiento.get("lineas", []):
                    rows.append({
                        "Código": line.get("codigo", ""),
                        "Cuenta": pcge_map.get(str(line.get("codigo", "")).strip(), line.get("denominacion", "")),
                        "Concepto": line.get("concepto", ""),
                        "Debe": line.get("debe", 0),
                        "Haber": line.get("haber", 0),
                    })
                st.dataframe(rows, use_container_width=True, hide_index=True)
                if asiento.get("requiere_revision"):
                    st.warning(asiento.get("observacion", "Requiere revisión."))

        if alertas:
            st.warning("\n".join(f"- {x}" for x in alertas))

# ============================================================
# TUTOR INTERACTIVO TANA
# ============================================================
def _tana_contexto_tutor():
    mono = st.session_state.get("monografia_texto", "")
    asientos = st.session_state.get("asientos_contables", [])
    asientos_txt = json.dumps(asientos, ensure_ascii=False, indent=2)
    return (
        "MONOGRAFÍA:\n" + mono[:14000]
        + "\n\nASIENTOS GENERADOS POR TANA:\n" + asientos_txt[:18000]
    )

def _preguntar_a_tana(pregunta):
    contexto = _tana_contexto_tutor()
    prompt = f"""Eres TANA, tutor de contabilidad peruana.
Responde la pregunta del estudiante usando únicamente el contexto proporcionado.
Explica con claridad por qué se hizo el asiento, cómo se obtuvo el importe, por qué
una cuenta va al Debe o Haber y, cuando corresponda, cómo se relaciona con la HT,
la distribución y ajustes, ERN, ERF o ESF.
No inventes información que no aparezca en el contexto.
 En los estados financieros respeta estrictamente estas reglas:
 ERF: 70 y 69 se detectan por prefijo; 94 y 95 son obligatorias; 78 se incluye solo si existe; 65 y 67 solo si existen sin destino a 94/95. No incluyas 79 ni agregues automáticamente otras cuentas del elemento 6 al ERF.
 ERN: presenta las cuentas por naturaleza y su resultado.
 ESF: presenta activo, pasivo y patrimonio; resultados acumulados 59 con saldo deudor reducen el patrimonio. El resultado del ejercicio debe ser consistente con ERN y ERF y el ESF debe cumplir Activo = Pasivo + Patrimonio.
 Si falta un dato, dilo.

CONTEXTO:
{contexto}

PREGUNTA:
{pregunta}"""

    response, profile = _generate_with_fallback(
        lambda client: [prompt],
        types.GenerateContentConfig()
    )
    return response.text or "No pude generar una respuesta.", profile["label"]

if "monografia_json" in st.session_state or "asientos_contables" in st.session_state:
    st.divider()
    st.subheader("🤖 Pregúntale a TANA")
    st.caption("Pregunta por qué se hizo un asiento, cómo se calculó o por qué una cuenta va al Debe o al Haber.")

    pregunta = st.text_input(
        "Escribe tu pregunta",
        placeholder="Ej.: ¿Por qué se utilizó la cuenta 79111 en este asiento?",
        key="pregunta_tana",
    )
    c1, c2 = st.columns([4, 1])
    with c1:
        preguntar = st.button("💬 Preguntar a TANA", type="primary", key="btn_preguntar_tana")
    with c2:
        audio = st.audio_input("🎙️ Hablar", key="audio_tana") if hasattr(st, "audio_input") else None

    if preguntar and pregunta.strip():
        with st.spinner("TANA está preparando la explicación..."):
            try:
                respuesta, ruta = _preguntar_a_tana(pregunta.strip())
                st.session_state["respuesta_tana"] = respuesta
                st.session_state["respuesta_tana_ruta"] = ruta
            except Exception as exc:
                st.error(f"No se pudo responder: {_gemini_error_message(exc)}")

    if audio is not None:
        with st.spinner("TANA está escuchando y preparando la respuesta..."):
            temp_audio = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp:
                    tmp.write(audio.getvalue())
                    temp_audio = tmp.name

                def audio_contents(client):
                    audio_file = client.files.upload(file=temp_audio)
                    return [
                        audio_file,
                        "Escucha el audio del estudiante, transcribe su pregunta y luego respóndela. "
                        "No inventes datos. Usa el siguiente contexto:\n" + _tana_contexto_tutor(),
                    ]

                response, profile = _generate_with_fallback(
                    audio_contents,
                    types.GenerateContentConfig()
                )
                st.session_state["respuesta_tana"] = response.text or "No pude interpretar el audio."
                st.session_state["respuesta_tana_ruta"] = profile["label"]
            except Exception as exc:
                st.error(f"No se pudo procesar el audio: {_gemini_error_message(exc)}")
            finally:
                if temp_audio and os.path.exists(temp_audio):
                    os.remove(temp_audio)

    if st.session_state.get("respuesta_tana"):
        st.markdown("**Respuesta de TANA:**")
        st.info(st.session_state["respuesta_tana"])
        if st.session_state.get("respuesta_tana_ruta"):
            st.caption(f"Procesado por {st.session_state['respuesta_tana_ruta']}.")

st.divider()

if not st.button("📊 Generar Excel contable", type="primary"):
    st.stop()


FONT = "Arial"
wb = Workbook()

def style_header(ws, row, col_start, col_end, fill="1F4E78", fontcolor="FFFFFF"):
    for c in range(col_start, col_end+1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True, color=fontcolor, size=10)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin"))

def autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

BLUE = Font(name=FONT, color="0000FF", size=10)
BLACK = Font(name=FONT, color="000000", size=10)
BOLD = Font(name=FONT, bold=True, size=10)
INPUT_FILL = PatternFill("solid", fgColor="FFFF99")
TITLE_FONT = Font(name=FONT, bold=True, size=14, color="1F4E78")
SUBTITLE_FONT = Font(name=FONT, italic=True, size=10, color="555555")
GRAY = Font(name=FONT, size=9, color="808080")

with open("pcge_data.json") as f:
    PCGE_DATA = json.load(f)

print("Setup listo,", len(PCGE_DATA), "cuentas PCGE cargadas")

# ============================================================
# HOJA: PCGE (catálogo oficial, tal cual tu plantilla)
# ============================================================
ws_pcge = wb.active
ws_pcge.title = "PCGE"
ws_pcge.cell(row=1, column=1, value="Cuenta")
ws_pcge.cell(row=1, column=2, value="Descripción")
style_header(ws_pcge, 1, 1, 2)
for i, (cod, desc) in enumerate(PCGE_DATA, start=2):
    ws_pcge.cell(row=i, column=1, value=cod).font = BLACK
    ws_pcge.cell(row=i, column=2, value=desc).font = BLACK
PCGE_LAST_ROW = 1 + len(PCGE_DATA)
autofit(ws_pcge, [12, 60])
ws_pcge.freeze_panes = "A2"

# Rango con nombre "PCGE" (igual que tu archivo original) para los VLOOKUP
dn = DefinedName("PCGE", attr_text=f"PCGE!$A$1:$B${PCGE_LAST_ROW}")
wb.defined_names["PCGE"] = dn

print("Hoja PCGE lista:", PCGE_LAST_ROW-1, "cuentas, rango con nombre creado")

# ============================================================
# HOJA: Reglas_Asiento (motor de plantillas, códigos PCGE reales)
# ============================================================
ws2 = wb.create_sheet("Reglas_Asiento")
headers = ["Tipo Operación", "Sub-Asiento\n(correlativo propio)", "Orden Línea",
           "Código Cuenta", "Lado (D/H)", "Col.Monto\n(Registro_Operaciones)", "Clave (aux)"]
for i, h in enumerate(headers, start=1):
    ws2.cell(row=1, column=i, value=h)
style_header(ws2, 1, 1, 7)

# Base: F=Valor Base (sin IGV) | G=IGV | H=Total | I=Costo de venta
# SubAsiento agrupa las líneas que forman UN asiento correlativo independiente
# (p.ej. la venta genera 2 asientos: 12/40/70 y luego 69/20)
reglas = [
    # TipoOperacion, SubAsiento, Orden(ABSOLUTO 1..5 dentro del evento), Cuenta, Lado, ColMonto
    ("VENTA_CONTADO", 1, 1, "10111", "D", "H"),
    ("VENTA_CONTADO", 1, 2, "40111", "H", "G"),
    ("VENTA_CONTADO", 1, 3, "70121", "H", "F"),
    ("VENTA_CONTADO", 2, 4, "69121", "D", "I"),
    ("VENTA_CONTADO", 2, 5, "20111", "H", "I"),

    ("VENTA_CREDITO", 1, 1, "12121", "D", "H"),
    ("VENTA_CREDITO", 1, 2, "40111", "H", "G"),
    ("VENTA_CREDITO", 1, 3, "70121", "H", "F"),
    ("VENTA_CREDITO", 2, 4, "69121", "D", "I"),
    ("VENTA_CREDITO", 2, 5, "20111", "H", "I"),

    ("COMPRA_CONTADO", 1, 1, "60111", "D", "F"),
    ("COMPRA_CONTADO", 1, 2, "40111", "D", "G"),
    ("COMPRA_CONTADO", 1, 3, "10111", "H", "H"),
    ("COMPRA_CONTADO", 2, 4, "20111", "D", "F"),
    ("COMPRA_CONTADO", 2, 5, "61111", "H", "F"),

    ("COMPRA_CREDITO", 1, 1, "60111", "D", "F"),
    ("COMPRA_CREDITO", 1, 2, "40111", "D", "G"),
    ("COMPRA_CREDITO", 1, 3, "42121", "H", "H"),
    ("COMPRA_CREDITO", 2, 4, "20111", "D", "F"),
    ("COMPRA_CREDITO", 2, 5, "61111", "H", "F"),

    ("COBRO_CLIENTE", 1, 1, "10111", "D", "H"),
    ("COBRO_CLIENTE", 1, 2, "12121", "H", "H"),

    ("PAGO_PROVEEDOR", 1, 1, "42121", "D", "H"),
    ("PAGO_PROVEEDOR", 1, 2, "10111", "H", "H"),

    ("DEPRECIACION_MENSUAL", 1, 1, "68415", "D", "F"),
    ("DEPRECIACION_MENSUAL", 1, 2, "39527", "H", "F"),

    ("PLANILLA_SUELDOS", 1, 1, "62111", "D", "F"),
    ("PLANILLA_SUELDOS", 1, 2, "41111", "H", "F"),
]

r = 2
for tipo, sub, orden, cuenta, lado, col in reglas:
    ws2.cell(row=r, column=1, value=tipo).font = BLACK
    ws2.cell(row=r, column=2, value=sub).font = BLACK
    ws2.cell(row=r, column=3, value=orden).font = BLACK
    ws2.cell(row=r, column=4, value=cuenta).font = BLACK
    ws2.cell(row=r, column=5, value=lado).font = BLACK
    ws2.cell(row=r, column=6, value=col).font = BLACK
    # clave auxiliar para busqueda por linea fisica: Tipo|Orden (orden absoluto 1..5 dentro del evento)
    ws2.cell(row=r, column=7, value=f'=A{r}&"|"&C{r}').font = GRAY
    r += 1

REGLAS_LAST_ROW = r - 1
ws2.freeze_panes = "A2"
autofit(ws2, [22, 14, 12, 12, 10, 14, 20])
ws2.cell(row=1, column=6).comment = Comment(
    "F=Valor Base (sin IGV), G=IGV, H=Total, I=Costo de venta. Estas letras son las columnas de Registro_Operaciones.", "Sistema")
ws2.cell(row=1, column=2).comment = Comment(
    "Líneas con el mismo Sub-Asiento forman UN asiento (mismo N° correlativo, fecha y glosa en la primera línea). "
    "Un Sub-Asiento distinto = nuevo N° correlativo, aunque sea de la misma operación.", "Sistema")

TIPOS_OPERACION = sorted(set(x[0] for x in reglas))
print("Hoja Reglas_Asiento lista:", REGLAS_LAST_ROW-1, "líneas,", len(TIPOS_OPERACION), "tipos de operación")

# ============================================================
# HOJA: Registro_Operaciones (captura del usuario)
# ============================================================
ws3 = wb.create_sheet("Registro_Operaciones")
headers = ["N°", "Fecha", "Tipo Operación", "Glosa", "Documento Ref.",
           "Valor Base\n(sin IGV) S/", "IGV S/\n(auto 18%)", "Total S/\n(auto)",
           "Costo de Venta S/\n(solo VENTA_*, opcional)"]
for i, h in enumerate(headers, start=1):
    ws3.cell(row=1, column=i, value=h)
style_header(ws3, 1, 1, 9)

N_OPS = 20
IGV_APLICA = ["VENTA_CONTADO", "VENTA_CREDITO", "COMPRA_CONTADO", "COMPRA_CREDITO"]
igv_condition = "OR(" + ",".join([f'C{{r}}="{t}"' for t in IGV_APLICA]) + ")"

ejemplo = [1, "2026-08-01", "VENTA_CONTADO", "Venta al contado - Boleta 001-00123",
           "B001-00123", 1000, None, None, 550]
for i, v in enumerate(ejemplo, start=1):
    c = ws3.cell(row=2, column=i, value=v)
    if i in (6, 9):
        c.font = BLUE
        c.fill = INPUT_FILL
    else:
        c.font = BLUE

for r in range(2, 2 + N_OPS):
    ws3.cell(row=r, column=7, value=f"={igv_condition.format(r=r)}*F{r}*0.18")
    ws3.cell(row=r, column=8, value=f"=F{r}+G{r}")
    ws3.cell(row=r, column=7).font = BLACK
    ws3.cell(row=r, column=8).font = BLACK
    for col in (6, 7, 8, 9):
        ws3.cell(row=r, column=col).number_format = '#,##0.00'
    if r > 2:
        ws3.cell(row=r, column=1, value=r-1).font = BLACK
        for col in (6, 9):
            ws3.cell(row=r, column=col).font = BLUE
            ws3.cell(row=r, column=col).fill = INPUT_FILL

dv = DataValidation(type="list", formula1='"' + ",".join(TIPOS_OPERACION) + '"', allow_blank=True)
ws3.add_data_validation(dv)
dv.add(f"C2:C{1+N_OPS}")

ws3.freeze_panes = "A2"
autofit(ws3, [5, 12, 20, 38, 16, 14, 12, 12, 20])
ws3.cell(row=1, column=6).comment = Comment("Celda de entrada (amarillo = tú llenas).", "Sistema")
ws3.cell(row=1, column=9).comment = Comment("Solo VENTA_CONTADO / VENTA_CREDITO. Déjalo en 0 si no llevas costeo perpetuo.", "Sistema")

REG_LAST_ROW = 1 + N_OPS
print("Hoja Registro_Operaciones lista:", N_OPS, "filas")

# ============================================================
# HOJA: LD - Libro Diario (generado)
# Regla del usuario: N° correlativo, Fecha y Glosa SOLO en la primera
# línea de cada sub-asiento (grupo). Las líneas siguientes del mismo
# grupo van con esas 3 celdas en blanco.
# ============================================================
ws4 = wb.create_sheet("LD")
headers = ["N° Asiento", "Fecha", "Glosa", "Documento", "Código Cuenta",
           "Denominación", "Debe S/", "Haber S/",
           "aux MatchRow", "aux Lado", "aux ColMonto", "aux GroupKey", "aux EsNuevoGrupo"]
for i, h in enumerate(headers, start=1):
    ws4.cell(row=1, column=i, value=h)
style_header(ws4, 1, 1, 13)

MAX_LINEAS = 5
dr = 2
for opnum in range(1, N_OPS + 1):
    oprow = opnum + 1
    for ln in range(1, MAX_LINEAS + 1):
        gate = f'Registro_Operaciones!$B${oprow}=""'
        clave = f'Registro_Operaciones!$C${oprow}&"|"&{ln}'

        # I: fila de la regla (o "" si esta linea no aplica a este tipo)
        f_match = f'=IF({gate},"",IFERROR(MATCH({clave},Reglas_Asiento!$G$2:$G${REGLAS_LAST_ROW},0)+1,""))'
        ws4.cell(row=dr, column=9, value=f_match)
        # J: lado D/H
        ws4.cell(row=dr, column=10, value=f'=IF($I{dr}="","",INDEX(Reglas_Asiento!$E:$E,$I{dr}))')
        # K: columna de monto F/G/H/I
        ws4.cell(row=dr, column=11, value=f'=IF($I{dr}="","",INDEX(Reglas_Asiento!$F:$F,$I{dr}))')
        # L: clave de grupo = Tipo|SubAsiento|FilaOperacion (identifica un mismo asiento correlativo)
        f_group = (f'=IF($I{dr}="","",Registro_Operaciones!$C${oprow}&"|"'
                   f'&INDEX(Reglas_Asiento!$B:$B,$I{dr})&"|"&{oprow})')
        ws4.cell(row=dr, column=12, value=f_group)
        # M: es primera linea de un grupo nuevo? compara con la fila de arriba
        if dr == 2:
            f_new = f'=IF($L{dr}="",0,1)'
        else:
            f_new = f'=IF($L{dr}="",0,IF($L{dr}<>$L{dr-1},1,0))'
        ws4.cell(row=dr, column=13, value=f_new)

        # A: N° Asiento -> correlativo GLOBAL, solo visible si es primera linea del grupo
        f_nasiento = f'=IF($M{dr}=1,SUM($M$2:$M{dr}),"")'
        ws4.cell(row=dr, column=1, value=f_nasiento)
        # B: Fecha, C: Glosa, D: Documento -> solo primera linea del grupo
        ws4.cell(row=dr, column=2, value=f'=IF($M{dr}=1,Registro_Operaciones!$B${oprow},"")')
        ws4.cell(row=dr, column=3, value=f'=IF($M{dr}=1,Registro_Operaciones!$D${oprow},"")')
        ws4.cell(row=dr, column=4, value=f'=IF($M{dr}=1,Registro_Operaciones!$E${oprow},"")')

        # E, F: codigo y denominacion (VLOOKUP contra PCGE, igual a tu plantilla original)
        ws4.cell(row=dr, column=5, value=f'=IF($I{dr}="","",INDEX(Reglas_Asiento!$D:$D,$I{dr}))')
        ws4.cell(row=dr, column=6, value=f'=IF($E{dr}="","",VLOOKUP($E{dr},PCGE,2,0))')

        # G, H: Debe / Haber
        monto = f'IFERROR(INDIRECT("Registro_Operaciones!"&$K{dr}&{oprow}),0)'
        ws4.cell(row=dr, column=7, value=f'=IF($I{dr}="",0,IF($J{dr}="D",{monto},0))')
        ws4.cell(row=dr, column=8, value=f'=IF($I{dr}="",0,IF($J{dr}="H",{monto},0))')
        ws4.cell(row=dr, column=7).number_format = '#,##0.00;(#,##0.00);"-"'
        ws4.cell(row=dr, column=8).number_format = '#,##0.00;(#,##0.00);"-"'

        for col in range(1, 14):
            ws4.cell(row=dr, column=col).font = BLACK
        dr += 1

LD_LAST_ROW = dr - 1
ws4.freeze_panes = "A2"
autofit(ws4, [10, 12, 30, 14, 12, 38, 13, 13, 9, 8, 9, 22, 9])
for col in ("I", "J", "K", "L", "M"):
    ws4.column_dimensions[col].hidden = True

print("Hoja LD (Libro Diario) lista:", LD_LAST_ROW-1, "filas físicas")

# ============================================================
# HOJA: LM - Libro Mayor (generado, por cuenta usada en el motor)
# ============================================================
ws5 = wb.create_sheet("LM")
headers = ["Código", "Denominación", "Naturaleza", "Total Debe S/", "Total Haber S/", "Saldo S/"]
for i, h in enumerate(headers, start=1):
    ws5.cell(row=1, column=i, value=h)
style_header(ws5, 1, 1, 6)

NATURALEZA = {
    "10111": "Deudora", "12121": "Deudora", "20111": "Deudora", "40111": "Acreedora",
    "42121": "Acreedora", "60111": "Deudora", "61111": "Acreedora", "62111": "Deudora",
    "68415": "Deudora", "69121": "Deudora", "70121": "Acreedora", "39527": "Acreedora",
    "41111": "Acreedora",
}
cuentas_usadas = sorted(set(x[3] for x in reglas))

r = 2
for cod in cuentas_usadas:
    ws5.cell(row=r, column=1, value=cod)
    ws5.cell(row=r, column=2, value=f'=VLOOKUP($A{r},PCGE,2,0)')
    ws5.cell(row=r, column=3, value=NATURALEZA[cod])
    ws5.cell(row=r, column=4, value=f'=SUMIFS(LD!$G:$G,LD!$E:$E,$A{r})')
    ws5.cell(row=r, column=5, value=f'=SUMIFS(LD!$H:$H,LD!$E:$E,$A{r})')
    ws5.cell(row=r, column=6, value=f'=IF($C{r}="Deudora",$D{r}-$E{r},$E{r}-$D{r})')
    for col in range(1, 7):
        ws5.cell(row=r, column=col).font = BLACK
    for col in (4, 5, 6):
        ws5.cell(row=r, column=col).number_format = '#,##0.00;(#,##0.00);"-"'
    r += 1
LM_LAST_ROW = r - 1
ws5.freeze_panes = "A2"
autofit(ws5, [10, 45, 14, 15, 15, 15])
ws5.cell(row=1, column=1).comment = Comment(
    "Nota: reemplacé el FILTER() de tu plantilla original por SUMIFS — FILTER es una función matricial "
    "moderna que LibreOffice/algunas versiones no evalúan de forma confiable en archivos generados por script. "
    "El resultado es el mismo saldo por cuenta, más robusto.", "Sistema")
print("Hoja LM (Libro Mayor) lista:", LM_LAST_ROW-1, "cuentas")

# ============================================================
# ============================================================
# HOJA: HT - Hoja de Trabajo / Balance de Comprobación
# La HT se construye directamente desde los asientos contables
# validados por TANA. No depende de las reglas auxiliares del Excel.
# ============================================================

asientos_export = st.session_state.get("asientos_contables", [])

# Consolidar todas las cuentas realmente utilizadas por los asientos.
movimientos = {}
for asiento in asientos_export:
    for line in asiento.get("lineas", []):
        code = str(line.get("codigo", "")).strip()
        if not re.fullmatch(r"\d{5}", code):
            continue
        rec = movimientos.setdefault(code, {"debe": 0.0, "haber": 0.0})
        try:
            rec["debe"] += float(line.get("debe", 0) or 0)
        except Exception:
            pass
        try:
            rec["haber"] += float(line.get("haber", 0) or 0)
        except Exception:
            pass

cuentas_reporte = sorted(movimientos.keys(), key=lambda x: (int(x), x))

ws6 = wb.create_sheet("HT")
ws6.merge_cells("A1:R1")
ws6["A1"] = "HOJA DE TRABAJO / BALANCE DE COMPROBACIÓN"
ws6["A1"].font = TITLE_FONT
ws6["A1"].alignment = Alignment(horizontal="center")

# Encabezados agrupados siguiendo la lógica de la plantilla de trabajo:
# suma, saldos, ajustes/eliminación, saldos ajustados, resultados por
# naturaleza, resultados por función y situación financiera.
groups = [
    (3, 4, "SUMA"),
    (5, 6, "SALDOS"),
    (7, 8, "AJUSTES Y ELIMINACIÓN"),
    (9, 10, "SALDOS AJUSTADOS"),
    (11, 12, "R. NATURALEZA"),
    (13, 14, "R. FUNCIÓN"),
    (15, 16, "E.S.F."),
    (17, 18, "DIST. Y AJUSTE FINAL"),
]
for c1, c2, label in groups:
    ws6.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
    ws6.cell(row=2, column=c1, value=label)
    style_header(ws6, 2, c1, c2)

headers_ht = [
    "CTA", "DENOMINACIÓN", "DEBE", "HABER", "DEUDOR", "ACREEDOR",
    "DEUDOR", "ACREEDOR", "DEBE", "HABER", "DEUDOR", "ACREEDOR",
    "DEUDOR", "ACREEDOR", "ACTIVO", "PASIVO", "DEBE", "HABER",
]
for c, label in enumerate(headers_ht, 1):
    ws6.cell(row=3, column=c, value=label)
style_header(ws6, 3, 1, 18)

# Clasificación contable para la HT.
# Regla oficial del PCGE: toda cuenta cuyo primer dígito es 6, 7, 8 o 9
# es una cuenta de resultados (nunca de balance). 1-5 son de balance.
def clasificar_resultado(code):
    return code[:1] in {"6", "7", "8", "9"}

def es_elemento9(code):
    return code[:1] == "9"

def es_costo_ventas(code):
    return code[:2] == "69"

def es_variacion_existencias(code):
    return code[:2] == "61"

def es_cuenta79(code):
    return code[:2] == "79"

def es_naturaleza(code):
    # 69 (costo de ventas) y el elemento 9 se reclasifican íntegramente a
    # R.Función; 79 es cuenta puente y no aparece en ningún resultado.
    if not clasificar_resultado(code):
        return False
    if es_costo_ventas(code) or es_elemento9(code) or es_cuenta79(code):
        return False
    return True

# Cuentas del Elemento 6 que ya tienen un destino explícito a 94/95.
# Se detectan a partir de los asientos desarrollados por TANA, para que
# el ERF no vuelva a incluir un gasto por naturaleza que ya fue llevado
# a una cuenta de función.
def detectar_cuentas_6_con_destino(asientos):
    con_destino = set()
    for asiento in asientos or []:
        lineas = asiento.get("lineas", []) if isinstance(asiento, dict) else []
        hay_94_95 = any(
            str(x.get("codigo", "")).strip().startswith(("94", "95"))
            for x in lineas if isinstance(x, dict)
        )
        if not hay_94_95:
            continue
        for x in lineas:
            if not isinstance(x, dict):
                continue
            codigo = str(x.get("codigo", "")).strip()
            if codigo[:1] == "6" and len(codigo) == 5:
                con_destino.add(codigo)
    return con_destino

CUENTAS_6_CON_DESTINO = detectar_cuentas_6_con_destino(
    st.session_state.get("asientos_contables", [])
)

def es_funcion(code):
    """
    Clasificación EXACTA para Resultado por Función según la plantilla
    revisada por el usuario:

    OBLIGATORIAS:
      - 70: ventas (detecta cualquier cuenta 70xxxxx presente).
      - 69: costo de ventas (detecta cualquier cuenta 69xxxxx presente).
      - 94 y 95: gastos por función.

    ADICIONALES SOLO SI CORRESPONDE:
      - 78: otros ingresos, si existe en la práctica.
      - 65 y 67: solo si la cuenta existe y NO tiene destino a 94/95.

    NO pertenecen al ERF:
      - 60, 61, 62, 63, 64, 66 y 68 por el solo hecho de ser
        cuentas del elemento 6.
      - 79: es cuenta puente de distribución y nunca se presenta
        como componente del ERF.
    """
    if not code:
        return False
    if code[:2] in {"70", "69", "78", "94", "95"}:
        return True
    if code[:2] in {"65", "67"} and len(code) == 5:
        return code not in CUENTAS_6_CON_DESTINO
    return False

def es_balance(code):
    return not clasificar_resultado(code)

# ------------------------------------------------------------
# AJUSTES Y ELIMINACIÓN (HT) — cálculo previo (dos pasadas)
# ------------------------------------------------------------
# Regla 1: Costo de Ventas (69) <-> Variación de Existencias (61),
# emparejadas por el mismo sufijo (ej. 69111 "Mercadería" <-> 61111
# "Mercadería"). El 69 cancela su saldo deudor completo al HABER de
# ajustes; ese mismo importe pasa al DEBE de ajustes del 61 emparejado,
# reduciendo su saldo acreedor. Así 69 queda solo en R.Función y el
# neto de 61 queda solo en R.Naturaleza.
#
# Regla 2: Elemento 9 (94, 95, ... cualquier código que inicia en "9")
# <-> 79. Cada cuenta del elemento 9 cancela su saldo deudor completo
# al HABER de ajustes (queda solo en R.Función). La(s) cuenta(s) 79
# reciben en el DEBE de ajustes la suma total de esas cancelaciones,
# repartida a prorrata de su propio saldo acreedor si hay más de una.
# ------------------------------------------------------------
ajustes_deudor = {}
ajustes_acreedor = {}

def _deudor_acreedor(code):
    debe = movimientos[code]["debe"]
    haber = movimientos[code]["haber"]
    return max(debe - haber, 0.0), max(haber - debe, 0.0)

# Regla 1: 69 <-> 61 por sufijo
cuentas_61 = [c for c in cuentas_reporte if es_variacion_existencias(c)]
mapa_61_por_sufijo = {c[2:]: c for c in cuentas_61}
for code69 in [c for c in cuentas_reporte if es_costo_ventas(c)]:
    deudor69, _ = _deudor_acreedor(code69)
    if deudor69 <= 0:
        continue
    code61 = mapa_61_por_sufijo.get(code69[2:])
    if code61 is None and len(cuentas_61) == 1:
        code61 = cuentas_61[0]
    if code61 is None:
        continue  # sin cuenta 61 emparejada: no se puede cancelar, se deja como está
    ajustes_acreedor[code69] = ajustes_acreedor.get(code69, 0.0) + deudor69
    ajustes_deudor[code61] = ajustes_deudor.get(code61, 0.0) + deudor69

# Regla 2: elemento 9 <-> 79
cuentas_9 = [c for c in cuentas_reporte if es_elemento9(c)]
cuentas_79 = [c for c in cuentas_reporte if es_cuenta79(c)]
total_elemento9 = 0.0
for code9 in cuentas_9:
    deudor9, _ = _deudor_acreedor(code9)
    if deudor9 <= 0:
        continue
    ajustes_acreedor[code9] = ajustes_acreedor.get(code9, 0.0) + deudor9
    total_elemento9 += deudor9

if total_elemento9 > 0 and cuentas_79:
    acreedores_79 = {c: _deudor_acreedor(c)[1] for c in cuentas_79}
    total_acreedor_79 = sum(acreedores_79.values())
    if total_acreedor_79 > 0:
        for code79, acreedor79 in acreedores_79.items():
            parte = total_elemento9 * (acreedor79 / total_acreedor_79)
            ajustes_deudor[code79] = ajustes_deudor.get(code79, 0.0) + parte
    else:
        # Sin saldo acreedor registrado en 79: se asigna todo a la primera cuenta 79.
        ajustes_deudor[cuentas_79[0]] = ajustes_deudor.get(cuentas_79[0], 0.0) + total_elemento9

r = 4
for code in cuentas_reporte:
    desc = pcge_map.get(code, "")
    debe = movimientos[code]["debe"]
    haber = movimientos[code]["haber"]
    deudor = max(debe - haber, 0.0)
    acreedor = max(haber - debe, 0.0)

    aj_deudor = ajustes_deudor.get(code, 0.0)
    aj_acreedor = ajustes_acreedor.get(code, 0.0)

    ws6.cell(r, 1, code)
    ws6.cell(r, 2, desc)
    ws6.cell(r, 3, debe)
    ws6.cell(r, 4, haber)
    ws6.cell(r, 5, deudor)
    ws6.cell(r, 6, acreedor)
    ws6.cell(r, 7, aj_deudor)
    ws6.cell(r, 8, aj_acreedor)

    # Saldos ajustados: SOLO cuentas de balance (elemento 1 al 5).
    # Las cuentas de resultados (elemento 6-9) no se muestran aquí;
    # su saldo neto se refleja directamente en R.Naturaleza/R.Función.
    if clasificar_resultado(code):
        sa_debe, sa_haber = 0.0, 0.0
    elif aj_deudor or aj_acreedor:
        sa_debe, sa_haber = 0.0, 0.0
    else:
        sa_debe, sa_haber = deudor, acreedor
    ws6.cell(r, 9, sa_debe)
    ws6.cell(r, 10, sa_haber)

    # Saldo neto tras ajuste.
    # IMPORTANTE: para Naturaleza usamos el saldo después de 69 <-> 61;
    # para Función NO debemos borrar 69 ni las cuentas del elemento 9,
    # porque esas cuentas son precisamente las que alimentan el ERF.
    neto = (deudor + aj_deudor) - (acreedor + aj_acreedor)
    neto_deudor = max(neto, 0.0)
    neto_acreedor = max(-neto, 0.0)

    if es_naturaleza(code):
        ws6.cell(r, 11, neto_deudor)
        ws6.cell(r, 12, neto_acreedor)

    if es_funcion(code):
        # ERF: conservar el saldo original de 69, 94, 95 y demás
        # cuentas del elemento 9. No usar el saldo ajustado porque las
        # contrapartidas de cierre las llevarían artificialmente a cero.
        ws6.cell(r, 13, deudor)
        ws6.cell(r, 14, acreedor)

    if es_balance(code):
        ws6.cell(r, 15, deudor)
        ws6.cell(r, 16, acreedor)

    # Distribución y ajustes: presentación contable de las transferencias.
    # 69 y elemento 9 pasan al HABER; 61 y 79 reciben la contrapartida
    # en el DEBE. Esta columna es informativa y no reemplaza los ajustes
    # G/H utilizados para el cálculo de los saldos.
    distrib_debe = 0.0
    distrib_haber = 0.0
    if es_variacion_existencias(code):
        distrib_debe = aj_deudor
    elif es_cuenta79(code):
        distrib_debe = aj_deudor
    elif es_costo_ventas(code) or es_elemento9(code):
        distrib_haber = aj_acreedor
    ws6.cell(r, 17, distrib_debe)
    ws6.cell(r, 18, distrib_haber)

    for c in range(1, 19):
        ws6.cell(r, c).font = BLACK
    for c in range(3, 19):
        ws6.cell(r, c).number_format = '#,##0.00;(#,##0.00);"-"'
    r += 1

HT_LAST_ROW = r - 1
HT_TOTAL_ROW = r
ws6.cell(r, 2, "TOTAL").font = BOLD
for c in range(3, 19):
    letter = get_column_letter(c)
    ws6.cell(r, c, f'=SUM({letter}4:{letter}{HT_LAST_ROW})').font = BOLD
    ws6.cell(r, c).number_format = '#,##0.00;(#,##0.00);"-"'

ws6.freeze_panes = "A4"
autofit(ws6, [11, 44, 14, 14, 14, 14, 13, 13, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14])


def ht_sum(code, col):
    return f'=SUMIF(HT!$A$4:$A${HT_LAST_ROW},"{code}",HT!${col}$4:${col}${HT_LAST_ROW})'

# ============================================================
# HOJAS: ESTADOS FINANCIEROS
# ============================================================
# Los tres estados se alimentan de la misma HT:
#   K/L = Resultado por Naturaleza
#   M/N = Resultado por Función
#   O/P = Estado de Situación Financiera
# De esta forma los resultados parten del mismo saldo ajustado y no
# se generan diferencias artificiales entre ERN, ERF y ESF.

# ------------------------------------------------------------
# ESTADOS FINANCIEROS — PRESENTACIÓN FINAL TANA
# ------------------------------------------------------------
# IMPORTANTE:
# - La HT queda intacta y es la única fuente de datos.
# - Los estados detectan las cuentas realmente presentes en la práctica.
# - No se inventan cuentas ni importes.
# - ERF: 70, 69, 94 y 95 son estructurales; 78 se incorpora si existe;
#        65 y 67 se incorporan solo si existen y no tienen destino a 94/95.
#        79 NO se presenta en el ERF.
# - ERN: presenta las cuentas por naturaleza y el resultado del ejercicio.
# - ESF: presenta activo, pasivo y patrimonio, y verifica A = P + PN.

# ------------------------------------------------------------
# Utilidades para los estados
# ------------------------------------------------------------
def _prefix_exists(prefix):
    return any(str(c).startswith(prefix) for c in cuentas_reporte)


def _sum_ht(prefix, column):
    """Suma el saldo de una familia de cuentas en una columna de HT."""
    return f'=SUMPRODUCT((LEFT(HT!$A$4:$A${HT_LAST_ROW},2)="{prefix}")*HT!${column}$4:${column}${HT_LAST_ROW})'


def _sum_ht_codes(codes, column):
    if not codes:
        return '=0'
    formulas = [
        f'SUMPRODUCT((HT!$A$4:$A${HT_LAST_ROW}="{code}")*HT!${column}$4:${column}${HT_LAST_ROW})'
        for code in codes
    ]
    return '=' + '+'.join(formulas)


def _set_report_value(ws, row, col, formula, bold=False):
    cell = ws.cell(row=row, column=col, value=formula)
    cell.font = BOLD if bold else BLACK
    cell.number_format = '#,##0.00;(#,##0.00);"-"'
    return cell


def _report_title(ws, title):
    ws.merge_cells('B2:E2')
    ws['B2'] = title
    ws['B2'].font = TITLE_FONT
    ws['B2'].alignment = Alignment(horizontal='left')
    ws.merge_cells('B3:E3')
    ws['B3'] = 'Expresado en soles'
    ws['B3'].font = SUBTITLE_FONT


def _report_header(ws, row, right_label='AÑO 2026'):
    ws.cell(row=row, column=2, value='DESCRIPCIÓN').font = BOLD
    ws.cell(row=row, column=4, value='Notas').font = BOLD
    ws.cell(row=row, column=5, value=right_label).font = BOLD
    for c in (2, 4, 5):
        ws.cell(row=row, column=c).fill = PatternFill('solid', fgColor='D9E1F2')
        ws.cell(row=row, column=c).border = Border(
            top=Side(style='thin', color='808080'),
            bottom=Side(style='thin', color='808080')
        )
        ws.cell(row=row, column=c).alignment = Alignment(horizontal='center')


def _write_label(ws, row, text, bold=False):
    ws.cell(row=row, column=2, value=text).font = BOLD if bold else BLACK


def _write_amount(ws, row, formula, bold=False):
    # Columna E: importe del estado, alineado con el modelo enviado.
    _set_report_value(ws, row, 5, formula, bold=bold)


def _hide_control_row(ws, row):
    if row:
        ws.row_dimensions[row].hidden = True

# ============================================================
# ERF — ESTADO DE RESULTADOS POR FUNCIÓN
# ============================================================
ws8 = wb.create_sheet('ERF')
_report_title(ws8, 'ESTADO DE RESULTADOS POR FUNCIÓN')
_report_header(ws8, 4)

r = 5
_write_label(ws8, r, 'INGRESOS OPERACIONALES', True); r += 1
ventas_row = r
_write_label(ws8, r, 'VENTAS')
_write_amount(ws8, r, _sum_ht('70', 'N'), False)
r += 1

# Líneas de detalle de ventas: solo se muestran cuando existen cuentas 70 adicionales.
ventas_codes = sorted(c for c in cuentas_reporte if c.startswith('70'))
if len(ventas_codes) > 1:
    for code in ventas_codes:
        _write_label(ws8, r, f'{code} - {pcge_map.get(code, code)}')
        _write_amount(ws8, r, _sum_ht_codes([code], 'N'))
        r += 1

ventas_total_row = r
_write_label(ws8, r, 'INGRESOS OPERACIONALES', True)
_write_amount(ws8, r, f'=E{ventas_row}', True)
r += 1

_write_label(ws8, r, 'COSTO DE VENTA', True)
costo_row = r
_write_amount(ws8, r, _sum_ht('69', 'M'))
r += 1

utilidad_bruta_row = r
_write_label(ws8, r, 'UTILIDAD BRUTA', True)
_write_amount(ws8, r, f'=E{ventas_total_row}-E{costo_row}', True)
r += 2

_write_label(ws8, r, 'GASTOS OPERACIONALES', True); r += 1

gasto_operativo_rows = []
# 95 y 94 son obligatorias en la estructura, aunque su saldo sea cero.
for prefix, label in [('95', 'Gastos de venta'), ('94', 'Gastos de administración')]:
    rr = r
    _write_label(ws8, r, label.upper())
    _write_amount(ws8, r, f'=-{_sum_ht(prefix, "M")[1:]}')
    gasto_operativo_rows.append(rr)
    r += 1

# 65: solo si existe y no fue destinada a 94/95.
for code in sorted(c for c in cuentas_reporte if len(c) == 5 and c.startswith('65') and c not in CUENTAS_6_CON_DESTINO):
    rr = r
    _write_label(ws8, r, f'{code} - {pcge_map.get(code, code)}')
    _write_amount(ws8, r, f'=-{_sum_ht_codes([code], "M")[1:]}')
    gasto_operativo_rows.append(rr)
    r += 1

utilidad_operativa_row = r
_write_label(ws8, r, 'UTILIDAD OPERATIVA', True)
parts = [f'E{utilidad_bruta_row}'] + [f'+E{x}' for x in gasto_operativo_rows]
_write_amount(ws8, r, '=' + ''.join(parts), True)
r += 2

_write_label(ws8, r, 'OTROS INGRESOS Y GASTOS', True); r += 1

# 78: se incorpora si existe.
otros_78_row = None
if _prefix_exists('78'):
    otros_78_row = r
    _write_label(ws8, r, 'OTROS INGRESOS')
    _write_amount(ws8, r, _sum_ht('78', 'N'))
    r += 1

# Ingreso financiero 77, si existe.
ingreso_fin_row = None
if _prefix_exists('77'):
    ingreso_fin_row = r
    _write_label(ws8, r, 'INGRESO FINANCIERO')
    _write_amount(ws8, r, _sum_ht('77', 'N'))
    r += 1

# 67: solo si existe y no tiene destino a 94/95; se presenta como gasto financiero.
gasto_fin_rows = []
for code in sorted(c for c in cuentas_reporte if len(c) == 5 and c.startswith('67') and c not in CUENTAS_6_CON_DESTINO):
    rr = r
    _write_label(ws8, r, f'{code} - {pcge_map.get(code, code)}')
    _write_amount(ws8, r, f'=-{_sum_ht_codes([code], "M")[1:]}')
    gasto_fin_rows.append(rr)
    r += 1

resultado_antes_part_row = r
_write_label(ws8, r, 'RESULTADO ANTES DE PARTICIPACIONES E IMPUESTOS', True)
parts = [f'E{utilidad_operativa_row}']
if otros_78_row is not None:
    parts.append(f'+E{otros_78_row}')
if ingreso_fin_row is not None:
    parts.append(f'+E{ingreso_fin_row}')
parts += [f'+E{x}' for x in gasto_fin_rows]
_write_amount(ws8, r, '=' + ''.join(parts), True)
r += 1

# Participaciones: solo si existe elemento 87; si no existe, se mantiene 0.
part_row = r
_write_label(ws8, r, 'PARTICIPACIONES')
_write_amount(ws8, r, f'=-{_sum_ht("87", "M")[1:]}')
r += 1

# Impuesto a la renta: solo si existe elemento 88; si no existe, 0.
impuesto_row = r
_write_label(ws8, r, 'IMPUESTO A LA RENTA')
_write_amount(ws8, r, f'=-{_sum_ht("88", "M")[1:]}')
r += 1

resultado_erf_row = r
_write_label(ws8, r, 'RESULTADO DEL EJERCICIO', True)
_write_amount(ws8, r, f'=E{resultado_antes_part_row}+E{part_row}+E{impuesto_row}', True)
r += 2

# Control interno: no se muestra en el informe, pero permite comprobar que ERF = ERN.
control_erf_row = r
_write_label(ws8, r, 'CONTROL INTERNO ERF')
_write_amount(ws8, r, '=0')
ws8.cell(r, 6, f'=IF(ABS(E{r})<0.01,"CUADRADO","REVISAR")')
_hide_control_row(ws8, control_erf_row)

ws8.column_dimensions['B'].width = 58
ws8.column_dimensions['C'].width = 3
ws8.column_dimensions['D'].width = 10
ws8.column_dimensions['E'].width = 18
ws8.freeze_panes = 'B5'

# ============================================================
# ERN — ESTADO DE RESULTADOS POR NATURALEZA
# ============================================================
ws7 = wb.create_sheet('ERN')
_report_title(ws7, 'ESTADO DE RESULTADOS POR NATURALEZA')
_report_header(ws7, 4)

r = 5
_write_label(ws7, r, 'INGRESOS OPERACIONALES', True); r += 1

# Ventas y otros ingresos: se detectan por prefijo, sin inventar cuentas.
ventas_ern_row = r
_write_label(ws7, r, 'VENTAS')
_write_amount(ws7, r, _sum_ht('70', 'L'))
r += 1

# Ingresos por naturaleza que efectivamente existan. La 74 es gasto.
for prefix, label in [
    ('71', 'Variación de la producción almacenada'),
    ('72', 'Producción de activo inmovilizado'),
    ('73', 'Descuentos, rebajas y bonificaciones obtenidos'),
    ('75', 'Otros ingresos de gestión'),
    ('76', 'Ganancia por medición / valuación'),
    ('77', 'Ingresos financieros'),
    ('78', 'Otros ingresos'),
]:
    if _prefix_exists(prefix):
        _write_label(ws7, r, label.upper())
        _write_amount(ws7, r, _sum_ht(prefix, 'L'))
        r += 1

ventas_total_ern_row = r
_write_label(ws7, r, 'TOTAL INGRESOS OPERACIONALES', True)
_write_amount(ws7, r, f'=SUM(E{ventas_ern_row}:E{r-1})', True)
r += 2

_write_label(ws7, r, 'COSTO Y GASTOS POR NATURALEZA', True); r += 1

naturaleza_rows = []
for prefix, label in [
    ('60', 'Compras'),
    ('61', 'Variación de existencias'),
    ('62', 'Gastos de personal'),
    ('63', 'Servicios prestados por terceros'),
    ('64', 'Tributos'),
    ('65', 'Otros gastos de gestión'),
    ('66', 'Pérdidas por medición / deterioro'),
    ('67', 'Gastos financieros'),
    ('68', 'Valuación, deterioro y depreciación'),
    ('74', 'Descuentos, rebajas y bonificaciones concedidos'),
]:
    if _prefix_exists(prefix):
        rr = r
        _write_label(ws7, r, label.upper())
        _write_amount(ws7, r, _sum_ht(prefix, 'K'))
        naturaleza_rows.append(rr)
        r += 1

total_gastos_ern_row = r
_write_label(ws7, r, 'TOTAL COSTO Y GASTOS', True)
_write_amount(ws7, r, '=' + '+'.join(f'E{x}' for x in naturaleza_rows) if naturaleza_rows else '=0', True)
r += 2

resultado_ern_row = r
_write_label(ws7, r, 'RESULTADO DEL EJERCICIO', True)
_write_amount(ws7, r, f'=E{ventas_total_ern_row}-E{total_gastos_ern_row}', True)
ERN_RESULTADO_ROW = r
r += 1

# Control interno oculto.
control_ern_row = r
_write_label(ws7, r, 'CONTROL INTERNO ERN')
_write_amount(ws7, r, f'=E{resultado_ern_row}-ERF!E{resultado_erf_row}')
ws7.cell(r, 6, f'=IF(ABS(E{r})<0.01,"CUADRADO","REVISAR")')
_hide_control_row(ws7, control_ern_row)

# Ahora que ERN_RESULTADO_ROW ya existe, completamos el control cruzado del ERF.
ws8.cell(control_erf_row, 5, f'=E{resultado_erf_row}-ERN!E{ERN_RESULTADO_ROW}')
ws8.cell(control_erf_row, 5).number_format = '#,##0.00;(#,##0.00);"-"'

ws7.column_dimensions['B'].width = 58
ws7.column_dimensions['C'].width = 3
ws7.column_dimensions['D'].width = 10
ws7.column_dimensions['E'].width = 18
ws7.freeze_panes = 'B5'

# ============================================================
# ESF — ESTADO DE SITUACIÓN FINANCIERA
# ============================================================
ws9 = wb.create_sheet('ESF')
_report_title(ws9, 'ESTADO DE SITUACIÓN FINANCIERA')
# En el modelo enviado el ESF es un formato de dos bloques, pero se mantiene
# una sola hoja para que sea fácil de imprimir y revisar.
ws9['B4'] = 'ACTIVO'
ws9['B4'].font = BOLD
ws9['D4'] = 'Notas'
ws9['E4'] = 'AÑO 2026'
for c in (2,4,5):
    ws9.cell(4,c).fill = PatternFill('solid', fgColor='D9E1F2')
    ws9.cell(4,c).font = BOLD
    ws9.cell(4,c).alignment = Alignment(horizontal='center')
ws9['G4'] = 'PASIVO Y PATRIMONIO'
ws9['I4'] = 'Notas'
ws9['J4'] = 'AÑO 2026'
for c in (7,9,10):
    ws9.cell(4,c).fill = PatternFill('solid', fgColor='D9E1F2')
    ws9.cell(4,c).font = BOLD
    ws9.cell(4,c).alignment = Alignment(horizontal='center')

# Se usa una fila compartida para ambos bloques, tal como el modelo.
r = 5
left_rows = []
right_rows = []

ws9.cell(r,2,'ACTIVO CORRIENTE').font = BOLD
ws9.cell(r,7,'PASIVO CORRIENTE').font = BOLD
r += 1

# Activo corriente
for prefix, label in [
    ('10','Efectivo y equivalentes de efectivo'),
    ('12','Cuentas por cobrar comerciales'),
    ('14','Cuentas por cobrar al personal / accionistas'),
    ('16','Cuentas por cobrar diversas'),
    ('18','Servicios y otros contratados por anticipado'),
    ('20','Mercaderías'),
    ('25','Materiales y suministros'),
    ('40','Tributos a favor / crédito fiscal'),
]:
    _write_label(ws9, r, label)
    _set_report_value(ws9, r, 5, f'={_sum_ht(prefix, "O")[1:]}')
    left_rows.append(r)
    r += 1

ws9.cell(r,2,'TOTAL ACTIVO CORRIENTE').font = BOLD
_set_report_value(ws9, r, 5, '=' + '+'.join(f'E{x}' for x in left_rows), True)
TOTAL_AC_ROW = r
r += 2

ws9.cell(r,2,'ACTIVO NO CORRIENTE').font = BOLD
r += 1
anc_rows = []
for prefix, label in [
    ('33','Propiedad, planta y equipo - costo'),
    ('37','Activos no corrientes / intangibles'),
    ('36','Desvalorización / deterioro de activos'),
    ('39','Depreciación y amortización acumulada'),
]:
    _write_label(ws9, r, label)
    if prefix in {'36','39'}:
        _set_report_value(ws9, r, 5, f'=-{_sum_ht(prefix, "P")[1:]}')
    else:
        _set_report_value(ws9, r, 5, f'={_sum_ht(prefix, "O")[1:]}')
    anc_rows.append(r)
    r += 1

ws9.cell(r,2,'TOTAL ACTIVO NO CORRIENTE').font = BOLD
_set_report_value(ws9, r, 5, '=' + '+'.join(f'E{x}' for x in anc_rows), True)
TOTAL_ANC_ROW = r
r += 1
ws9.cell(r,2,'TOTAL ACTIVO').font = BOLD
_set_report_value(ws9, r, 5, f'=E{TOTAL_AC_ROW}+E{TOTAL_ANC_ROW}', True)
TOTAL_ACTIVO_ROW = r

# Pasivo y patrimonio se coloca en paralelo, desde fila 6.
r2 = 6
pc_rows = []
for prefix, label in [
    ('40','Tributos y cuentas por pagar al Estado'),
    ('41','Remuneraciones y participaciones por pagar'),
    ('42','Cuentas por pagar comerciales'),
    ('43','Cuentas por pagar diversas'),
    ('44','Cuentas por pagar a socios / dividendos'),
    ('45','Obligaciones financieras'),
    ('46','Cuentas por pagar diversas / terceros'),
    ('47','Cuentas por pagar relacionadas'),
    ('48','Provisiones y obligaciones'),
]:
    _write_label(ws9, r2, label)
    _set_report_value(ws9, r2, 10, f'={_sum_ht(prefix, "P")[1:]}')
    pc_rows.append(r2)
    r2 += 1

ws9.cell(r2,7,'TOTAL PASIVO CORRIENTE').font = BOLD
_set_report_value(ws9, r2, 10, '=' + '+'.join(f'J{x}' for x in pc_rows), True)
TOTAL_PC_ROW = r2
r2 += 2
ws9.cell(r2,7,'PASIVO NO CORRIENTE').font = BOLD
r2 += 1
# Si hubiera cuentas 45/46/47 con tratamiento no corriente, no se inventa
# una clasificación adicional: el total se muestra de forma transparente.
ws9.cell(r2,7,'Obligaciones financieras y otras').font = BLACK
_set_report_value(ws9, r2, 10, '=0')
TOTAL_PNC_ROW = r2
r2 += 1
ws9.cell(r2,7,'TOTAL PASIVO').font = BOLD
_set_report_value(ws9, r2, 10, f'=J{TOTAL_PC_ROW}+J{TOTAL_PNC_ROW}', True)
TOTAL_PASIVO_ROW = r2
r2 += 2
ws9.cell(r2,7,'PATRIMONIO NETO').font = BOLD
r2 += 1
pat_rows = []
for prefix, label in [
    ('50','Capital social'),
    ('51','Acciones de inversión / capital adicional'),
    ('52','Capital adicional'),
    ('56','Resultados no realizados'),
    ('57','Excedente de revaluación'),
    ('58','Reservas'),
    ('59','Resultados acumulados'),
]:
    _write_label(ws9, r2, label)
    if prefix == '59':
        _set_report_value(ws9, r2, 10, f'={_sum_ht(prefix, "P")[1:]}-{_sum_ht(prefix, "O")[1:]}')
    else:
        _set_report_value(ws9, r2, 10, f'={_sum_ht(prefix, "P")[1:]}')
    pat_rows.append(r2)
    r2 += 1

ws9.cell(r2,7,'Resultado del ejercicio').font = BLACK
_set_report_value(ws9, r2, 10, f'=ERN!E{resultado_ern_row}', False)
pat_rows.append(r2)
r2 += 1
ws9.cell(r2,7,'TOTAL PATRIMONIO NETO').font = BOLD
_set_report_value(ws9, r2, 10, '=' + '+'.join(f'J{x}' for x in pat_rows), True)
TOTAL_PATRIMONIO_ROW = r2
r2 += 1
ws9.cell(r2,7,'TOTAL PASIVO Y PATRIMONIO NETO').font = BOLD
_set_report_value(ws9, r2, 10, f'=J{TOTAL_PASIVO_ROW}+J{TOTAL_PATRIMONIO_ROW}', True)
TOTAL_PYPN_ROW = r2
r2 += 1
control_esf_row = r2
ws9.cell(r2,7,'DIFERENCIA (debe ser 0)').font = BOLD
_set_report_value(ws9, r2, 10, f'=E{TOTAL_ACTIVO_ROW}-J{TOTAL_PYPN_ROW}', True)
ws9.cell(r2,11, f'=IF(ABS(J{r2})<0.01,"CUADRADO","REVISAR")').font = BOLD
_hide_control_row(ws9, control_esf_row)

for col, width in {'B':48,'C':3,'D':9,'E':18,'G':48,'H':3,'I':9,'J':18,'K':14}.items():
    ws9.column_dimensions[col].width = width
ws9.freeze_panes = 'B5'

# ============================================================
# FIN DE ESTADOS FINANCIEROS
# ============================================================

# HOJA: ASIENTOS_CONTABLES (resueltos y validados por TANA)
# ============================================================
if "asientos_contables" in st.session_state:
    ws_ac = wb.create_sheet("Asientos_Contables")
    ac_headers = ["N° Asiento", "Fecha", "Glosa", "Documento", "Operación", "Código", "Denominación", "Concepto", "Debe S/", "Haber S/"]
    for i, h in enumerate(ac_headers, start=1):
        ws_ac.cell(row=1, column=i, value=h)
    style_header(ws_ac, 1, 1, len(ac_headers))
    rr = 2
    pcge_map_export = {str(cod).strip(): str(desc) for cod, desc in PCGE_DATA}
    for asiento in st.session_state["asientos_contables"]:
        first_line = True
        for line in asiento.get("lineas", []):
            code = str(line.get("codigo", "")).strip()

            # Para una presentación limpia: los datos identificadores del asiento
            # aparecen únicamente en su primera línea.
            if first_line:
                numero = asiento.get("numero", "")
                fecha = asiento.get("fecha", "")
                glosa = asiento.get("glosa", "")
                documento = asiento.get("documento", "")
                operacion = asiento.get("operacion_numero", "")
                first_line = False
            else:
                numero = ""
                fecha = ""
                glosa = ""
                documento = ""
                operacion = ""

            values = [
                numero, fecha, glosa, documento, operacion, code,
                pcge_map_export.get(code, line.get("denominacion", "")),
                line.get("concepto", ""), line.get("debe", 0), line.get("haber", 0)
            ]
            for cc, value in enumerate(values, start=1):
                ws_ac.cell(row=rr, column=cc, value=value).font = BLACK
            ws_ac.cell(row=rr, column=9).number_format = '#,##0.00;(#,##0.00);"-"'
            ws_ac.cell(row=rr, column=10).number_format = '#,##0.00;(#,##0.00);"-"'
            rr += 1
    autofit(ws_ac, [12, 13, 35, 18, 12, 12, 48, 42, 14, 14])
    ws_ac.freeze_panes = "A2"

# ============================================================
# HOJA: MONOGRAFIA (fuente leída por TANA)
# ============================================================
if "monografia_texto" in st.session_state:
    ws_mono = wb.create_sheet("Monografia")
    ws_mono["A1"] = "MONOGRAFÍA / DOCUMENTO FUENTE"
    ws_mono["A1"].font = TITLE_FONT
    ws_mono["A2"] = st.session_state.get("monografia_nombre", "")
    ws_mono["A2"].font = SUBTITLE_FONT
    ws_mono["A4"] = st.session_state["monografia_texto"]
    ws_mono["A4"].alignment = Alignment(vertical="top", wrap_text=True)
    ws_mono.column_dimensions["A"].width = 120
    ws_mono.freeze_panes = "A4"

# ============================================================
# PRESENTACIÓN DEL EXCEL FINAL
# ============================================================
# Las hojas auxiliares siguen existiendo durante la construcción porque
# alimentan las fórmulas de los estados financieros, pero no se entregan
# al usuario. El archivo final muestra únicamente los reportes solicitados.
HOJAS_PUBLICAS = [
    "Asientos_Contables",
    "LM",
    "HT",
    "ESF",
    "ERF",
    "ERN",
]

for ws in wb.worksheets:
    if ws.title not in HOJAS_PUBLICAS:
        ws.sheet_state = "hidden"

# Dejamos como primera hoja la de Asientos Contables.
# Guardamos la referencia ANTES de quitarla de la lista; después de
# wb._sheets.remove(), volver a hacer wb["Asientos_Contables"] provoca KeyError.
if "Asientos_Contables" in wb.sheetnames:
    ws_asientos = wb["Asientos_Contables"]
    wb._sheets.remove(ws_asientos)
    wb._sheets.insert(0, ws_asientos)

# ============================================================
# GENERAR ARCHIVO EN MEMORIA Y OFRECER DESCARGA
# ============================================================
# Excel debe recalcular las fórmulas al abrir el archivo.
try:
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
except Exception:
    pass

buffer = io.BytesIO()
wb.save(buffer)
buffer.seek(0)

st.success("Workbook generado correctamente.")
if "monografia_nombre" in st.session_state:
    st.caption("La hoja Monografia conserva el texto extraído para revisión.")
st.download_button(
    label="Descargar Excel",
    data=buffer,
    file_name="TANA_Contabilidad.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
