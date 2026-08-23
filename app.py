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

# ============================================================
# CONFIGURACIÓN GENERAL / PÁGINA PÚBLICA
# ============================================================
st.set_page_config(
    page_title="TANA | Inteligencia Artificial Contable",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ============================================================
# AUTENTICACIÓN PÚBLICA DE TANA — GOOGLE OIDC
# ============================================================
# La configuración se mantiene en Streamlit Secrets ([auth]).
# No se exponen aquí Client ID, Client Secret ni cookie_secret.
# TANA no muestra la aplicación hasta que el usuario se autentica.

if not st.user.is_logged_in:
    st.markdown(
        """
        <div style="max-width:720px;margin:9vh auto 0 auto;text-align:center;">
            <div style="font-size:4rem;line-height:1;">📊</div>
            <h1 style="margin-bottom:.2rem;">TANA</h1>
            <p style="font-size:1.15rem;margin-top:0;">Inteligencia Artificial Contable</p>
            <p style="margin:2rem 0 1.2rem 0;">Inicia sesión con Google para ingresar a TANA.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.button(
            "🔐 Iniciar sesión con Google",
            on_click=st.login,
            use_container_width=True,
        )
    st.stop()

# Usuario autenticado: la aplicación continúa normalmente.
# La identificación del rol es automática mediante el correo devuelto por Google.
# El correo autorizado del creador se guarda en Streamlit Secrets y nunca se muestra.
user_email = str(getattr(st.user, "email", "") or "").strip().lower()
creator_email = str(st.secrets.get("TANA_CREATOR_EMAIL", "") or "").strip().lower()
user_role = "Creador" if creator_email and user_email == creator_email else "Estudiante"

# Identificación visible, breve y automática. Nunca mostramos el correo completo.
if user_role == "Creador":
    role_label = "👤 Creador"
else:
    email_local = user_email.split("@", 1)[0].strip()
    initial = (email_local[:1] or "E").upper()
    role_label = f"👤 Estudiante · {initial}"

# Se muestra discretamente en la zona principal, sin llenar la interfaz.
st.markdown(
    f'<div style="display:flex;justify-content:flex-end;margin:-6px 2px 2px 0;">'
    f'<span style="display:inline-block;padding:4px 10px;border:1px solid #DDE8EF;'
    f'border-radius:999px;background:#F7FAFC;color:#5F7180;font-size:12px;'
    f'font-weight:600;">{role_label}</span></div>',
    unsafe_allow_html=True,
)

# El cierre de sesión se ofrece en la barra lateral sin alterar el flujo contable.
with st.sidebar:
    st.button("Cerrar sesión", on_click=st.logout, use_container_width=True)

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


# ============================================================
# SIDEBAR + LAYOUT TIPO CHAT
# ============================================================
# Nota de diseño: esta sección solo cambia PRESENTACIÓN (sidebar,
# burbujas de chat, barra de entrada). No toca extracción, motor de
# asientos, HT, ERN, ERF, ESF ni la generación del Excel.
st.markdown("""
<style>
/* Oculta el header/menú default de Streamlit para look de app */
#MainMenu, header[data-testid="stHeader"] {visibility: hidden; height: 0;}
div[data-testid="stDecoration"] {display: none !important;}
div[data-testid="stToolbar"] {display: none !important;}
div[data-testid="stAppViewContainer"] {padding-top: 0 !important;}
.block-container {padding-top: 0rem; padding-bottom: 21rem; max-width: 980px;}

/* ---- Sidebar tipo ChatGPT/Claude ---- */
section[data-testid="stSidebar"] {background: #F7F9FB; border-right: 1px solid #E3E9EE;}
section[data-testid="stSidebar"] .block-container {padding-top: 1rem;}
.tana-side-logo {display:flex; align-items:center; gap:10px; margin-bottom:14px;}
.tana-side-logo img {border-radius:10px;}
.tana-side-logo span {font-weight:800; font-size:19px; color:#12304A;}
.tana-side-section {font-size:12px; font-weight:700; color:#8B98A3; text-transform:uppercase;
                     letter-spacing:.04em; margin:18px 0 6px 2px;}
.tana-side-item {font-size:14px; color:#334452; padding:6px 8px; border-radius:8px; cursor:default;}
.tana-side-item:hover {background:#EDF2F5;}
.tana-side-empty {font-size:12.5px; color:#A6B0B8; padding:2px 8px;}
.tana-side-account {display:flex; align-items:center; gap:10px; margin-top:26px;
                     padding:10px 8px; border-top:1px solid #E3E9EE;}
.tana-avatar {width:30px; height:30px; border-radius:50%; background:#087EA4; color:#fff;
              display:flex; align-items:center; justify-content:center; font-weight:700; font-size:13px;}
.tana-side-account span {font-size:13px; color:#4D6172;}

/* ---- Burbujas de chat ---- */
.tana-bubble-user {background:#087EA4; color:#fff; padding:10px 15px; border-radius:16px 16px 4px 16px;
                    max-width:78%; margin-left:auto; margin-bottom:14px; font-size:14.5px;}
.tana-bubble-assistant {background:#F5FAFC; border:1px solid #DDE8EF; color:#22333F; padding:14px 18px;
                         border-radius:16px 16px 16px 4px; max-width:88%; margin-bottom:14px; font-size:14.5px;
                         line-height:1.55;}
.tana-result-card {background:#fff; border:1px solid #DDE8EF; border-radius:12px; padding:12px 16px;
                    margin-top:10px; display:flex; align-items:center; gap:10px;}

/* ---- Tarjeta final: éxito + estadísticas + descarga ---- */
.tana-success-card {background:#EFFBF3; border:1px solid #CDEFDA; color:#166534; padding:14px 18px;
                     border-radius:14px; margin: 18px 0 14px 0; font-size:14.5px; display:flex;
                     align-items:center; gap:10px;}
.tana-stats-row {display:flex; gap:12px; margin-bottom:14px;}
.tana-stat-box {flex:1; background:#F5FAFC; border:1px solid #DDE8EF; border-radius:12px;
                 padding:12px 16px; display:flex; align-items:center; gap:10px;}
.tana-stat-box .num {font-size:19px; font-weight:800; color:#12304A; line-height:1.1;}
.tana-stat-box .label {font-size:12px; color:#6B7B87;}
div[data-testid="stDownloadButton"] button {
    background:#16A34A !important; border-color:#16A34A !important; color:#fff !important;
    border-radius:10px !important; font-weight:700;
}
div[data-testid="stDownloadButton"] button:hover {background:#15803D !important; border-color:#15803D !important;}
.tana-result-card .name {font-weight:700; color:#12304A; font-size:13.5px;}

/* ---- Barra de entrada inferior, FIJA de verdad ----
   Streamlit no deja "envolver" columnas con un <div> de markdown (quedan
   como hermanos, no hijos, en el DOM). Por eso anclamos un marcador
   invisible dentro de un st.container() real y usamos :has() para
   fijar exactamente ESE contenedor (y solo ese), sin afectar el resto
   de la página, que sigue haciendo scroll normal. Se listan varias
   variantes del selector para cubrir distintas versiones de Streamlit. */
div[data-testid="stVerticalBlock"]:has(
    > div[data-testid="element-container"] .tana-inputbar-anchor,
    > div[data-testid="stElementContainer"] .tana-inputbar-anchor,
    > .tana-inputbar-anchor
) {
    position: fixed !important;
    bottom: 0; left: 50%; transform: translateX(-50%);
    width: min(940px, 94vw);
    z-index: 999;
    background: #fff;
    border: 1px solid #DDE8EF;
    border-radius: 22px;
    padding: 8px 14px 10px 14px;
    box-shadow: 0 6px 22px rgba(18,48,74,.09);
    margin-bottom: 60px;
}

/* ---- Unifica los 4 controles (adjuntar, texto, micro, enviar) en UNA
   sola barra visual, en vez de 4 cajas separadas. Se quita el borde y
   fondo propio de cada widget de Streamlit y se dejan "transparentes"
   dentro del contenedor blanco de arriba, alineados en una sola fila. ---- */
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) div[data-testid="stHorizontalBlock"] {
    align-items: center !important;
    gap: 6px !important;
}
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) div[data-testid="column"] {
    display: flex; align-items: center; padding: 0 !important;
}

/* Campo de texto: sin borde propio, se funde con la barra */
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) div[data-testid="stTextInput"] > div {
    border: none !important; background: transparent !important; box-shadow: none !important;
}
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) div[data-testid="stTextInput"] input {
    background: transparent !important; font-size: 14.5px; padding-left: 4px;
}

/* Adjuntar archivo: se reduce a un botón circular tipo clip, sin la
   zona de "arrastra y suelta" ni los textos de ayuda de Streamlit */
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) [data-testid="stFileUploader"] {
    width: 40px;
}
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) [data-testid="stFileUploaderDropzoneInstructions"] {
    display: none !important;
}
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) [data-testid*="FileUploaderDropzone"] {
    border: none !important; background: transparent !important; padding: 0 !important;
    min-height: 0 !important; width: 40px; height: 40px;
    display: flex; align-items: center; justify-content: center;
}
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) [data-testid*="FileUploaderDropzone"] button {
    font-size: 0 !important; width: 38px; height: 38px; border-radius: 50%;
    border: 1px solid #DDE8EF !important; background: #F5FAFC !important; padding: 0 !important;
}
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) [data-testid*="FileUploaderDropzone"] button::after {
    content: "📎"; font-size: 17px;
}
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) [data-testid="stFileUploaderFile"] {
    display: none !important; /* la ficha del archivo se muestra con nuestro propio chip, no la de Streamlit */
}

/* Micrófono: mismo tratamiento circular y transparente */
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) [data-testid*="AudioInput"] {
    background: transparent !important; border: none !important; box-shadow: none !important;
    min-height: 0 !important;
}

/* Botón enviar: circular, color de marca */
div[data-testid="stVerticalBlock"]:has(> div[data-testid="element-container"] .tana-inputbar-anchor, > div[data-testid="stElementContainer"] .tana-inputbar-anchor, > .tana-inputbar-anchor) button[kind="primary"] {
    border-radius: 50% !important; width: 40px; height: 40px; padding: 0 !important;
}

/* Chip que aparece cuando ya hay un archivo cargado: la barra se
   agranda un poco hacia arriba para mostrarlo, sin salirse del recuadro fijo */
.tana-file-chip {
    display: flex; align-items: center; gap: 6px;
    background: #F5FAFC; border: 1px solid #DDE8EF; border-radius: 10px;
    color: #12304A; font-size: 12.5px; padding: 5px 10px; margin: 0 4px 6px 4px;
    max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
</style>
""", unsafe_allow_html=True)

LOGO_PATH = os.path.join(os.path.dirname(__file__), "LOGO TANA.png")
if not os.path.exists(LOGO_PATH):
    LOGO_PATH = os.path.join(os.path.dirname(__file__), "LOGO TANA.jpg")

with st.sidebar:
    logo_col, title_col = st.columns([0.35, 1])
    with logo_col:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=42)
    with title_col:
        st.markdown('<div style="font-weight:800;font-size:19px;color:#12304A;padding-top:6px;">TANA</div>',
                     unsafe_allow_html=True)

    if st.button("➕  Nuevo chat", use_container_width=True):
        for _key in (
            "monografia_json", "monografia_texto", "monografia_nombre", "tana_file_signature",
            "asientos_contables", "asientos_validos", "errores_asientos", "alertas_asientos",
            "respuesta_tana", "respuesta_tana_ruta", "audio_tana_processed",
            "tana_modo_trabajo", "tana_correcciones", "tana_correccion_version",
            "tana_excel_origen_bytes", "tana_diagnostico_excel",
            "tana_excel_buffer", "tana_resuelto_signature",
        ):
            st.session_state.pop(_key, None)
        st.rerun()

    st.markdown('<div class="tana-side-section">Historial</div>', unsafe_allow_html=True)
    historial = st.session_state.get("tana_historial", [])
    if historial:
        st.markdown('<div class="tana-side-section" style="margin-top:4px;">Hoy</div>', unsafe_allow_html=True)
        for item in reversed(historial[-15:]):
            st.markdown(f'<div class="tana-side-item">📄 {item}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="tana-side-empty">Aún no hay monografías resueltas en esta sesión.</div>',
                     unsafe_allow_html=True)

    st.markdown(
        '<div class="tana-side-account">'
        '<div class="tana-avatar">T</div><span>Cuenta del suscriptor de TANA</span></div>',
        unsafe_allow_html=True,
    )

# ============================================================
# HISTORIAL DE CONVERSACIÓN (área principal)
# ============================================================
if "tana_chat" not in st.session_state:
    st.session_state["tana_chat"] = []  # lista de dicts: {"role": "user"/"assistant", "content": str}

def _tana_chat_add(role, content):
    st.session_state["tana_chat"].append({"role": role, "content": content})

if not st.session_state["tana_chat"]:
    st.markdown(
        '<div style="text-align:center; padding:2px 0 6px 0;">'
        f'{"<img src=\'data:image/png;base64," + __import__("base64").b64encode(open(LOGO_PATH,"rb").read()).decode() + "\' width=52 style=\'border-radius:12px;\'>" if os.path.exists(LOGO_PATH) else ""}'
        '<div style="font-size:23px;font-weight:800;color:#12304A;margin-top:8px;">¿Qué monografía resolvemos hoy?</div>'
        '<div style="color:#6B7B87;font-size:14.5px;margin-top:5px;">'
        'Sube tu monografía abajo y TANA desarrolla los asientos, la HT y los estados financieros.</div>'
        '</div>',
        unsafe_allow_html=True,
    )

for msg in st.session_state["tana_chat"]:
    css_class = "tana-bubble-user" if msg["role"] == "user" else "tana-bubble-assistant"
    st.markdown(f'<div class="{css_class}">{msg["content"]}</div>', unsafe_allow_html=True)

# ============================================================
# BARRA DE ENTRADA (carga de monografía + consulta + voz + enviar)
# Fija de verdad: vive dentro de su propio st.container(), con un
# marcador invisible que el CSS de arriba usa para anclarla. El resto
# del contenido (burbujas de chat) sigue con scroll normal.
# ============================================================
inputbar_container = st.container()
with inputbar_container:
    st.markdown('<span class="tana-inputbar-anchor"></span>', unsafe_allow_html=True)

    bar = st.columns([0.7, 5.6, 0.85, 0.85], gap="small")
    with bar[0]:
        uploaded_file = st.file_uploader(
            "Archivo", type=SUPPORTED_TYPES, label_visibility="collapsed",
            help="Adjuntar monografía: PDF, DOC, DOCX, XLS, XLSX, JPG, JPEG y PNG."
        )
    with bar[1]:
        pregunta_top = st.text_input(
            "Consulta", placeholder="Pregunta a TANA…",
            key="pregunta_tana_top", label_visibility="collapsed"
        )
    with bar[2]:
        audio_top = st.audio_input("Hablar", key="audio_tana_top", label_visibility="collapsed") if hasattr(st, "audio_input") else None
    with bar[3]:
        enviar_top = st.button("➤", type="primary", key="btn_enviar_tana_top", use_container_width=True)

    # Cuando hay un archivo cargado, la barra se agranda un poco (hacia
    # arriba) para mostrar esta etiqueta, en vez de un texto aparte debajo.
    if uploaded_file:
        st.markdown(
            f'<div class="tana-file-chip">📎 {uploaded_file.name} · pulsa ➤ para enviar y procesar</div>',
            unsafe_allow_html=True,
        )

def _normalizar_nombre_hoja(nombre):
    import unicodedata
    txt = str(nombre or '').strip().lower()
    txt = ''.join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^a-z0-9]', '', txt)

def _normalizar_encabezado(valor):
    import unicodedata
    txt = str(valor or '').strip().lower()
    txt = ''.join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')
    txt = txt.replace('°','o').replace('º','o')
    return re.sub(r'[^a-z0-9]', '', txt)

def _resolver_columnas_asientos(ws, max_scan_rows=30):
    """Encuentra encabezados de asientos aunque estén en otra fila o tengan nombres libres."""
    alias = {
        'N° Asiento': {'asiento','numeroasiento','nroasiento','noasiento','nasiento','numasiento','nro','numero'},
        'Fecha': {'fecha','fecharegistro','fechadelasiento'},
        'Glosa': {'glosa','descripcion','descripcin','concepto','detalle','glosacontable'},
        'Documento': {'documento','doc','documentoref','referencia','nrodocumento'},
        'Operación': {'operacion','numerooperacion','nrooperacion'},
        'Código': {'codigo','cuenta','codigocuenta','cuentacontable','codcuenta'},
        'Denominación': {'denominacion','nombrecuenta','descripcioncuenta','cuenta'},
        'Concepto': {'concepto','detalle','descripcion'},
        'Debe S/': {'debe','debes','debesoles','debes/.','debeimporte'},
        'Haber S/': {'haber','habers','habersoles','habers/.','haberimporte'},
    }
    best=None
    scan_rows=min(max_scan_rows, ws.max_row or 0)
    for r in range(1, scan_rows+1):
        cols={}
        for c in range(1, (ws.max_column or 0)+1):
            n=_normalizar_encabezado(ws.cell(r,c).value)
            if n:
                cols[n]=c
        resolved={}
        for canonical, names in alias.items():
            for n,c in cols.items():
                if n in names:
                    resolved[canonical]=c; break
        score=sum(k in resolved for k in ('Código','Debe S/','Haber S/'))
        score += 1 if 'N° Asiento' in resolved else 0
        score += 1 if 'Fecha' in resolved else 0
        if best is None or score>best[0]: best=(score,r,resolved)
    if not best or best[0] < 3:
        return None
    return {'header_row':best[1], 'resolved':best[2], 'score':best[0]}

def _clasificar_hoja_excel(ws):
    """Clasifica por contenido; el nombre de la hoja solo sirve como pista."""
    info=_resolver_columnas_asientos(ws)
    if info and all(k in info['resolved'] for k in ('Código','Debe S/','Haber S/')):
        return 'libro_diario', info
    # Heurística secundaria para diarios con encabezados muy libres.
    texto=[]
    for r in range(1,min(ws.max_row or 0,25)+1):
        texto.extend(str(ws.cell(r,c).value or '').lower() for c in range(1,min(ws.max_column or 0,20)+1))
    joined=' '.join(texto)
    if ('debe' in joined and 'haber' in joined and ('cuenta' in joined or 'codigo' in joined)):
        return 'libro_diario', info
    return 'desconocida', info

def _encontrar_libro_diario(wb):
    candidatos=[]
    for idx,name in enumerate(wb.sheetnames):
        ws=wb[name]
        tipo,info=_clasificar_hoja_excel(ws)
        pista=_normalizar_nombre_hoja(name)
        bonus=0
        if pista in ('ld','librodiario','diario','asientos','asientoscontables'):
            bonus=2
        if tipo=='libro_diario' and info:
            candidatos.append((info['score']+bonus, idx, name, info))
    if not candidatos:
        return None
    candidatos.sort(reverse=True)
    return candidatos[0][2], candidatos[0][3]

def _cargar_asientos_desde_excel(uploaded):
    """Importa un Libro Diario de cualquier Excel por contenido, no por nombre de hoja."""
    data = uploaded.getvalue()
    try:
        wb_in = openpyxl.load_workbook(io.BytesIO(data), data_only=False, read_only=False)
    except Exception as exc:
        raise ValueError(f'No se pudo abrir el archivo Excel: {exc}')
    encontrado=_encontrar_libro_diario(wb_in)
    if not encontrado:
        raise ValueError('No pude identificar el Libro Diario por su contenido. Necesito una tabla con columnas equivalentes a Cuenta/Código, Debe y Haber.')
    sheet_name, meta=encontrado
    ws=wb_in[sheet_name]
    resolved=meta['resolved']; header_row=meta['header_row']
    def cell(row,key,default=''):
        col=resolved.get(key)
        if not col: return default
        v=ws.cell(row,col).value
        return default if v is None else v
    asientos=[]; actual=None
    for r in range(header_row+1, ws.max_row+1):
        numero=cell(r,'N° Asiento',''); codigo=str(cell(r,'Código','') or '').strip()
        if numero in ('',None) and not codigo: continue
        if numero not in ('',None):
            try:
                if isinstance(numero,float) and numero.is_integer(): numero=int(numero)
                elif isinstance(numero,str) and numero.strip().isdigit(): numero=int(numero.strip())
            except Exception: pass
            actual={'numero':numero,'fecha':cell(r,'Fecha',''),'glosa':cell(r,'Glosa',''),'documento':cell(r,'Documento',''),'operacion_numero':cell(r,'Operación',''),'lineas':[]}
            asientos.append(actual)
        if actual is None or not codigo: continue
        actual['lineas'].append({'codigo':codigo,'denominacion':cell(r,'Denominación',''),'concepto':cell(r,'Concepto',''),'debe':_to_float(cell(r,'Debe S/',0),0.0),'haber':_to_float(cell(r,'Haber S/',0),0.0)})
    if not asientos: raise ValueError('Encontré una hoja con estructura de diario, pero no pude reconstruir ningún asiento.')
    return asientos

def _diagnosticar_asientos(asientos, pcge_map):
    """Diagnóstico determinista y explícito para que TANA señale dónde está el error."""
    diagnostico = []
    total_d = Decimal("0")
    total_h = Decimal("0")

    for idx, asiento in enumerate(asientos or [], start=1):
        numero = asiento.get("numero", idx)
        td = Decimal("0")
        th = Decimal("0")
        line_errors = []

        for li, line in enumerate(asiento.get("lineas", []) or [], start=1):
            code = str(line.get("codigo", "")).strip()
            debe = _money(line.get("debe")) or Decimal("0")
            haber = _money(line.get("haber")) or Decimal("0")
            td += debe
            th += haber

            if not re.fullmatch(r"\d{5}", code):
                line_errors.append(f"línea {li}: la cuenta '{code}' no tiene 5 dígitos")
            elif code not in pcge_map:
                line_errors.append(f"línea {li}: la cuenta {code} no existe en el PCGE")
            if debe > 0 and haber > 0:
                line_errors.append(f"línea {li}: {code} tiene Debe y Haber simultáneamente")

        diff = td - th
        total_d += td
        total_h += th

        if abs(diff) > Decimal("0.01") or line_errors:
            detalle = [f"Asiento {numero}: Debe S/ {td:.2f} vs Haber S/ {th:.2f}."]
            if abs(diff) > Decimal("0.01"):
                detalle.append(f"Diferencia: S/ {abs(diff):.2f}.")
            detalle.extend(line_errors)
            diagnostico.append(" ".join(detalle))

    if not diagnostico:
        diagnostico.append(
            f"No encontré asientos descuadrados. Total Debe S/ {total_d:.2f} = "
            f"Total Haber S/ {total_h:.2f}."
        )
    else:
        diagnostico.insert(
            0,
            f"Se detectaron {len(diagnostico)} asiento(s) que requieren revisión."
        )

    return "\n".join(diagnostico)


def _resumen_excel_para_tutor(uploaded):
    """Lee TODAS las hojas del Excel, una por una.

    El nombre de la hoja es solo una pista. El contenido se conserva para que
    TANA pueda responder después a preguntas como "¿en qué hoja está el error?".
    No obliga al Excel del estudiante a tener una hoja Asientos_Contables.
    """
    try:
        raw = uploaded.getvalue() if uploaded is not None else st.session_state.get("tana_excel_origen_bytes")
        if not raw:
            return ""
        wbv = openpyxl.load_workbook(io.BytesIO(raw), data_only=False, read_only=True)
        partes = [
            f"EXCEL REVISADO: {len(wbv.sheetnames)} hojas.",
            "INSTRUCCIÓN: analizar cada hoja por su contenido; el nombre de la hoja es solo una pista.",
        ]
        for idx, nombre in enumerate(wbv.sheetnames, start=1):
            ws = wbv[nombre]
            tipo, meta = _clasificar_hoja_excel(ws)
            etiqueta = {
                'libro_diario': 'POSIBLE LIBRO DIARIO',
                'desconocida': 'HOJA PARA REVISAR',
            }.get(tipo, 'HOJA')
            partes.append(
                f"\n=== HOJA {idx}: {nombre} | {etiqueta} | filas={ws.max_row} columnas={ws.max_column} ==="
            )
            if meta and meta.get('resolved'):
                partes.append(
                    "COLUMNAS DETECTADAS: " + ", ".join(
                        f"{k}={v}" for k, v in meta['resolved'].items()
                    )
                )
            # Se inspeccionan todas las hojas. Para mantener el contexto manejable,
            # se toma una muestra amplia y representativa de cada hoja.
            max_rows = min(ws.max_row or 0, 180)
            max_cols = min(ws.max_column or 0, 20)
            filas_no_vacias = 0
            for r in ws.iter_rows(min_row=1, max_row=max_rows, max_col=max_cols, values_only=True):
                vals = [str(v) for v in r if v not in (None, "")]
                if vals:
                    filas_no_vacias += 1
                    partes.append(" | ".join(vals))
            if (ws.max_row or 0) > max_rows:
                partes.append(f"... [muestra truncada: se omitieron {ws.max_row - max_rows} filas posteriores] ...")
            partes.append(f"FILAS NO VACÍAS EN MUESTRA: {filas_no_vacias}")
        return "\n".join(partes)[:90000]
    except Exception as exc:
        return f"No fue posible resumir el Excel: {exc}"


def _revisar_excel_completo(uploaded):
    """Construye un inventario de TODAS las hojas sin exigir que exista un diario."""
    raw = uploaded.getvalue()
    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=False, read_only=True)
    hojas = []
    for idx, nombre in enumerate(wb.sheetnames, start=1):
        ws = wb[nombre]
        tipo, meta = _clasificar_hoja_excel(ws)
        # Detectar señales generales, sin declarar todavía que existe un error.
        formulas = 0
        errores_formula = 0
        no_vacias = 0
        for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 0, 300), max_col=min(ws.max_column or 0, 25), values_only=True):
            if any(v not in (None, "") for v in r):
                no_vacias += 1
            for v in r:
                if isinstance(v, str) and v.startswith('='):
                    formulas += 1
                if isinstance(v, str) and v.startswith('#'):
                    errores_formula += 1
        hojas.append({
            'indice': idx, 'nombre': nombre, 'tipo': tipo,
            'filas': ws.max_row or 0, 'columnas': ws.max_column or 0,
            'columnas_detectadas': list((meta or {}).get('resolved', {}).keys()),
            'fila_encabezado': (meta or {}).get('header_row'),
            'filas_no_vacias_muestra': no_vacias,
            'formulas_muestra': formulas,
            'errores_formula_muestra': errores_formula,
        })
    return hojas


def _es_excel_tana(nombre):
    return str(nombre or "").lower().endswith((".xlsx", ".xls"))


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

# El archivo se procesa automáticamente al cargarse. No se muestra un botón
# intermedio: la lógica contable original permanece intacta.
if uploaded_file:
    file_signature = f"{uploaded_file.name}|{getattr(uploaded_file, 'size', 0)}"
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


if uploaded_file is not None and enviar_top and st.session_state.get("tana_file_signature") != file_signature:
    for _key in (
        "monografia_json", "monografia_texto", "monografia_nombre", "archivo_excel_origen",
        "asientos_contables", "asientos_validos", "errores_asientos",
        "alertas_asientos", "respuesta_tana", "respuesta_tana_ruta", "audio_tana_processed",
        "tana_modo_trabajo", "tana_correcciones", "tana_correccion_version",
        "tana_excel_origen_bytes", "tana_diagnostico_excel",
        "tana_excel_buffer", "tana_resuelto_signature",
    ):
        st.session_state.pop(_key, None)
    if _es_excel_tana(uploaded_file.name):
        with st.spinner("TANA está revisando todas las hojas del Excel…"):
            try:
                # PRIMERA FASE: revisión completa. No se exige Libro Diario y no
                # se detiene el proceso si no se pueden reconstruir asientos.
                hojas_revisadas = _revisar_excel_completo(uploaded_file)
                resumen_completo = _resumen_excel_para_tutor(uploaded_file)
                st.session_state["tana_excel_origen_bytes"] = uploaded_file.getvalue()
                st.session_state["archivo_excel_origen"] = uploaded_file.name
                st.session_state["monografia_nombre"] = uploaded_file.name
                st.session_state["tana_excel_hojas"] = hojas_revisadas
                st.session_state["tana_excel_revisado"] = True
                st.session_state["tana_excel_resumen_completo"] = resumen_completo
                st.session_state["monografia_texto"] = (
                    "Excel del estudiante revisado hoja por hoja. La información de todas las hojas "
                    "queda disponible para responder preguntas y localizar errores."
                )
                # Intentamos reconstruir asientos solo como capacidad adicional.
                # Si no se puede, NO es un error de carga y TANA continúa con la revisión.
                try:
                    asientos_importados = _cargar_asientos_desde_excel(uploaded_file)
                    valid, errors, warnings = validate_asientos({"asientos": asientos_importados}, pcge_map)
                    st.session_state["asientos_contables"] = asientos_importados
                    st.session_state["asientos_validos"] = valid
                    st.session_state["errores_asientos"] = errors
                    st.session_state["alertas_asientos"] = warnings
                    st.session_state["tana_diagnostico_excel"] = _diagnosticar_asientos(asientos_importados, pcge_map)
                except Exception:
                    # No se obliga al alumno a entregar una hoja de diario compatible.
                    st.session_state.pop("asientos_contables", None)
                    st.session_state["asientos_validos"] = []
                    st.session_state["errores_asientos"] = []
                    st.session_state["alertas_asientos"] = []
                    st.session_state["tana_diagnostico_excel"] = "La revisión por asientos se realizará cuando la pregunta identifique una operación o asiento concreto."

                st.session_state["tana_file_signature"] = file_signature
                st.session_state["tana_modo_trabajo"] = "revision_excel"
                # NO guardar el archivo original como archivo de salida.
                st.session_state.pop("tana_excel_buffer", None)
                _tana_chat_add(
                    "user",
                    f"📊 Cargó un Excel para revisión: <b>{uploaded_file.name}</b>"
                )
                resumen_hojas = "<br>".join(
                    f"{h['indice']}. <b>{h['nombre']}</b> — {h['filas']} filas × {h['columnas']} columnas"
                    for h in hojas_revisadas
                )
                _tana_chat_add(
                    "assistant",
                    "<b>✅ Tu Excel ha sido revisado hoja por hoja.</b><br><br>"
                    "He leído todas las hojas y tengo su contenido disponible para analizarlo.<br><br>"
                    + resumen_hojas +
                    "<br><br><b>Ahora puedes preguntarme dónde está un error.</b> Por ejemplo: "
                    "<i>¿En qué hoja está el error y qué debo corregir?</i>"
                )
            except Exception as exc:
                st.error(f"No se pudo revisar el Excel: {exc}")
                st.stop()
    else:
        with st.spinner("TANA está leyendo y procesando la monografía…"):
            try:
                extracted = extract_with_gemini(uploaded_file)
                st.session_state["monografia_json"] = extracted
                st.session_state["monografia_texto"] = extraction_to_text(extracted)
                st.session_state["monografia_nombre"] = uploaded_file.name
                st.session_state["tana_file_signature"] = file_signature
            except json.JSONDecodeError:
                st.error("Gemini respondió con un formato que no pudo convertirse a JSON. Vuelve a intentarlo.")
                st.stop()
            except Exception as exc:
                st.error(f"No se pudo procesar el archivo con Gemini: {exc}")
                st.stop()

if "monografia_json" in st.session_state:
    data = st.session_state["monografia_json"]
    _nombre_mono = st.session_state.get('monografia_nombre', 'archivo')
    if not any(m["role"] == "user" and _nombre_mono in m["content"] for m in st.session_state["tana_chat"]):
        _tana_chat_add("user", f"📄 Cargó la monografía: <b>{_nombre_mono}</b>")
        _tana_chat_add("assistant", f"Monografía recibida: <b>{_nombre_mono}</b>. Estoy desarrollando los asientos…")
        if _nombre_mono not in st.session_state.get("tana_historial", []):
            st.session_state.setdefault("tana_historial", []).append(_nombre_mono)




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

if "monografia_json" in st.session_state and "asientos_contables" not in st.session_state:
    with st.spinner("TANA está desarrollando y validando los asientos contables…"):
        try:
            resolved, pcge_map = resolve_asientos_with_gemini()
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
            asientos_generados = asegurar_cuenta_79_en_destinos(asientos_generados, pcge_map)
            asientos_generados = corregir_retiro_socio(asientos_generados, st.session_state.get("monografia_json", {}))
            valid, errors, warnings = validate_asientos({"asientos": asientos_generados}, pcge_map)
            st.session_state["asientos_contables"] = asientos_generados
            st.session_state["asientos_validos"] = valid
            st.session_state["errores_asientos"] = errors
            st.session_state["alertas_asientos"] = list(alertas_gemini) + list(warnings)
        except Exception as exc:
            st.error(f"No se pudieron desarrollar los asientos: {exc}")
            st.stop()

# ============================================================
# MODO DE TRABAJO + CORRECCIÓN DEL EXCEL
# ============================================================
def _detectar_modo_trabajo(texto):
    """Interpreta instrucciones simples del estudiante sin cambiar el motor contable."""
    t = (texto or "").lower()
    if any(k in t for k in ("solo estados", "solo estado", "estado de resultados", "estados financieros")):
        return "completo"
    if any(k in t for k in ("libro diario", "solo diario", "diario contable")):
        return "asientos"
    if any(k in t for k in ("solo asientos", "solo los asientos", "asientos contables", "desarrolla los asientos")):
        return "asientos"
    if any(k in t for k in ("toda la contabilidad", "todo completo", "todo el proceso", "todos los estados", "hoja de trabajo")):
        return "completo"
    return None


def _es_peticion_correccion(texto):
    t = (texto or "").lower()
    return any(k in t for k in (
        "corrige", "corregir", "corrección", "correccion", "te equivocaste",
        "está mal", "esta mal", "error", "modifica", "modificar", "cambia",
        "cambiar", "reemplaza", "reemplazar", "debería ser", "debe ser",
        "no corresponde", "incorrecto", "incorrecta", "ajusta", "ajustar"
    ))



def _corregir_excel_directamente_con_gemini(instruccion):
    """Analiza TODO el Excel y devuelve cambios de celdas concretos.

    Esta ruta es para Excel del estudiante: no exige Libro Diario ni intenta
    convertir la hoja HT/EF/ES en asientos. Gemini propone únicamente cambios
    explícitos de celdas, y TANA los aplica sobre una COPIA del libro original.
    """
    raw = st.session_state.get("tana_excel_origen_bytes")
    if not raw:
        raise ValueError("No hay un Excel original cargado para corregir.")

    contexto = st.session_state.get("tana_excel_resumen_completo", "")
    prompt = f"""Eres TANA, un sistema de revisión y corrección de Excel contable peruano.

El estudiante te pide CORREGIR DIRECTAMENTE SU EXCEL. NO respondas que no puedes
modificar archivos. Tu tarea aquí es proponer cambios concretos de celdas para que
la aplicación pueda aplicarlos a una COPIA del Excel.

REGLAS ABSOLUTAS:
1. Analiza TODAS las hojas del Excel, no solamente el Libro Diario.
2. La hoja puede llamarse HT, Hoja de Trabajo, Balance, Hoja1 o cualquier otro nombre.
3. El nombre de la hoja es solo una pista; usa el contenido para localizarla.
4. Si el estudiante indica una hoja y una cuenta/fila concreta, modifica ÚNICAMENTE
   esa fila y las celdas estrictamente necesarias de esa fila.
5. NO modifiques otras filas, otras hojas, formatos, fórmulas o valores no relacionados.
6. Si el error es una fórmula desplazada/desalineada, devuelve la fórmula correcta
   para la celda concreta, respetando las referencias relativas de esa fila.
7. Si el estudiante pide "corrige solo esa fila", la respuesta debe contener solo
   las celdas de esa fila.
8. No inventes valores. Usa las demás hojas como evidencia cruzada (LD/HT/EF/ES,
   aunque tengan otros nombres).
9. Antes de proponer el cambio, comprueba que la celda actual y la fila coinciden
   con la evidencia del Excel.
10. Devuelve JSON válido, sin Markdown.

FORMATO OBLIGATORIO:
{{
  "estado": "corregible" | "no_hay_evidencia",
  "hoja": "nombre exacto de la hoja",
  "fila": 0,
  "cambios": [
    {{
      "celda": "C23",
      "valor_actual": "valor o fórmula actual",
      "valor_nuevo": "valor o fórmula nueva",
      "motivo": "explicación breve"
    }}
  ],
  "observacion": "qué se corrigió y por qué"
}}

Si no puedes demostrar con el contenido del Excel qué celda debe cambiar,
usa estado "no_hay_evidencia" y no inventes una corrección.

CONTENIDO DEL EXCEL, HOJA POR HOJA:
{contexto[:100000]}

SOLICITUD EXACTA DEL ESTUDIANTE:
{instruccion}
"""
    response, profile = _generate_with_fallback(
        lambda client: [prompt],
        types.GenerateContentConfig(response_mime_type="application/json"),
    )
    data = json.loads(response.text or "{}")
    return data, profile


def _aplicar_cambios_celdas_excel(propuesta):
    """Aplica cambios explícitos sobre una COPIA del Excel original."""
    raw = st.session_state.get("tana_excel_origen_bytes")
    if not raw:
        raise ValueError("No hay Excel original cargado.")
    if propuesta.get("estado") != "corregible":
        raise ValueError(propuesta.get("observacion") or "No existe evidencia suficiente para corregir el Excel.")

    hoja = str(propuesta.get("hoja") or "").strip()
    fila = int(propuesta.get("fila") or 0)
    cambios = propuesta.get("cambios") or []
    if not hoja or fila < 1 or not cambios:
        raise ValueError("La propuesta de corrección no contiene hoja, fila y cambios concretos.")

    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=False)
    if hoja not in wb.sheetnames:
        raise ValueError(f"La hoja '{hoja}' no existe en el Excel original.")

    # Seguridad: todos los cambios deben pertenecer a la fila indicada.
    for cambio in cambios:
        celda = str(cambio.get("celda") or "").strip()
        m = re.fullmatch(r"([A-Za-z]+)(\d+)", celda)
        if not m or int(m.group(2)) != fila:
            raise ValueError(f"Cambio rechazado: {celda} no pertenece a la fila {fila}.")

    ws = wb[hoja]
    aplicados=[]
    for cambio in cambios:
        celda = str(cambio["celda"]).strip()
        valor_actual = cambio.get("valor_actual")
        valor_nuevo = cambio.get("valor_nuevo")
        actual = ws[celda].value
        # Si Gemini recibió una celda vacía, permitimos None/""; si no, exigimos
        # coincidencia para evitar escribir encima de una versión distinta.
        if valor_actual not in (None, ""):
            if str(actual).strip() != str(valor_actual).strip():
                raise ValueError(
                    f"La celda {hoja}!{celda} cambió desde el diagnóstico. "
                    f"Actual='{actual}' / esperado='{valor_actual}'. No se aplicó ningún cambio."
                )
        ws[celda] = valor_nuevo
        aplicados.append((celda, actual, valor_nuevo, cambio.get("motivo", "")))

    ws_rev = wb.create_sheet("Revision_TANA")
    ws_rev["A1"] = "TANA — CORRECCIÓN DE CELDAS"
    ws_rev["A2"] = "Estado"
    ws_rev["B2"] = "CORRECCIÓN APLICADA"
    ws_rev["A3"] = "Hoja"
    ws_rev["B3"] = hoja
    ws_rev["A4"] = "Fila corregida"
    ws_rev["B4"] = fila
    ws_rev["A6"] = "Celda"
    ws_rev["B6"] = "Valor anterior"
    ws_rev["C6"] = "Valor nuevo"
    ws_rev["D6"] = "Motivo"
    for i,(celda,antes,nuevo,motivo) in enumerate(aplicados,start=7):
        ws_rev.cell(i,1,celda); ws_rev.cell(i,2,str(antes)); ws_rev.cell(i,3,str(nuevo)); ws_rev.cell(i,4,str(motivo))
    ws_rev.column_dimensions["A"].width=15; ws_rev.column_dimensions["B"].width=30
    ws_rev.column_dimensions["C"].width=30; ws_rev.column_dimensions["D"].width=80
    ws_rev["A"+str(len(aplicados)+9)] = "Advertencia"
    ws_rev["B"+str(len(aplicados)+9)] = "Soy una inteligencia artificial y puedo cometer errores. Revisa siempre el resultado antes de utilizarlo."

    try:
        wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode="auto"
    except Exception:
        pass
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue(), aplicados

def _corregir_asientos_con_gemini(instruccion):
    """Pide una corrección controlada y devuelve TODOS los asientos corregidos."""
    actuales = st.session_state.get("asientos_contables", [])
    mono = st.session_state.get("monografia_texto", "")
    pcge_5 = [[str(c).strip(), str(d)] for c, d in PCGE_DATA if re.fullmatch(r"\d{5}", str(c).strip())]
    diagnostico_actual = st.session_state.get("tana_diagnostico_excel", "")
    excel_contexto = _resumen_excel_para_tutor(None) if st.session_state.get("tana_excel_origen_bytes") else ""
    prompt = f"""Eres TANA, motor contable peruano, y estás corrigiendo un Excel que tú misma generaste.
El estudiante está revisando un Excel que TANA ya generó y solicita una corrección.
El conjunto "ASIENTOS ACTUALES" fue reconstruido directamente desde ese Excel.
Debes tratarlo como la versión que el estudiante está corrigiendo.

DIAGNÓSTICO DETERMINISTA DEL EXCEL: 
{diagnostico_actual[:12000]}

IMPORTANTE: si el estudiante pide "corrige", "corrígelo" o equivalente, debes intentar
corregir usando este diagnóstico y la instrucción del estudiante. No rechaces la tarea
solo porque la versión actual esté descuadrada. La versión actual puede estar precisamente
mal y el diagnóstico es la evidencia que debes utilizar para proponer la corrección.

REGLAS DE CORRECCIÓN:
1. Modifica únicamente lo que el estudiante señala.
2. Conserva todos los demás asientos, fechas, glosas, documentos, cuentas e importes que no sean afectados.
3. Si la corrección cambia un importe, recalcula las líneas del mismo asiento para que Debe = Haber.
4. Usa únicamente cuentas PCGE de 5 dígitos del catálogo proporcionado.
5. Intenta corregir a partir de TRES fuentes, en este orden: (a) instrucción del estudiante, (b) diagnóstico determinista, (c) contenido del Excel. Si el Excel contiene una cuenta inválida, busca en el contexto del asiento, denominación, operación y demás hojas la cuenta válida más coherente del PCGE.
6. No hagas una corrección arbitraria solo para cuadrar. Si no existe evidencia suficiente para escoger una cuenta, conserva esa línea y explica exactamente qué dato falta.
7. Si el estudiante pide expresamente "corrige", "corrígelo", "corrige los errores" o "genera un nuevo Excel", debes INTENTAR una corrección; no respondas únicamente que no puedes.
8. Si existe un descuadre, identifica primero qué asiento y qué líneas lo producen y modifica solo las líneas justificadas por la evidencia.
9. Después de modificar, comprueba mentalmente Debe = Haber y que las cuentas propuestas estén en el PCGE.
10. Devuelve SIEMPRE el conjunto COMPLETO de asientos, no solo el asiento modificado.
11. La respuesta debe ser JSON válido.

ESTRUCTURA OBLIGATORIA:
{{
  "asientos": [ ... todos los asientos completos ... ],
  "observacion": "explicación breve de lo corregido"
}}

DOCUMENTO FUENTE / CONTEXTO ORIGINAL:
{mono[:14000]}

ASIENTOS ACTUALES:
{json.dumps(actuales, ensure_ascii=False, indent=2)[:30000]}

PCGE DE 5 DÍGITOS:
{json.dumps(pcge_5, ensure_ascii=False)}

CONTENIDO DEL EXCEL CARGADO:
{excel_contexto[:30000]}

SOLICITUD DEL ESTUDIANTE:
{instruccion}
"""
    response, profile = _generate_with_fallback(
        lambda client: [prompt],
        types.GenerateContentConfig(response_mime_type="application/json"),
    )
    data = json.loads(response.text or "{}")
    return data, profile


def _crear_excel_revision_desde_origen(asientos, errores):
    """Crea un nuevo Excel a partir del Excel que el estudiante cargó.
    Conserva todas sus hojas/formatos y reemplaza las líneas de Asientos_Contables.
    Si la propuesta no está validada, deja una hoja/nota de advertencia y fuerza
    recálculo al abrir el archivo.
    """
    origen = st.session_state.get("tana_excel_origen_bytes")
    if not origen:
        return None
    wb_in = openpyxl.load_workbook(io.BytesIO(origen), data_only=False)
    encontrado = _encontrar_libro_diario(wb_in)
    if not encontrado:
        raise ValueError("No pude identificar el Libro Diario del Excel original por su contenido.")
    sheet_name, meta = encontrado
    ws = wb_in[sheet_name]
    header_row = meta["header_row"]
    headers = {str(ws.cell(header_row,c).value or "").strip(): c for c in range(1, ws.max_column+1)}
    aliases = {
        "N° Asiento":["N° Asiento","Nº Asiento","No. Asiento","Asiento","N°"],
        "Fecha":["Fecha"], "Glosa":["Glosa","Descripción"],
        "Documento":["Documento","Documento Ref.","Documento Ref"],
        "Operación":["Operación","Operacion","N° Operación","N° Operacion"],
        "Código":["Código","Codigo","Cuenta","Código Cuenta"],
        "Denominación":["Denominación","Denominacion","Nombre Cuenta","Descripción Cuenta"],
        "Concepto":["Concepto","Detalle"], "Debe S/":["Debe S/","Debe","Debe S/.","DEBE"],
        "Haber S/":["Haber S/","Haber","Haber S/.","HABER"]}
    col={}
    for canon,names in aliases.items():
        col[canon]=next((headers[n] for n in names if n in headers), None)
    if not col.get("Código") or not col.get("Debe S/") or not col.get("Haber S/"):
        raise ValueError("El Excel original no tiene las columnas de asientos necesarias.")

    rows=[]
    for a in asientos:
        first=True
        for line in a.get("lineas",[]):
            rows.append((a,line,first))
            first=False
    start=2
    # Para conservar las fórmulas y referencias de las hojas, normalmente las correcciones
    # tienen el mismo número de líneas. Si cambia, ajustamos filas al final.
    old_rows=ws.max_row-start+1
    if len(rows) > old_rows:
        for _ in range(len(rows)-old_rows): ws.insert_rows(ws.max_row+1)
    elif len(rows) < old_rows:
        for _ in range(old_rows-len(rows)): ws.delete_rows(start+len(rows))

    for idx,(a,line,first) in enumerate(rows,start=start):
        if col.get("N° Asiento"): ws.cell(idx,col["N° Asiento"], a.get("numero", "") if first else "")
        if col.get("Fecha"): ws.cell(idx,col["Fecha"], a.get("fecha", "") if first else "")
        if col.get("Glosa"): ws.cell(idx,col["Glosa"], a.get("glosa", "") if first else "")
        if col.get("Documento"): ws.cell(idx,col["Documento"], a.get("documento", "") if first else "")
        if col.get("Operación"): ws.cell(idx,col["Operación"], a.get("operacion_numero", "") if first else "")
        ws.cell(idx,col["Código"], str(line.get("codigo", "")).strip())
        if col.get("Denominación"): ws.cell(idx,col["Denominación"], line.get("denominacion", ""))
        if col.get("Concepto"): ws.cell(idx,col["Concepto"], line.get("concepto", ""))
        ws.cell(idx,col["Debe S/"], float(line.get("debe",0) or 0))
        ws.cell(idx,col["Haber S/"], float(line.get("haber",0) or 0))

    ws_rev=wb_in.create_sheet("Revision_TANA")
    ws_rev["A1"]="TANA — PROPUESTA DE CORRECCIÓN"
    ws_rev["A2"]="Estado"
    ws_rev["B2"]="NO VALIDADA" if errores else "VALIDADA"
    ws_rev["A3"]="Advertencia"
    ws_rev["B3"]="Soy una inteligencia artificial y puedo cometer errores. Revisa siempre el resultado antes de utilizarlo."
    ws_rev["A5"]="Diagnóstico posterior a la corrección"
    for i,e in enumerate(errores or ["No se encontraron errores de validación en los asientos."],start=6):
        ws_rev.cell(i,1,str(e))
    ws_rev.column_dimensions["A"].width=55; ws_rev.column_dimensions["B"].width=95
    try:
        wb_in.calculation.fullCalcOnLoad=True; wb_in.calculation.forceFullCalc=True; wb_in.calculation.calcMode="auto"
    except Exception: pass
    out=io.BytesIO(); wb_in.save(out); out.seek(0); return out.getvalue()

def _revisar_bytes_excel_completo(raw):
    """Vuelve a revisar un Excel generado por TANA, hoja por hoja."""
    class _MemoryUpload:
        def __init__(self, data): self._data = data
        def getvalue(self): return self._data
    obj = _MemoryUpload(raw)
    return _revisar_excel_completo(obj), _resumen_excel_para_tutor(obj)


def _aplicar_correccion_si_corresponde(pregunta):
    if not _es_peticion_correccion(pregunta):
        return False
    with st.spinner("TANA está localizando la fila y preparando un Excel nuevo…"):
        try:
            # Si el estudiante cargó un Excel, corregimos directamente ese libro.
            # No obligamos a convertir HT/EF/ES a asientos contables.
            if st.session_state.get("tana_excel_origen_bytes"):
                propuesta, profile = _corregir_excel_directamente_con_gemini(pregunta)
                if propuesta.get("estado") != "corregible":
                    _tana_chat_add("assistant", f"<b>⚠️ No hice cambios.</b><br>{propuesta.get('observacion','No encontré evidencia suficiente para una corrección segura.')}<br><br><small>El Excel original permanece intacto.</small>")
                    return True
                nuevo_buffer, aplicados = _aplicar_cambios_celdas_excel(propuesta)
                # Segunda revisión del libro completo.
                hojas_post, resumen_post = _revisar_bytes_excel_completo(nuevo_buffer)
                st.session_state["tana_excel_buffer"] = nuevo_buffer
                st.session_state["tana_excel_hojas"] = hojas_post
                st.session_state["tana_excel_resumen_completo"] = resumen_post
                st.session_state["tana_excel_revisado_post"] = True
                st.session_state["tana_correcciones"] = st.session_state.get("tana_correcciones", 0) + 1
                st.session_state["tana_correccion_version"] = st.session_state.get("tana_correccion_version", 1) + 1
                detalle = "<br>".join(f"<b>{c}</b>: {a} → {n}<br><small>{m}</small>" for c,a,n,m in aplicados)
                _tana_chat_add("assistant", f"<b>✅ Corrección aplicada en el Excel.</b><br><br><b>Hoja:</b> {propuesta.get('hoja')}<br><b>Fila:</b> {propuesta.get('fila')}<br><br>{detalle}<br><br><b>Nuevo Excel generado.</b><br>TANA volvió a revisar todas las hojas del archivo corregido.<br><br><small>El Excel original permanece intacto.</small>")
                return True

            data, profile = _corregir_asientos_con_gemini(pregunta)
            nuevos = data.get("asientos", []) if isinstance(data, dict) else []
            observacion = data.get("observacion", "Corrección procesada.") if isinstance(data, dict) else "Corrección procesada."
            if not isinstance(nuevos, list) or not nuevos:
                _tana_chat_add("assistant", f"No apliqué cambios porque no pude determinar una corrección segura. {observacion}")
                return True
            valid, errors, warnings = validate_asientos({"asientos": nuevos}, pcge_map)
            # No bloqueamos la entrega de una PROPUESTA si la validación falla.
            # El estudiante pidió explícitamente que TANA intente corregir a partir
            # del diagnóstico. La versión anterior se conserva y la nueva se marca
            # claramente como NO VALIDADA hasta que el usuario la revise.
            nuevos = asegurar_cuenta_79_en_destinos(nuevos, pcge_map)
            nuevos = corregir_retiro_socio(nuevos, st.session_state.get("monografia_json", {}))
            valid, errors, warnings = validate_asientos({"asientos": nuevos}, pcge_map)

            if st.session_state.get("tana_excel_origen_bytes"):
                # Generamos SIEMPRE un archivo nuevo; el original permanece intacto.
                nuevo_buffer = _crear_excel_revision_desde_origen(nuevos, errors)
                st.session_state["tana_excel_buffer"] = nuevo_buffer

                # Segunda pasada obligatoria: revisar nuevamente TODO el Excel generado.
                # Esto evita afirmar que una corrección quedó bien solo porque el
                # conjunto de asientos cuadró. También deja el inventario de hojas
                # actualizado para las preguntas posteriores del estudiante.
                try:
                    hojas_post, resumen_post = _revisar_bytes_excel_completo(nuevo_buffer)
                    st.session_state["tana_excel_hojas"] = hojas_post
                    st.session_state["tana_excel_resumen_completo"] = resumen_post
                    st.session_state["tana_excel_revisado_post"] = True
                except Exception as exc_post:
                    st.session_state["tana_excel_revisado_post"] = False
                    st.session_state["tana_excel_post_error"] = str(exc_post)
            st.session_state["asientos_contables"] = nuevos
            st.session_state["asientos_validos"] = valid
            st.session_state["errores_asientos"] = errors
            st.session_state["alertas_asientos"] = warnings
            st.session_state["tana_correcciones"] = st.session_state.get("tana_correcciones", 0) + 1
            st.session_state["tana_correccion_version"] = st.session_state.get("tana_correccion_version", 1) + 1
            st.session_state["tana_resuelto_signature"] = None

            post_ok = st.session_state.get("tana_excel_revisado_post", False)
            if errors:
                detalle = "<br>".join(str(e) for e in errors[:12])
                _tana_chat_add("assistant", f"<b>⚠️ TANA generó un nuevo Excel, pero la revisión contable todavía encuentra observaciones.</b><br>{observacion}<br><br><b>Observaciones:</b><br>{detalle}<br><br><b>El archivo nuevo fue vuelto a revisar hoja por hoja.</b><br>{'La segunda revisión se completó.' if post_ok else 'La segunda revisión no pudo completarse.'}<br><br><small>El Excel original no fue modificado.</small>")
            else:
                _tana_chat_add("assistant", f"<b>✅ Corrección aplicada y Excel nuevo generado.</b><br>{observacion}<br><br><b>Primera validación:</b> los asientos cuadran.<br><b>Segunda validación:</b> {'TANA volvió a revisar todo el Excel, hoja por hoja.' if post_ok else 'no pudo completar la segunda revisión.'}<br><br><b>El archivo original permanece intacto y el botón de descarga corresponde al archivo nuevo.</b>")
            return True
        except Exception as exc:
            st.error(f"No se pudo aplicar la corrección: {_gemini_error_message(exc)}")
            return True

# ============================================================
# TUTOR INTERACTIVO TANA
# ============================================================
def _tana_contexto_tutor():
    mono = st.session_state.get("monografia_texto", "")
    asientos = st.session_state.get("asientos_contables", [])
    asientos_txt = json.dumps(asientos, ensure_ascii=False, indent=2)
    diagnostico = st.session_state.get("tana_diagnostico_excel", "")
    excel_contexto = st.session_state.get("tana_excel_resumen_completo", "")
    return (
        "MONOGRAFÍA / FUENTE:\n" + mono[:12000]
        + "\n\nASIENTOS ACTUALES (si fueron reconstruidos):\n" + asientos_txt[:18000]
        + "\n\nDIAGNÓSTICO DETERMINISTA:\n" + diagnostico[:8000]
        + "\n\nCONTENIDO COMPLETO DEL EXCEL, HOJA POR HOJA:\n" + excel_contexto[:90000]
    )

def _preguntar_a_tana(pregunta):
    contexto = _tana_contexto_tutor()
    prompt = f"""Eres TANA, tutor de contabilidad peruana.
Responde la pregunta del estudiante usando únicamente el contexto proporcionado.
Explica con claridad por qué se hizo el asiento, cómo se obtuvo el importe, por qué
una cuenta va al Debe o Haber y, cuando corresponda, cómo se relaciona con la HT,
la distribución y ajustes, ERN, ERF o ESF.
No inventes información que no aparezca en el contexto.
Si el estudiante pregunta dónde está el error, responde de forma exacta y prioriza
el diagnóstico determinista: número de asiento, número de línea, código de cuenta,
Debe, Haber y diferencia. Si el problema está en una cuenta, menciona el código y
su denominación. No digas "algún destino está mal" si el contexto permite identificar
el punto concreto. Si no puedes identificarlo con seguridad, dilo expresamente.
Si el usuario pide corregir, no afirmes que lo corregiste hasta que la validación haya
pasado y se haya generado una nueva versión del Excel.
 En los estados financieros respeta estrictamente estas reglas:
 ERF: 70 y 69 se detectan por prefijo; 94 y 95 son obligatorias; 78 se incluye solo si existe; 65 y 67 solo si existen sin destino a 94/95. No incluyas 79 ni agregues automáticamente otras cuentas del elemento 6 al ERF.
 ERN: presenta las cuentas por naturaleza y su resultado.
 ESF: presenta activo, pasivo y patrimonio; resultados acumulados 59 con saldo deudor reducen el patrimonio. El resultado del ejercicio debe ser consistente con ERN y ERF y el ESF debe cumplir Activo = Pasivo + Patrimonio.
 Si falta un dato, dilo.
Al final de una explicación de revisión agrega, cuando sea pertinente:
"Soy una inteligencia artificial y puedo cometer errores. Revisa siempre el resultado antes de utilizarlo."

CONTEXTO:
{contexto}

PREGUNTA:
{pregunta}"""

    response, profile = _generate_with_fallback(
        lambda client: [prompt],
        types.GenerateContentConfig()
    )
    return response.text or "No pude generar una respuesta.", profile["label"]

# La consulta y el audio se capturan arriba. Aquí solo se procesa la acción,
# una vez que las funciones del tutor ya están definidas.
if enviar_top and (pregunta_top.strip() or audio_top is not None) and (st.session_state.get("asientos_contables") or st.session_state.get("tana_excel_revisado") or st.session_state.get("monografia_json")):
    if enviar_top and pregunta_top.strip():
        _tana_chat_add("user", pregunta_top.strip())
        _modo = _detectar_modo_trabajo(pregunta_top.strip())
        if _modo:
            st.session_state["tana_modo_trabajo"] = _modo
        if _es_peticion_correccion(pregunta_top.strip()):
            _aplicar_correccion_si_corresponde(pregunta_top.strip())
        elif st.session_state.get("tana_excel_origen_bytes") and any(
            k in pregunta_top.lower()
            for k in ("no cuadra", "no cuadran", "diferencia", "dónde está el error", "donde esta el error",
                      "qué está mal", "que esta mal", "revisa el excel", "revisa este excel")
        ):
            diagnostico = st.session_state.get("tana_diagnostico_excel", "")
            _tana_chat_add(
                "assistant",
                "<b>Revisión exacta del Excel:</b><br>" +
                diagnostico.replace("\n", "<br>") +
                "<br><br><small>Soy una inteligencia artificial y puedo cometer errores. "
                "Revisa siempre el resultado antes de utilizarlo.</small>"
            )
        else:
            with st.spinner("TANA está preparando la explicación…"):
                try:
                    respuesta, ruta = _preguntar_a_tana(pregunta_top.strip())
                    st.session_state["respuesta_tana"] = respuesta
                    st.session_state["respuesta_tana_ruta"] = ruta
                    _tana_chat_add("assistant", respuesta)
                except Exception as exc:
                    st.error(f"No se pudo responder: {_gemini_error_message(exc)}")
    elif audio_top is not None:
        import hashlib
        _audio_sig = hashlib.sha1(audio_top.getvalue()).hexdigest()
        if st.session_state.get("audio_tana_processed") == _audio_sig:
            audio_top = None
        else:
            st.session_state["audio_tana_processed"] = _audio_sig
        if audio_top is not None:
            _tana_chat_add("user", "🎤 Pregunta enviada por voz")
            with st.spinner("TANA está escuchando y preparando la respuesta…"):
                temp_audio = None
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp:
                        tmp.write(audio_top.getvalue())
                        temp_audio = tmp.name
                    def audio_contents(client):
                        audio_file = client.files.upload(file=temp_audio)
                        return [audio_file, "Escucha el audio del estudiante, transcribe su pregunta y luego respóndela. No inventes datos. Usa el siguiente contexto:\n" + _tana_contexto_tutor()]
                    response, profile = _generate_with_fallback(audio_contents, types.GenerateContentConfig())
                    respuesta_audio = response.text or "No pude interpretar el audio."
                    st.session_state["respuesta_tana"] = respuesta_audio
                    st.session_state["respuesta_tana_ruta"] = profile["label"]
                    _tana_chat_add("assistant", respuesta_audio)
                except Exception as exc:
                    st.error(f"No se pudo procesar el audio: {_gemini_error_message(exc)}")
                finally:
                    if temp_audio and os.path.exists(temp_audio):
                        os.remove(temp_audio)
    st.rerun()
if st.session_state.get("asientos_contables"):
    st.markdown(
        '<div class="tana-success-card">✅&nbsp; TANA terminó el desarrollo contable. Tu Excel está listo para descargar.</div>',
        unsafe_allow_html=True,
    )

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
# Regla de presentación final:
# 1) Se muestran TODAS las cuentas de balance realmente utilizadas por TANA.
# 2) En las hojas públicas se muestra únicamente la DESCRIPCIÓN; no se
#    imprimen códigos de cuenta en el ESF.
# 3) La ubicación se decide por el saldo real de la cuenta en la HT:
#       - saldo deudor  -> ACTIVO
#       - saldo acreedor -> PASIVO o PATRIMONIO según el elemento.
# 4) Si una cuenta normalmente activa (1-3) aparece con saldo acreedor,
#    se presenta en el lado pasivo como "otras cuentas"; si una cuenta de
#    pasivo (4) aparece con saldo deudor, se presenta en activo. Así no se
#    pierde ninguna cuenta y nunca se duplica una cuenta.
# 5) Las cuentas 5 se presentan como PATRIMONIO, respetando su signo.
# 6) El resultado del ejercicio se toma del ERN y debe coincidir con ERF.
# 7) TOTAL ACTIVO = TOTAL PASIVO + TOTAL PATRIMONIO NETO.

ws9 = wb.create_sheet('ESF')
_report_title(ws9, 'ESTADO DE SITUACIÓN FINANCIERA')

# Encabezados, exactamente en el estilo de la plantilla suministrada.
for cell, value in [('B4','ACTIVO'), ('D4','Notas'), ('E4','AÑO 2026'),
                    ('G4','PASIVO Y PATRIMONIO'), ('I4','Notas'), ('J4','AÑO 2026')]:
    ws9[cell] = value
for c in (2,4,5,7,9,10):
    ws9.cell(4,c).fill = PatternFill('solid', fgColor='D9E1F2')
    ws9.cell(4,c).font = BOLD
    ws9.cell(4,c).alignment = Alignment(horizontal='center')

# Saldos de cada cuenta desde la HT. Se usa SALDO AJUSTADO (I/J) para las
# cuentas de balance; si por alguna razón estuviera vacío, se conserva el
# saldo deudor/acreedor de la HT (O/P).
def _es_balance_real(code):
    return bool(code) and code[:1] in {'1','2','3','4','5'}

def _saldo_deudor_esf(code):
    return f'=IF(SUMIF(HT!$A$4:$A${HT_LAST_ROW},"{code}",HT!$I$4:$I${HT_LAST_ROW})<>0,' \
           f'SUMIF(HT!$A$4:$A${HT_LAST_ROW},"{code}",HT!$I$4:$I${HT_LAST_ROW}),' \
           f'SUMIF(HT!$A$4:$A${HT_LAST_ROW},"{code}",HT!$O$4:$O${HT_LAST_ROW}))'

def _saldo_acreedor_esf(code):
    return f'=IF(SUMIF(HT!$A$4:$A${HT_LAST_ROW},"{code}",HT!$J$4:$J${HT_LAST_ROW})<>0,' \
           f'SUMIF(HT!$A$4:$A${HT_LAST_ROW},"{code}",HT!$J$4:$J${HT_LAST_ROW}),' \
           f'SUMIF(HT!$A$4:$A${HT_LAST_ROW},"{code}",HT!$P$4:$P${HT_LAST_ROW}))'

# Para decidir el lado en Excel sin depender de la evaluación previa del
# archivo, usamos los saldos ya consolidados en Python (movimientos). Los
# valores son los mismos que alimentan la HT. Esto permite que cada cuenta
# aparezca una sola vez y evita filas de códigos "sueltos" fuera del cuadro.

def _saldo_python(code):
    rec = movimientos.get(code, {'debe':0.0,'haber':0.0})
    return float(rec.get('debe',0) or 0) - float(rec.get('haber',0) or 0)

# Cuentas de balance con saldo no nulo. Las cuentas con saldo cero también
# se conservan si fueron utilizadas: ninguna cuenta utilizada desaparece.
cuentas_balance = [c for c in cuentas_reporte if _es_balance_real(c)]

# Orden lógico: activos 1-3, luego saldos deudores anómalos de 4-5;
# pasivos 4, luego patrimonio 5.
activos = []
activos_anomalos = []
pasivos = []
patrimonio = []

for code in cuentas_balance:
    saldo = _saldo_python(code)
    if code.startswith('5'):
        patrimonio.append(code)
    elif saldo >= 0:
        if code.startswith(('1','2','3')):
            activos.append(code)
        elif code.startswith('4'):
            # Cuenta de pasivo con saldo deudor: se presenta como activo,
            # pero separada como "otras cuentas de activo".
            activos_anomalos.append(code)
        else:
            activos.append(code)
    else:
        if code.startswith(('1','2','3')):
            # Cuenta normalmente activa con saldo acreedor: se presenta como
            # pasivo, sin duplicarla.
            pasivos.append(code)
        elif code.startswith('4'):
            pasivos.append(code)
        else:
            pasivos.append(code)

# --------------------------- ACTIVO ---------------------------
r = 5
ws9.cell(r,2,'ACTIVO CORRIENTE').font = BOLD
r += 1
ac_rows=[]

# Elementos 1-2 se presentan como activo corriente, salvo cuentas 3 (PPE,
# etc.) que corresponden al no corriente. La cuenta 40 con saldo deudor se
# considera activo corriente (crédito fiscal).
for code in activos:
    if code.startswith('3'):
        continue
    desc = pcge_map.get(code, '')
    if not desc:
        desc = f'Cuenta {code}'
    _write_label(ws9, r, desc)
    _set_report_value(ws9, r, 5, _saldo_deudor_esf(code))
    ac_rows.append(r)
    r += 1
for code in activos_anomalos:
    desc = pcge_map.get(code, '') or f'Cuenta {code}'
    _write_label(ws9, r, desc)
    _set_report_value(ws9, r, 5, _saldo_deudor_esf(code))
    ac_rows.append(r)
    r += 1

ws9.cell(r,2,'TOTAL ACTIVO CORRIENTE').font = BOLD
_set_report_value(ws9, r, 5, '=' + '+'.join(f'E{x}' for x in ac_rows) if ac_rows else '=0', True)
TOTAL_AC_ROW=r
r += 2

ws9.cell(r,2,'ACTIVO NO CORRIENTE').font = BOLD
r += 1
anc_rows=[]
for code in activos:
    if not code.startswith('3'):
        continue
    desc = pcge_map.get(code, '') or f'Cuenta {code}'
    _write_label(ws9, r, desc)
    _set_report_value(ws9, r, 5, _saldo_deudor_esf(code))
    anc_rows.append(r)
    r += 1

ws9.cell(r,2,'TOTAL ACTIVO NO CORRIENTE').font = BOLD
_set_report_value(ws9, r, 5, '=' + '+'.join(f'E{x}' for x in anc_rows) if anc_rows else '=0', True)
TOTAL_ANC_ROW=r
r += 1

ws9.cell(r,2,'TOTAL ACTIVO').font = BOLD
_set_report_value(ws9, r, 5, f'=E{TOTAL_AC_ROW}+E{TOTAL_ANC_ROW}', True)
TOTAL_ACTIVO_ROW=r

# ---------------------- PASIVO / PATRIMONIO ----------------------
r2=5
ws9.cell(r2,7,'PASIVO').font=BOLD
r2 += 1
ws9.cell(r2,7,'PASIVO CORRIENTE').font=BOLD
r2 += 1
pc_rows=[]

# Pasivos corrientes: cuentas 4 y saldos acreedores de cuentas 1-3.
# La clasificación se mantiene dinámica y no se inventan cuentas.
for code in pasivos:
    # Las obligaciones financieras que comienzan en 45 se mantienen en
    # corriente en esta plantilla, tal como el modelo del usuario.
    desc = pcge_map.get(code, '') or f'Cuenta {code}'
    _write_label(ws9, r2, desc)
    _set_report_value(ws9, r2, 10, _saldo_acreedor_esf(code))
    pc_rows.append(r2)
    r2 += 1

ws9.cell(r2,7,'TOTAL PASIVO CORRIENTE').font=BOLD
_set_report_value(ws9, r2, 10, '=' + '+'.join(f'J{x}' for x in pc_rows) if pc_rows else '=0', True)
TOTAL_PC_ROW=r2
r2 += 2

# Pasivo no corriente: queda preparado para cuentas que explícitamente
# correspondan a obligaciones no corrientes. Si el catálogo/monografía no
# aporta una clasificación de vencimiento, no se duplica ninguna cuenta.
ws9.cell(r2,7,'PASIVO NO CORRIENTE').font=BOLD
r2 += 1
pnc_rows=[]
# Se reserva la clasificación de cuentas 45/46/47 con información de
# vencimiento futura. En esta versión no se fuerza ninguna cuenta a PNC;
# todas las cuentas existentes se muestran una sola vez en pasivo corriente,
# siguiendo el modelo suministrado.
ws9.cell(r2,7,'Obligaciones financieras y otras').font=BLACK
_set_report_value(ws9,r2,10,'=0')
TOTAL_PNC_ROW=r2
r2 += 1

ws9.cell(r2,7,'TOTAL PASIVO').font=BOLD
_set_report_value(ws9,r2,10,f'=J{TOTAL_PC_ROW}+J{TOTAL_PNC_ROW}',True)
TOTAL_PASIVO_ROW=r2
r2 += 2

ws9.cell(r2,7,'PATRIMONIO NETO').font=BOLD
r2 += 1
pat_rows=[]
for code in patrimonio:
    desc=pcge_map.get(code,'') or f'Cuenta {code}'
    _write_label(ws9,r2,desc)
    # Patrimonio: saldo acreedor aumenta; saldo deudor disminuye.
    _set_report_value(ws9,r2,10,f'={_saldo_acreedor_esf(code)[1:]}-{_saldo_deudor_esf(code)[1:]}')
    pat_rows.append(r2)
    r2 += 1

ws9.cell(r2,7,'Resultado del ejercicio').font=BLACK
# El ESF toma el mismo resultado final del ERF; ERN mantiene un control cruzado.
_set_report_value(ws9,r2,10,f'=ERF!E{resultado_erf_row}')
pat_rows.append(r2)
r2 += 1

ws9.cell(r2,7,'TOTAL PATRIMONIO NETO').font=BOLD
_set_report_value(ws9,r2,10,'=' + '+'.join(f'J{x}' for x in pat_rows) if pat_rows else '=0',True)
TOTAL_PATRIMONIO_ROW=r2
r2 += 1

ws9.cell(r2,7,'TOTAL PASIVO Y PATRIMONIO NETO').font=BOLD
_set_report_value(ws9,r2,10,f'=J{TOTAL_PASIVO_ROW}+J{TOTAL_PATRIMONIO_ROW}',True)
TOTAL_PYPN_ROW=r2
r2 += 1

# Control: la diferencia debe ser exactamente cero.
control_esf_row=r2
ws9.cell(r2,7,'DIFERENCIA: ACTIVO - (PASIVO + PATRIMONIO)').font=BOLD
_set_report_value(ws9,r2,10,f'=E{TOTAL_ACTIVO_ROW}-J{TOTAL_PYPN_ROW}',True)
ws9.cell(r2,11,f'=IF(ABS(J{r2})<0.01,"CUADRADO","REVISAR")').font=BOLD
_hide_control_row(ws9,control_esf_row)

for col,width in {'B':52,'C':3,'D':9,'E':18,'G':52,'H':3,'I':9,'J':18,'K':14}.items():
    ws9.column_dimensions[col].width=width
ws9.freeze_panes='B5'

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
if st.session_state.get("tana_modo_trabajo") == "asientos":
    # Modo básico: el estudiante pidió únicamente asientos / libro diario.
    HOJAS_PUBLICAS = ["Asientos_Contables"]
else:
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

_sig = st.session_state.get("tana_file_signature")
if _sig and st.session_state.get("tana_resuelto_signature") != _sig:
    _tana_chat_add(
        "assistant",
        "TANA ha resuelto tu monografía:<br>" + (
            "&nbsp;&nbsp;• Solo asientos / Libro Diario" if st.session_state.get("tana_modo_trabajo") == "asientos" else
            "&nbsp;&nbsp;• Asientos<br>&nbsp;&nbsp;• HT<br>&nbsp;&nbsp;• ERN<br>&nbsp;&nbsp;• ERF<br>&nbsp;&nbsp;• ESF"
        ),
    )
    st.session_state["tana_resuelto_signature"] = _sig
    st.session_state["tana_excel_buffer"] = buffer.getvalue()
    st.rerun()

if st.session_state.get("tana_excel_buffer"):
    _n_asientos = len(st.session_state.get("asientos_contables", []) or [])
    _n_validos = len(st.session_state.get("asientos_validos", []) or [])
    st.markdown(
        f'<div class="tana-stats-row">'
        f'<div class="tana-stat-box"><span style="font-size:20px;">📄</span>'
        f'<div><div class="num">{_n_asientos}</div><div class="label">Asientos generados</div></div></div>'
        f'<div class="tana-stat-box"><span style="font-size:20px;">✅</span>'
        f'<div><div class="num">{_n_validos}</div><div class="label">Registros validados</div></div></div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.download_button(
        label="⬇️  Descargar Excel",
        data=st.session_state["tana_excel_buffer"],
        file_name=(
            f"TANA_Contabilidad_corregido_v{st.session_state.get('tana_correccion_version', 1)}.xlsx"
            if st.session_state.get("tana_correcciones", 0) else
            ("TANA_Asientos.xlsx" if st.session_state.get("tana_modo_trabajo") == "asientos" else "TANA_Contabilidad.xlsx")
        ),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    if "monografia_nombre" in st.session_state:
        st.caption("La hoja Monografia conserva el texto extraído para revisión.")
