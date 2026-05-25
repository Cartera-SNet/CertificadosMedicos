"""
Bot Certificaciones Médicas SOAT — v2.0
Migración API REST — SIS / Esculapio Soft
"""

import os, csv, time, threading, zipfile, logging, json, shutil, re, base64, tempfile
from datetime import datetime, timedelta
from pathlib import Path
from io import BytesIO
from copy import deepcopy

from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for

try:
    import requests as req_lib
    REQUESTS_OK = True
except ImportError:
    REQUESTS_OK = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    from PIL import Image
    PIL_OK = True
except ImportError:
    PIL_OK = False

# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# APIs Esculapio — los servidores y empresas vienen de la red
# ─────────────────────────────────────────────────────────────
API_SERVIDORES = "https://appsintranet.esculapiosis.com/ApiCampbell/api/Usuarios/obtener-servidores"
API_EMPRESAS   = "https://appsintranet.esculapiosis.com/ApiCampbell/api/Usuarios/obtener-empresas"

# Cache en memoria para no re-consultar en cada petición
_cache_servidores = None   # lista de dicts con $id, Descripcion, Ip_Conexion, Puerto, Usuario, Clave
_cache_servidores_ts = 0   # timestamp


def _get_servidores_api():
    """Obtiene la lista de servidores desde la API. Cache de 5 minutos."""
    import time
    global _cache_servidores, _cache_servidores_ts
    if _cache_servidores and (time.time() - _cache_servidores_ts) < 300:
        return _cache_servidores
    try:
        resp = req_lib.get(API_SERVIDORES, timeout=10, verify=False)
        resp.raise_for_status()
        data = resp.json()
        _cache_servidores    = data if isinstance(data, list) else []
        _cache_servidores_ts = time.time()
        return _cache_servidores
    except Exception as e:
        log(f"Error consultando servidores API: {e}", "warn")
        return _cache_servidores or []


def _get_empresas_api(ip, puerto, usuario, password):
    """Valida credenciales y obtiene empresas del usuario en ese servidor."""
    try:
        params = {
            "IpConexion":   ip,
            "BdConexion":   "bd",
            "PortConexion": puerto,
            "Usuario":      usuario,
            "PasswordUsu":  password,
        }
        resp = req_lib.get(API_EMPRESAS, params=params, timeout=15, verify=False)
        resp.raise_for_status()
        data = resp.json()
        return data if isinstance(data, list) else []
    except req_lib.exceptions.HTTPError as e:
        log(f"API empresas HTTP error: {e}", "warn")
        return []
    except Exception as e:
        log(f"API empresas error: {e}", "warn")
        return []


API_BASE = "https://appsintranet.esculapiosis.com/ApiCampbell/api/Reportes/obtener-certificacionmedica"

# ─────────────────────────────────────────────────────────────
# FLASK APP
# ─────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(message)s")
L = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────
# FUENTES CALIBRI/CAMBRIA — descarga automática al arrancar
# Carlito = clon libre de Calibri (métricas idénticas)
# Caladea = clon libre de Cambria (métricas idénticas)
# ─────────────────────────────────────────────────────────────
FONTS_DIR = Path(__file__).parent / "fonts"

def _descargar_fuentes():
    """Descarga Carlito y Caladea si no están en la carpeta fonts/."""
    import urllib.request
    FONTS_DIR.mkdir(exist_ok=True)
    fuentes = {
        "Carlito-Regular.ttf": "https://github.com/googlefonts/carlito/raw/main/fonts/ttf/Carlito-Regular.ttf",
        "Carlito-Bold.ttf":    "https://github.com/googlefonts/carlito/raw/main/fonts/ttf/Carlito-Bold.ttf",
        "Caladea-Regular.ttf": "https://github.com/huertatipografica/Caladea/raw/master/fonts/ttf/Caladea-Regular.ttf",
        "Caladea-Bold.ttf":    "https://github.com/huertatipografica/Caladea/raw/master/fonts/ttf/Caladea-Bold.ttf",
    }
    headers = {"User-Agent": "Mozilla/5.0"}
    for nombre, url in fuentes.items():
        dest = FONTS_DIR / nombre
        if dest.exists() and dest.stat().st_size > 10000:
            continue  # ya existe
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=20) as r:
                dest.write_bytes(r.read())
            logging.getLogger(__name__).info(f"Fuente descargada: {nombre}")
        except Exception as e:
            logging.getLogger(__name__).warning(f"No se pudo descargar {nombre}: {e}")

# Descargar fuentes al iniciar
try:
    _descargar_fuentes()
except Exception:
    pass


app      = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "sis-soat-2025-key-default")
app.permanent_session_lifetime = timedelta(hours=8)  # Sesión expira en 8 horas
BASE_DIR = Path(__file__).parent

# ── HTTPS forzado (solo activo en Railway/producción) ──────────────────
@app.before_request
def forzar_https():
    """Redirige HTTP → HTTPS cuando corre detrás del proxy de Railway."""
    proto = request.headers.get("X-Forwarded-Proto", "")
    if proto == "http":
        url = request.url.replace("http://", "https://", 1)
        return redirect(url, code=301)
FORMATO  = BASE_DIR / "FORMATO.docx"
DOWNLOAD_DIR = BASE_DIR / "downloads"
DOWNLOAD_DIR.mkdir(exist_ok=True)

# Estado global del proceso
S = {
    "running": False, "stopping": False, "logs": [],
    "stats": {"total": 0, "ok": 0, "err": 0},
    "finished": False, "error": None, "zip_path": None,
    "session_dir": None, "completadas": set(),
    "exitosos": [], "errores": [],
}
LK = threading.Lock()

class StopBotException(Exception):
    pass

# ─────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────
def log(msg, lvl="info"):
    ts = datetime.now().strftime("%H:%M:%S")
    with LK:
        S["logs"].append({"ts": ts, "msg": msg, "level": lvl})
    L.info(f"[{ts}] {msg}")

def reset():
    with LK:
        S.update({
            "running": False, "stopping": False, "logs": [],
            "stats": {"total": 0, "ok": 0, "err": 0},
            "finished": False, "error": None, "zip_path": None,
            "session_dir": None, "completadas": set(),
            "exitosos": [], "errores": [],
        })

# ─────────────────────────────────────────────────────────────
# AUTENTICACIÓN
# ─────────────────────────────────────────────────────────────
def get_users():
    """
    Lee usuarios desde variable de entorno BOT_USERS.
    Formato JSON: {"admin": "password123", "usuario1": "clave456"}
    Si no existe, usa usuario por defecto: admin / soat2025
    """
    raw = os.environ.get("BOT_USERS", "")
    if raw:
        try:
            return json.loads(raw)
        except Exception:
            pass
    return {"admin": os.environ.get("BOT_PASSWORD", "soat2025")}

# Almacén de intentos fallidos por IP: {ip: {"intentos": N, "bloqueado_hasta": timestamp}}
_login_intentos = {}
_LOGIN_MAX_INTENTOS = 5      # máximo intentos fallidos
_LOGIN_VENTANA_SEG  = 600    # ventana de 10 minutos
_LOGIN_BLOQUEO_SEG  = 900    # bloqueo de 15 minutos

def _check_rate_limit(ip):
    """
    Verifica si la IP está bloqueada.
    Retorna (permitido: bool, mensaje: str, segundos_restantes: int)
    """
    ahora = time.time()
    datos = _login_intentos.get(ip, {})

    bloqueado_hasta = datos.get("bloqueado_hasta", 0)
    if bloqueado_hasta > ahora:
        restantes = int(bloqueado_hasta - ahora)
        minutos   = restantes // 60
        return False, f"IP bloqueada. Intenta en {minutos} min {restantes%60} seg.", restantes

    # Limpiar si la ventana expiró
    primer_intento = datos.get("primer_intento", 0)
    if ahora - primer_intento > _LOGIN_VENTANA_SEG:
        _login_intentos.pop(ip, None)

    return True, "", 0

def _registrar_intento_fallido(ip):
    """Registra un intento fallido y bloquea si supera el límite."""
    ahora = time.time()
    datos = _login_intentos.get(ip, {})

    if not datos or ahora - datos.get("primer_intento", 0) > _LOGIN_VENTANA_SEG:
        datos = {"intentos": 0, "primer_intento": ahora}

    datos["intentos"] += 1

    if datos["intentos"] >= _LOGIN_MAX_INTENTOS:
        datos["bloqueado_hasta"] = ahora + _LOGIN_BLOQUEO_SEG
        log(f"IP {ip} bloqueada por {_LOGIN_BLOQUEO_SEG//60} min ({datos['intentos']} intentos fallidos)", "warn")

    _login_intentos[ip] = datos

def _registrar_login_exitoso(ip):
    """Limpia el contador de intentos tras login exitoso."""
    _login_intentos.pop(ip, None)

def login_requerido(f):
    """Decorador que protege rutas — redirige a /login si no autenticado."""
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("autenticado"):
            if request.is_json or request.path.startswith("/api/"):
                return jsonify({"error": "No autorizado", "login_required": True}), 401
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return decorated

# ─────────────────────────────────────────────────────────────
# PERSISTENCIA (reanudación)
# ─────────────────────────────────────────────────────────────
def cargar_progreso(session_dir):
    p = session_dir / "progreso.json"
    if p.exists():
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            c = data.get("completadas", [])
            return set(c) if isinstance(c, list) else set(c.keys())
        except Exception as e:
            log(f"Warning progreso: {e}", "warn")
    return set()

def guardar_progreso(session_dir, completadas):
    try:
        with open(session_dir / "progreso.json", "w", encoding="utf-8") as f:
            json.dump({"completadas": list(completadas),
                       "actualizado": datetime.now().isoformat()}, f, indent=2)
    except Exception as e:
        log(f"Warning guardar progreso: {e}", "warn")

def guardar_casos_sesion(session_dir, casos, config):
    try:
        with open(session_dir / "sesion_casos.json", "w", encoding="utf-8") as f:
            json.dump({"casos": casos, "config": config,
                       "guardado": datetime.now().isoformat()},
                      f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"Warning guardar casos: {e}", "warn")

def cargar_casos_sesion(session_dir):
    p = session_dir / "sesion_casos.json"
    if not p.exists():
        return [], {}
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("casos", []), data.get("config", {})
    except Exception as e:
        log(f"Warning cargar casos: {e}", "warn")
        return [], {}

def buscar_sesion_pendiente():
    for d in DOWNLOAD_DIR.iterdir():
        if d.is_dir() and d.name.startswith("sesion_"):
            if (d / "progreso.json").exists() and not (d / "finalizado.txt").exists():
                return d
    return None

def marcar_sesion_finalizada(session_dir):
    (session_dir / "finalizado.txt").write_text(datetime.now().isoformat())

# ─────────────────────────────────────────────────────────────
# LLAMADA A LA API
# ─────────────────────────────────────────────────────────────
def llamar_api(config, no_cuenta):
    """Llama a la API de certificación médica y retorna el JSON del primer resultado."""
    params = {
        "IpConexion":   config["ip"],
        "BdConexion":   "bd",
        "PortConexion": config["puerto"],
        "Usuario":      config["usuario"],
        "PasswordUsu":  config["password"],
        "codigoEmp":    config["cod_empresa"],
        "NoCuenta":     no_cuenta,
    }
    try:
        resp = req_lib.get(
            API_BASE,
            params=params,
            timeout=120,
            verify=False,
            stream=False,
        )

        status_code = resp.status_code

        if status_code == 500:
            raise Exception(f"Sin certificado médico para caso {no_cuenta} en empresa '{config.get('cod_empresa','')}' — omitiendo")
        elif status_code == 404:
            raise Exception(f"Caso {no_cuenta} no encontrado en el servidor")
        elif status_code in (401, 403):
            raise Exception(f"Sin autorización para caso {no_cuenta} (HTTP {status_code})")
        elif status_code != 200:
            raise Exception(f"Error del servidor HTTP {status_code} para caso {no_cuenta}")

        content = resp.content
        text    = content.decode('utf-8', errors='replace')

        data = json.loads(text)
        if isinstance(data, list) and len(data) > 0:
            item = data[0]
            firma = item.get("FirmaMedico", "")
            if firma:
                log(f"  → Firma recibida: {len(firma):,} caracteres", "info")
            return item
        elif isinstance(data, dict):
            return data
        return None
    except req_lib.exceptions.Timeout:
        raise Exception(f"Timeout (120s) consultando caso {no_cuenta} — servidor lento")
    except req_lib.exceptions.ConnectionError:
        raise Exception(f"Sin conexión al servidor {config.get('ip','')}:{config.get('puerto','')}")
    except Exception as e:
        raise Exception(str(e))

# ─────────────────────────────────────────────────────────────
# FORMATEO DE FECHAS
# ─────────────────────────────────────────────────────────────
MESES_ES = {
    1:"ene", 2:"feb", 3:"mar", 4:"abr", 5:"may", 6:"jun",
    7:"jul", 8:"ago", 9:"sep", 10:"oct", 11:"nov", 12:"dic"
}

def formatear_fecha_nac(fecha_str):
    """1985-12-04T00:00:00 → 04-dic.-1985"""
    try:
        dt = datetime.fromisoformat(fecha_str.replace("T", " ").split(" ")[0])
        mes = MESES_ES[dt.month]
        return f"{dt.day:02d}-{mes}.-{dt.year}"
    except Exception:
        return fecha_str

# ─────────────────────────────────────────────────────────────
# FIRMA: BASE64 BMP → PNG TEMPORAL
# ─────────────────────────────────────────────────────────────
def procesar_firma_base64(base64_string, ruta_salida=None):
    """
    Convierte una cadena Base64 (como la de FirmaMedico) en una imagen PNG.
    Base64 → Binary → PIL Image → PNG con fondo transparente.
    La firma BMP real tiene ~483,424 caracteres en base64.
    """
    try:
        if not base64_string or len(base64_string) < 100:
            log(f"  → Firma demasiado corta ({len(base64_string)} chars), ignorada", "warn")
            return None

        # 1. Padding por si la cadena viene incompleta en el transporte
        missing_padding = len(base64_string) % 4
        if missing_padding:
            base64_string += '=' * (4 - missing_padding)

        # 2. Decodificar Base64 → bytes binarios
        image_data = base64.b64decode(base64_string)

        # Verificar cabecera BMP (debe empezar con 'BM' = 0x42 0x4D)
        if len(image_data) < 2 or image_data[0:2] != b'BM':
            log(f"  → Firma no es BMP válido (header: {image_data[0:4].hex()})", "warn")
            # Intentar de todas formas
        
        log(f"  → Firma decodificada: {len(image_data):,} bytes", "info")

        # 3. Abrir con PIL (soporta BMP, PNG, JPG, etc.)
        img = Image.open(BytesIO(image_data))
        log(f"  → Firma imagen: {img.size[0]}x{img.size[1]}px modo={img.mode}", "info")

        # 4. Convertir a RGBA para manejo de transparencia
        img = img.convert("RGBA")

        # 5. Fondo blanco → transparente
        data = img.getdata()
        new_data = []
        for r, g, b, a in data:
            if r > 240 and g > 240 and b > 240:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append((r, g, b, a))
        img.putdata(new_data)

        # 6. Recortar whitespace usando PIL getbbox() — sin dependencia numpy
        # Los BMP de Esculapio tienen padding blanco que infla el tamaño real
        try:
            # Usar solo el canal alpha para detectar el área con contenido
            alpha = img.split()[3]
            bbox = alpha.getbbox()  # (left, top, right, bottom) del contenido no-transparente
            if bbox:
                PAD = 8
                left   = max(0, bbox[0] - PAD)
                top    = max(0, bbox[1] - PAD)
                right  = min(img.width,  bbox[2] + PAD)
                bottom = min(img.height, bbox[3] + PAD)
                img = img.crop((left, top, right, bottom))
        except Exception:
            pass  # si falla, usar imagen completa

        # 6. Guardar o retornar
        if ruta_salida:
            img.save(str(ruta_salida), format="PNG")
            log(f"  → Firma PNG guardada: {Path(ruta_salida).stat().st_size:,} bytes", "info")
            return ruta_salida

        return img

    except Exception as e:
        log(f"Error al procesar la firma: {e}", "warn")
        return None


def firma_b64_a_png(b64_str, ruta_png: Path) -> bool:
    """Wrapper que usa procesar_firma_base64 y retorna bool."""
    resultado = procesar_firma_base64(b64_str, ruta_salida=str(ruta_png))
    return resultado is not None

# ─────────────────────────────────────────────────────────────
# GENERACIÓN DEL DOCX RELLENADO
# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# GENERACIÓN DEL DOCX RELLENADO
# ─────────────────────────────────────────────────────────────
def _colapsar_runs(parrafo):
    """
    Las variables en el DOCX están fragmentadas en múltiples runs
    (ej: '{', 'FirmaMedico', '}' en runs separados).
    Este helper colapsa todos los runs en uno solo preservando el
    formato del primer run, para que el reemplazo funcione.
    """
    from docx.oxml.ns import qn
    if not parrafo.runs:
        return

    texto_completo = parrafo.text
    if not texto_completo.strip():
        return

    # Guardar propiedades del primer run con contenido
    primer_run = None
    for r in parrafo.runs:
        if r.text:
            primer_run = r
            break
    if not primer_run:
        return

    font_name  = primer_run.font.name
    font_size  = primer_run.font.size
    bold       = primer_run.bold
    italic     = primer_run.italic

    # Eliminar todos los runs del párrafo a nivel XML
    p_elem = parrafo._p
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    for r in p_elem.findall(f'{{{ns}}}r'):
        p_elem.remove(r)

    # Agregar un único run con el texto completo
    run_nuevo = parrafo.add_run(texto_completo)
    if font_name:  run_nuevo.font.name = font_name
    if font_size:  run_nuevo.font.size = font_size
    if bold is not None:   run_nuevo.bold   = bold
    if italic is not None: run_nuevo.italic = italic



# ─────────────────────────────────────────────────────────────
# GENERACIÓN DIRECTA DE PDF CON REPORTLAB
# Layout idéntico al PDF de referencia (MUESTRA.pdf)
# Sin Word ni LibreOffice — funciona igual en Windows y Railway
# ─────────────────────────────────────────────────────────────
def generar_pdf_directo(datos_api, no_cuenta, ruta_png_firma, session_dir) -> Path:
    """
    PDF 100% idéntico al MUESTRA.pdf de Crystal Reports.
    Todas las coordenadas son absolutas, medidas con pdfminer.
    NO hay lógica dinámica — posiciones fijas exactas.
    """
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    W, H = LETTER  # 612 x 792

    def _reg(name, path):
        try:
            pdfmetrics.registerFont(TTFont(name, str(path)))
            return True
        except Exception:
            return False

    FD   = FONTS_DIR
    LLIB = "/usr/share/fonts/truetype/liberation"

    # Fuentes Crystal (mismas del MUESTRA) + Carlito como fallback completo
    _reg("CR1", FD/"CrystalF1.ttf")
    _reg("CR2", FD/"CrystalF2.ttf")
    if _reg("F1", FD/"Carlito-Bold.ttf") and _reg("F2", FD/"Carlito-Regular.ttf"):
        F1 = "F1"; F2 = "F2"
    elif _reg("F1", f"{LLIB}/LiberationSans-Bold.ttf") and _reg("F2", f"{LLIB}/LiberationSans-Regular.ttf"):
        F1 = "F1"; F2 = "F2"
    else:
        F1 = "Helvetica-Bold"; F2 = "Helvetica"

    # Construir charset de Crystal para seleccionar fuente por string
    try:
        from fontTools.ttLib import TTFont as _TTFTools
        _cmap1 = set(_TTFTools(str(FD/"CrystalF1.ttf")).getBestCmap().keys())
        _cmap2 = set(_TTFTools(str(FD/"CrystalF2.ttf")).getBestCmap().keys())
    except Exception:
        _cmap1 = set(); _cmap2 = set()

    def sf(text, bold=False):
        """Selecciona Crystal si todos los chars están disponibles, si no Carlito."""
        cmap = _cmap1 if bold else _cmap2
        if cmap and all(ord(ch) in cmap for ch in text):
            return "CR1" if bold else "CR2"
        return F1 if bold else F2

    nombre_empresa      = str(datos_api.get("nombre_empresa", ""))
    paciente            = str(datos_api.get("Paciente", ""))
    tipo_documento      = str(datos_api.get("TipoDocumento", ""))
    sexo                = str(datos_api.get("sexo", ""))
    cod_servicio        = str(datos_api.get("CodServicio", ""))
    nom_servicio        = str(datos_api.get("NomServicio", ""))
    edad                = str(int(float(datos_api.get("Edad", 0))))
    medida_edad         = str(datos_api.get("Medida_Edad", ""))
    fecha_nac           = formatear_fecha_nac(str(datos_api.get("Fecha_Nac", "")))
    fecha_ingreso       = str(datos_api.get("Fecha_Ingreso", ""))
    medico              = str(datos_api.get("Medico", ""))
    tarjeta_profesional = str(datos_api.get("Tarjeta_Profesional", ""))

    pdf_path = session_dir / f"CM_{no_cuenta}.pdf"
    c = canvas.Canvas(str(pdf_path), pagesize=LETTER)

    # Convierte y_top (medido desde arriba) → y ReportLab (desde abajo)
    def yt(y_top): return H - y_top

    # ══════════════════════════════════════════════════════════
    # TÍTULO — CIDFont+F1 sz=11.977, centrado
    # Línea 1: x=117.24→466.93  y_top=24.21
    # Línea 2: x=208.20→376.42  y_top=38.73
    # ══════════════════════════════════════════════════════════
    c.setFont(sf("CERTIFICADO DE ATENCION MEDICA PARA VICTIMAS DE", bold=True), 11.977)
    c.drawCentredString(W/2, yt(24.21), "CERTIFICADO DE ATENCION MEDICA PARA VICTIMAS DE")
    c.setFont(sf("ACCIDENTES DE TRÁNSITO", bold=True), 11.977)
    c.drawCentredString(W/2, yt(38.73), "ACCIDENTES DE TRÁNSITO")

    # ══════════════════════════════════════════════════════════
    # CUADRO BORDE
    # x=27.84→581.04  y_top=54.72→172.92  w=553.2  h=118.2
    # ══════════════════════════════════════════════════════════
    c.setLineWidth(0.8)
    c.roundRect(27.84, yt(172.92), 553.2, 118.2, 4, stroke=1, fill=0)

    # ══════════════════════════════════════════════════════════
    # CONTENIDO DEL CUADRO — CIDFont+F2 sz=9.941
    # ══════════════════════════════════════════════════════════
    c.setFont(F2, 9.941)

    # Línea 1: intro — x=31.92  y_top=62.07
    # En MUESTRA es una sola línea que cabe en el ancho del cuadro
    # Si el nombre de empresa es muy largo, hacer wrap en y=62.07 y 74.07
    intro = (f"El suscrito médico de la Institución Prestadora de Servicios  "
             f"{nombre_empresa} certifica que atendió a:")
    MAX_W = 541.99 - 31.92  # ancho exacto del MUESTRA

    if c.stringWidth(intro, F2, 9.941) <= MAX_W:
        # Cabe en una línea — exacto al MUESTRA
        c.setFont(sf(intro), 9.941)
        c.drawString(31.92, yt(67.07), intro)
        # Línea 2: paciente
        c.setFont(sf(f"Señor(a)   {paciente}"), 9.941)
        c.drawString(31.92, yt(79.07), f"Señor(a)   {paciente}")
    else:
        # Wrap: partir en dos líneas usando las mismas y del MUESTRA
        words = intro.split(); lines = []; line = ""
        for word in words:
            test = (line + " " + word).strip()
            if c.stringWidth(test, F2, 9.941) <= MAX_W:
                line = test
            else:
                if line: lines.append(line)
                line = word
        if line: lines.append(line)
        c.drawString(31.92, yt(62.07), lines[0])
        if len(lines) > 1:
            c.setFont(sf(lines[1]), 9.941)
        c.drawString(31.92, yt(79.07), lines[1])
        # Paciente en siguiente línea disponible
        c.setFont(sf(f"Señor(a)   {paciente}"), 9.941)
        c.drawString(31.92, yt(79.07 + 12.0), f"Señor(a)   {paciente}")

    # ── COLUMNA IZQUIERDA ──────────────────────────────────────
    # Identificación: x=35.64  y_top=113.91
    # Sexo:           x=35.64  y_top=132.63
    # Servicio:       x=35.64  y_top=152.07
    c.setFont(sf(f"Identificación: {tipo_documento}"), 9.941)
    c.drawString(35.64, yt(118.91), f"Identificación: {tipo_documento}")
    c.setFont(sf(f"Sexo:  {sexo}"), 9.941)
    c.drawString(35.64, yt(137.63), f"Sexo:  {sexo}")
    c.setFont(sf(f"Servicio Ingreso:  {cod_servicio} - {nom_servicio}"), 9.941)
    c.drawString(35.64, yt(157.07), f"Servicio Ingreso:  {cod_servicio} - {nom_servicio}")

    # ── COLUMNA DERECHA ────────────────────────────────────────
    # Edad:         x=365.64  y_top=112.35
    # Fecha Nac:    x=365.64  y_top=132.63
    # Fecha Ingr:   x=365.64  y_top=152.07
    c.setFont(sf(f"Edad: {edad} {medida_edad}"), 9.941)
    c.drawString(365.64, yt(117.35), f"Edad: {edad} {medida_edad}")
    c.setFont(sf(f"Fecha Nacimiento:  {fecha_nac}"), 9.941)
    c.drawString(365.64, yt(137.63), f"Fecha Nacimiento:  {fecha_nac}")
    c.setFont(sf(f"Fecha y hora ingreso:  {fecha_ingreso}"), 9.941)
    c.drawString(365.64, yt(157.07), f"Fecha y hora ingreso:  {fecha_ingreso}")

    # ══════════════════════════════════════════════════════════
    # EL SUSCRITO — CIDFont+F1 sz=9.941, centrado
    # Línea 1: x=54.00→523.31  y_top=183.63
    # Línea 2: x=146.28→428.54 y_top=195.63
    # ══════════════════════════════════════════════════════════
    c.setFont(sf("EL SUSCRITO MEDICO CERTIFICA LA RELACION CAUSAL DIRECTA ENTRE LAS LESIONES QUE", bold=True), 9.941)
    c.drawCentredString(W/2, yt(183.63),
        "EL SUSCRITO MEDICO CERTIFICA LA RELACION CAUSAL DIRECTA ENTRE LAS LESIONES QUE")
    c.setFont(sf("PRESENTA EL PACIENTE Y EL ACCIDENTE DE TRANSITO.", bold=True), 9.941)
    c.drawCentredString(W/2, yt(195.63),
        "PRESENTA EL PACIENTE Y EL ACCIDENTE DE TRANSITO.")

    # ══════════════════════════════════════════════════════════
    # FIRMA — zona x=322.2→565.2, entre y_top=205.57 y y_top=275
    # ── FIRMA + LÍNEA ─────────────────────────────────────────
    # MUESTRA exacto: línea en y_top=282.24, Dr en y_top=284.07
    # Orden visual de arriba a abajo: FIRMA → LÍNEA → Dr.
    FIRMA_X  = 322.2
    FIRMA_XR = 565.2
    FIRMA_ZW = FIRMA_XR - FIRMA_X   # 243 pts ancho
    LINEA_Y_TOP = 282.24             # posición exacta del MUESTRA

    if ruta_png_firma and ruta_png_firma.exists():
        try:
            from PIL import Image as _PI
            img = _PI.open(str(ruta_png_firma))
            iw, ih = img.size
            ratio = iw / ih if ih > 0 else 2.5
            # Zona firma: de y_top=220 a y_top=270 (exacto del MUESTRA)
            # 50pt de alto disponible, centrado horizontalmente en zona derecha
            FIRMA_MAX_H = 55.0
            FIRMA_MAX_W = FIRMA_ZW * 0.75
            fw = min(FIRMA_MAX_W, FIRMA_MAX_H * ratio)
            fh = fw / ratio
            if fh > FIRMA_MAX_H: fh = FIRMA_MAX_H; fw = fh * ratio
            fx = FIRMA_X + (FIRMA_ZW - fw) / 2
            # Firma debe quedar CLARAMENTE arriba de la línea
            # fy = base de la imagen, se dibuja hacia arriba
            # Subimos 4pt extra para que haya espacio visible entre firma y línea
            fy = yt(LINEA_Y_TOP) + 4
            c.drawImage(str(ruta_png_firma), fx, fy, fw, fh, mask="auto")
        except Exception as e:
            log(f"  → Firma PDF error: {e}", "warn")

    # Línea horizontal — dibujada DESPUÉS de la firma (encima visualmente)
    c.setLineWidth(0.96)
    c.line(322.2, yt(LINEA_Y_TOP), 565.2, yt(LINEA_Y_TOP))

    # DR. y REG.M. — posición original del MUESTRA
    c.setFont(sf(f"Dr. {medico}"), 9.941)
    c.drawString(322.20, yt(293.07), f"Dr. {medico}")
    c.setFont(sf(f"Reg.M. {tarjeta_profesional}"), 8.983)
    c.drawString(322.20, yt(305.11), f"Reg.M. {tarjeta_profesional}")

    c.save()
    log(f"  → PDF generado directamente (reportlab)", "info")
    return pdf_path
def generar_docx(datos_api, no_cuenta, ruta_png_firma, session_dir) -> Path:
    """
    Copia FORMATO.docx, reemplaza variables en text boxes y cuerpo,
    inserta firma correctamente usando add_picture() directo.
    """
    from docx import Document
    from docx.shared import Cm
    import lxml.etree as etree
    import copy

    reemplazos = {
        "{nombre_empresa}":      str(datos_api.get("nombre_empresa", "")),
        "{Paciente}":            str(datos_api.get("Paciente", "")),
        "{TipoDocumen}":         str(datos_api.get("TipoDocumento", "")),
        "{s}":                   str(datos_api.get("sexo", "")),
        "{c}":                   str(datos_api.get("CodServicio", "")),
        "{codserv}":             str(datos_api.get("NomServicio", "")),
        "{Edad}":                str(int(float(datos_api.get("Edad", 0)))),
        "{medidae}":             str(datos_api.get("Medida_Edad", "")),
        "{Fecha_Nac}":           formatear_fecha_nac(str(datos_api.get("Fecha_Nac", ""))),
        "{Fecha_I}":             str(datos_api.get("Fecha_Ingreso", "")),
        "{Medico}":              str(datos_api.get("Medico", "")),
        "{Tarjeta_Profesional}": str(datos_api.get("Tarjeta_Profesional", "")),
    }

    ruta_docx = session_dir / f"CM_{no_cuenta}.docx"
    shutil.copy(FORMATO, ruta_docx)
    doc = Document(str(ruta_docx))

    W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # ── FIX 1: Centrar título ────────────────────────────────────────────
    # El estilo Ttulo1 no tiene jc=center definido → LibreOffice lo deja izquierda
    # Lo forzamos desde código en el párrafo P0
    import lxml.etree as _etree_init
    if doc.paragraphs:
        p0 = doc.paragraphs[0]
        pPr0 = p0._p.find(f'{{{W_NS}}}pPr')
        if pPr0 is None:
            pPr0 = _etree_init.SubElement(p0._p, f'{{{W_NS}}}pPr')
            p0._p.insert(0, pPr0)
        jc0 = pPr0.find(f'{{{W_NS}}}jc')
        if jc0 is None:
            jc0 = _etree_init.SubElement(pPr0, f'{{{W_NS}}}jc')
        jc0.set(f'{{{W_NS}}}val', 'center')

    # ── FIX 2: Ampliar text box columna derecha (Fecha y hora ingreso cortada) ──
    # Shape3 (columna derecha) mide 6.63cm → texto se corta
    # La ampliamos a 7.8cm desde el código
    TARGET_CX_EMU = str(int(7.8 / 2.54 * 914400))  # 7.8cm en EMU
    import lxml.etree as _etree2
    body_xml = _etree2.tostring(doc.element.body, encoding='unicode')
    # Count wsp shapes and find Shape3 (index 2, 0-based among wps:wsp)
    # Shape3 is the right column with Edad/FechaNac/FechaIngreso
    shapes_found = 0
    for elem in doc.element.body.iter('{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}wsp'):
        ext_elem = elem.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
        if ext_elem is not None:
            cx_val = ext_elem.get('cx', '0')
            cx_cm = int(cx_val) / 914400 * 2.54
            # Shape3 is ~6.63cm (right column) - identify by size
            if 6.0 < cx_cm < 7.0:
                ext_elem.set('cx', TARGET_CX_EMU)
                shapes_found += 1

    REMOVE_TAGS = {
        f'{{{W_NS}}}r', f'{{{W_NS}}}proofErr',
        f'{{{W_NS}}}bookmarkStart', f'{{{W_NS}}}bookmarkEnd',
    }

    # ── Firma: normalizar tamaño y posicionar centrada en zona del Dr. ──
    if ruta_png_firma and ruta_png_firma.exists():
        from docx.shared import Twips, Pt
        from docx.oxml.ns import qn as _qn
        import lxml.etree as _etree
        from PIL import Image as _PILImg

        # Tamaño: ancho base 5cm, alto máximo 1.8cm (igual al PDF de referencia)
        # MIN 4cm para que firmas cuadradas no queden muy angostas
        ANCHO_BASE = 5.0
        ALTO_MAX   = 1.8
        MIN_ANCHO  = 4.0
        try:
            img_tmp = _PILImg.open(str(ruta_png_firma))
            w_px, h_px = img_tmp.size
            img_tmp.close()
            ratio    = w_px / h_px if h_px > 0 else 2.0
            ancho_cm = ANCHO_BASE
            alto_cm  = ancho_cm / ratio
            if alto_cm > ALTO_MAX:
                alto_cm  = ALTO_MAX
                ancho_cm = max(MIN_ANCHO, alto_cm * ratio)
            firma_width = Cm(ancho_cm)
        except Exception:
            ancho_cm    = 5.0
            firma_width = Cm(5.0)

        # ── Eliminar space_before exagerado de párrafos vacíos alrededor de la firma ──
        # Párrafos 13 (space=129 twips) y 15 (space=104 twips) crean el hueco grande
        # ── Ajustar spacing de TODOS los párrafos vacíos del documento ──────
        # El FORMATO.docx original tiene:
        # - P15-P22: 8 empty paras (8 líneas de espacio innecesario)
        # - P23: before=202 twips (enorme)
        # - P25: before=129 twips (antes de firma)
        # - P27: before=104 twips (después de firma)
        # Los corregimos desde código sin tocar el DOCX
        for para in doc.paragraphs:
            if para.text.strip():
                continue  # solo párrafos vacíos
            pPr = para._p.find(f'{{{W_NS}}}pPr')
            if pPr is None:
                continue
            sp = pPr.find(f'{{{W_NS}}}spacing')
            if sp is None:
                continue
            before_val = int(sp.get(f'{{{W_NS}}}before', '0'))
            if before_val > 100:
                # Reducir valores grandes a moderados
                sp.set(f'{{{W_NS}}}before', '20')

        # Ajuste fino alrededor de la firma
        firma_idx = next((i for i, p in enumerate(doc.paragraphs)
                          if '{FirmaMedico}' in p.text), None)
        if firma_idx is not None:
            spacing_targets = {
                -9: 0,   # P15-P22 vacíos: zerear (8 empty paras)
                -8: 0,
                -7: 0,
                -6: 0,
                -5: 0,
                -4: 0,
                -3: 0,
                -2: 0,
                -1: 60,  # P25 (justo antes firma): 60 twips
                 1: 40,  # P27 (justo después firma): 40 twips
                 2: 0,
            }
            for offset, target_twips in spacing_targets.items():
                idx = firma_idx + offset
                if 0 <= idx < len(doc.paragraphs):
                    para_v = doc.paragraphs[idx]
                    if not para_v.text.strip():
                        pPr = para_v._p.find(f'{{{W_NS}}}pPr')
                        if pPr is not None:
                            sp = pPr.find(f'{{{W_NS}}}spacing')
                            if sp is not None:
                                if target_twips == 0:
                                    sp.attrib.pop(f'{{{W_NS}}}before', None)
                                    sp.attrib.pop(f'{{{W_NS}}}after', None)
                                else:
                                    sp.set(f'{{{W_NS}}}before', str(target_twips))

        for i, para in enumerate(doc.paragraphs):
            if '{FirmaMedico}' in para.text:
                p_elem = para._p
                for child in list(p_elem):
                    if child.tag != f'{{{W_NS}}}pPr':
                        p_elem.remove(child)
                # Centrar la firma en la zona del Dr.
                # Zona: empieza en 10.73cm, ancho 8.57cm → centro 15.015cm
                ZONA_CENTRO_CM = 15.015
                indent_cm = ZONA_CENTRO_CM - (ancho_cm / 2)
                indent_twips = int(indent_cm / 2.54 * 1440)
                para.paragraph_format.left_indent = Twips(indent_twips)
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after  = Pt(0)
                run = para.add_run()
                run.add_picture(str(ruta_png_firma), width=firma_width)
                log(f"  → Firma insertada en documento OK", "info")
                break

        # ── Borde superior al párrafo Dr. (línea horizontal) ──
        for para in doc.paragraphs:
            if '{Medico}' in para.text:
                pPr = para._p.find(f'{{{W_NS}}}pPr')
                if pPr is None:
                    pPr = _etree.SubElement(para._p, f'{{{W_NS}}}pPr')
                    para._p.insert(0, pPr)
                pBdr = pPr.find(f'{{{W_NS}}}pBdr')
                if pBdr is None:
                    pBdr = _etree.SubElement(pPr, f'{{{W_NS}}}pBdr')
                for old in pBdr.findall(f'{{{W_NS}}}top'):
                    pBdr.remove(old)
                top = _etree.SubElement(pBdr, f'{{{W_NS}}}top')
                top.set(f'{{{W_NS}}}val',   'single')
                top.set(f'{{{W_NS}}}sz',    '6')
                top.set(f'{{{W_NS}}}space', '8')
                top.set(f'{{{W_NS}}}color', '000000')
                break


    def reemplazar_p_elem(p_elem, texto_override=None):
        """Reemplaza variables preservando rPr original."""
        texts = [t.text or '' for t in p_elem.iter(f'{{{W_NS}}}t')]
        texto = texto_override or ''.join(texts)

        if '{FirmaMedico}' in texto:
            return  # ya manejado arriba
        if not any(k in texto for k in reemplazos):
            return

        primer_rpr = None
        for child in p_elem:
            if child.tag == f'{{{W_NS}}}r':
                rpr = child.find(f'{{{W_NS}}}rPr')
                if rpr is not None:
                    primer_rpr = copy.deepcopy(rpr)
                break

        nuevo = texto
        for k, v in reemplazos.items():
            nuevo = nuevo.replace(k, v)

        for child in list(p_elem):
            if child.tag in REMOVE_TAGS:
                p_elem.remove(child)

        nuevo_r = etree.SubElement(p_elem, f'{{{W_NS}}}r')
        if primer_rpr is not None:
            nuevo_r.insert(0, primer_rpr)
        nuevo_t = etree.SubElement(nuevo_r, f'{{{W_NS}}}t')
        nuevo_t.text = nuevo
        nuevo_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    # ── 1. Text boxes ──
    for p_elem in doc.element.body.iter(f'{{{W_NS}}}p'):
        parent = p_elem.getparent()
        in_txbx = False
        while parent is not None:
            if parent.tag.split('}')[-1] == 'txbxContent':
                in_txbx = True
                break
            parent = parent.getparent()
        if in_txbx:
            reemplazar_p_elem(p_elem)

    # ── 2. Cuerpo principal ──
    for para in doc.paragraphs:
        reemplazar_p_elem(para._p, texto_override=para.text)

    # ── 3. Tablas ──
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for para in celda.paragraphs:
                    reemplazar_p_elem(para._p, texto_override=para.text)

    doc.save(str(ruta_docx))
    return ruta_docx


# ─────────────────────────────────────────────────────────────
# CONVERSIÓN DOCX → PDF (multiplataforma)
# ─────────────────────────────────────────────────────────────
def docx_a_pdf(ruta_docx: Path, session_dir: Path) -> Path:
    """
    Convierte docx a PDF.
    - En Windows: usa docx2pdf (requiere Word instalado)
    - En Linux/Railway: usa LibreOffice CLI
    """
    import platform

    # Intentar primero con docx2pdf (solo funciona en Windows con Word)
    if platform.system() == "Windows":
        try:
            from docx2pdf import convert
            pdf_path = session_dir / (ruta_docx.stem + ".pdf")
            convert(str(ruta_docx), str(pdf_path))
            if not pdf_path.exists():
                raise Exception("PDF no generado")
            return pdf_path
        except ImportError:
            raise Exception("docx2pdf no instalado. Ejecuta: pip install docx2pdf")
        except Exception as e:
            raise Exception(f"Error convirtiendo a PDF con docx2pdf: {e}")

    # En Linux/Railway, usar LibreOffice
    try:
        # Importar función del módulo de conversión
        import convertidor_pdf as cpdf
        return cpdf.docx_a_pdf_libreoffice(ruta_docx, session_dir)
    except Exception as e:
        raise Exception(f"Error convirtiendo a PDF (LibreOffice): {e}")

# ─────────────────────────────────────────────────────────────
# REPORTE EXCEL
# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# PROCESO PRINCIPAL (hilo)
# ─────────────────────────────────────────────────────────────
def ejecutar_proceso(casos, config, session_dir, completadas_previas):
    zipf = None
    try:
        total = len(casos)
        completadas_set = set(completadas_previas)
        exitosos = []
        errores  = []

        zip_path = session_dir / f"certificados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zipf = zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED)

        with LK:
            S["session_dir"] = session_dir
            S["stats"]["total"] = total

        pendientes = [c for c in casos if str(c) not in completadas_set]
        ya_hechos  = total - len(pendientes)

        if ya_hechos > 0:
            log(f"Reanudando sesión — {ya_hechos} ya completados, {len(pendientes)} pendientes", "warn")
            with LK:
                S["stats"]["ok"] = ya_hechos

        for caso in pendientes:
            if S.get("stopping", False):
                raise StopBotException("Detenido por usuario")

            # Soportar tanto dict {NoCuenta, NoFactura} como string legacy
            if isinstance(caso, dict):
                no_cuenta  = str(caso.get("NoCuenta", "")).strip()
                no_factura = str(caso.get("NoFactura", "")).strip()
            else:
                no_cuenta  = str(caso).strip()
                no_factura = no_cuenta

            # Si no tiene factura, usar el número de cuenta como nombre
            nombre_pdf = no_factura if no_factura and no_factura != "None" else no_cuenta

            log(f"Procesando — Caso: {no_cuenta} | Factura: {nombre_pdf}", "info")

            success = False
            motivo_error = ""

            for intento in range(2):
                if S.get("stopping", False):
                    raise StopBotException("Detenido por usuario")
                try:
                    # 1. Llamar API
                    log(f"  → Consultando API...", "info")
                    datos = llamar_api(config, no_cuenta)

                    if not datos:
                        raise Exception("La API no retornó datos para este caso")

                    log(f"  → Paciente: {datos.get('Paciente','?')}", "info")
                    log(f"  → Medico: {datos.get('Medico','(vacío)')}", "info")

                    # 2. Convertir firma (Base64 → PNG)
                    ruta_firma_png = None
                    firma_b64 = datos.get("FirmaMedico", "")
                    cedula     = datos.get("Cedula", no_cuenta)

                    if firma_b64 and PIL_OK:
                        nombre_archivo_firma = f"firma_{cedula}.png"
                        ruta_firma = session_dir / nombre_archivo_firma

                        # Diagnóstico: longitud del base64
                        log(f"  → Firma base64 recibida: {len(firma_b64):,} caracteres", "info")

                        imagen_procesada = procesar_firma_base64(firma_b64, str(ruta_firma))
                        if imagen_procesada:
                            ruta_firma_png = ruta_firma
                            # Verificar que el PNG existe y tiene contenido real
                            from PIL import Image as PILImg
                            try:
                                img_check = PILImg.open(str(ruta_firma_png))
                                log(f"  → Firma PNG: {img_check.size[0]}x{img_check.size[1]}px modo={img_check.mode}", "info")
                            except Exception as ec:
                                log(f"  → Firma PNG error verificación: {ec}", "warn")
                                ruta_firma_png = None
                        else:
                            log(f"  → Firma falló al procesar, se omite", "warn")
                    elif firma_b64 and not PIL_OK:
                        log(f"  → Pillow no instalado — instala con: pip install Pillow", "warn")
                    else:
                        log(f"  → Campo FirmaMedico vacío en respuesta API", "warn")

                    # 3. Generar PDF directamente con reportlab
                    log(f"  → Generando PDF...", "info")
                    ruta_pdf_raw = generar_pdf_directo(datos, no_cuenta, ruta_firma_png, session_dir)

                    # Liberar PNG de firma
                    if ruta_firma_png and ruta_firma_png.exists():
                        try:
                            ruta_firma_png.unlink()
                        except Exception:
                            pass
                        ruta_firma_png = None

                    # 5. Renombrar PDF como CM_{NoFactura}
                    pdf_final = session_dir / f"CM_{nombre_pdf}.pdf"
                    if ruta_pdf_raw != pdf_final:
                        if pdf_final.exists():
                            pdf_final.unlink()
                        ruta_pdf_raw.rename(pdf_final)

                    # 6. Agregar al ZIP
                    zipf.write(pdf_final, arcname=f"CM_{nombre_pdf}.pdf")

                    success = True
                    log(f"  ✓ CM_{nombre_pdf}.pdf generado", "ok")
                    break

                except StopBotException:
                    raise
                except Exception as e:
                    motivo_error = str(e)
                    log(f"  Error intento {intento+1}/2: {e}", "error")
                    if intento == 0:
                        log(f"  Reintentando...", "warn")
                        time.sleep(2)

            if success:
                with LK:
                    S["stats"]["ok"] += 1
                    S["exitosos"].append(no_cuenta)
                    S["completadas"].add(no_cuenta)
                completadas_set.add(no_cuenta)
                exitosos.append((no_cuenta, nombre_pdf))
                guardar_progreso(session_dir, completadas_set)
            else:
                with LK:
                    S["stats"]["err"] += 1
                    S["errores"].append((no_cuenta, motivo_error))
                errores.append((no_cuenta, motivo_error))
                log(f"  ✗ Caso {no_cuenta} / Factura {nombre_pdf} fallido: {motivo_error}", "error")

            time.sleep(0.5)  # Pausa entre casos para no saturar la API

        # Reporte final
        reporte = generar_excel_reporte(session_dir, exitosos, errores)
        if reporte and reporte.exists():
            zipf.write(reporte, arcname="reporte_resultados.xlsx")

        zipf.close()
        with LK:
            S["zip_path"] = zip_path.name

        if not S["stopping"] and len(completadas_set) >= total:
            marcar_sesion_finalizada(session_dir)
            log("=" * 50, "info")
            log(f"✅ Proceso completado — Total:{total} OK:{len(exitosos)} Errores:{len(errores)}", "ok")
        else:
            log("Sesión guardada para reanudar", "warn")

        if errores:
            log("Casos con error:", "warn")
            for nc, mot in errores:
                log(f"  - {nc}: {mot}", "error")

    except StopBotException:
        log("Bot detenido por usuario", "warn")
    except Exception as e:
        log(f"Error crítico: {e}", "error")
        with LK:
            S["error"] = str(e)
    finally:
        if zipf:
            try:
                zipf.close()
            except Exception:
                pass
        with LK:
            S["running"]  = False
            S["finished"] = True


def generar_excel_reporte(session_dir, exitosos, errores) -> Path:
    """exitosos = [(no_cuenta, no_factura), ...] | errores = [(no_cuenta, motivo), ...]"""
    if not EXCEL_OK:
        return None
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Resultados"
        headers = ["NoCuenta", "NoFactura", "Archivo PDF", "Estado", "Motivo"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="1a3a52", end_color="1a3a52", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        row = 2
        for nc, nf in exitosos:
            ws.cell(row=row, column=1, value=nc)
            ws.cell(row=row, column=2, value=nf)
            ws.cell(row=row, column=3, value=f"CM_{nf}.pdf")
            c = ws.cell(row=row, column=4, value="OK")
            c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            ws.cell(row=row, column=5, value="")
            row += 1
        for nc, motivo in errores:
            ws.cell(row=row, column=1, value=nc)
            ws.cell(row=row, column=2, value="")
            ws.cell(row=row, column=3, value="")
            c = ws.cell(row=row, column=4, value="ERROR")
            c.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            ws.cell(row=row, column=5, value=motivo)
            row += 1
        for col_letter, w in zip(["A","B","C","D","E"], [18,18,22,10,50]):
            ws.column_dimensions[col_letter].width = w
        path = session_dir / "reporte_resultados.xlsx"
        wb.save(path)
        return path
    except Exception as e:
        log(f"Warning reporte: {e}", "warn")
        return None
        if reporte and reporte.exists():
            zipf.write(reporte, arcname="reporte_resultados.xlsx")

        zipf.close()
        with LK:
            S["zip_path"] = zip_path.name

        if not S["stopping"] and len(completadas_set) >= total:
            marcar_sesion_finalizada(session_dir)
            log("=" * 50, "info")
            log(f"✅ Proceso completado — Total:{total} OK:{len(exitosos)} Errores:{len(errores)}", "ok")
        else:
            log("Sesión guardada para reanudar", "warn")

        if errores:
            log("Casos con error:", "warn")
            for nc, mot in errores:
                log(f"  - {nc}: {mot}", "error")


# ─────────────────────────────────────────────────────────────
# RUTAS FLASK
# ─────────────────────────────────────────────────────────────
@app.route("/login", methods=["GET"])
def login_page():
    if session.get("autenticado"):
        return redirect(url_for("index"))
    return render_template("login.html")

@app.route("/api/auth/login", methods=["POST"])
def auth_login():
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "").split(",")[0].strip()

    # Verificar rate limit
    permitido, msg_bloqueo, _ = _check_rate_limit(ip)
    if not permitido:
        return jsonify({"error": msg_bloqueo}), 429

    data     = request.get_json() or {}
    usuario  = data.get("usuario", "").strip()
    password = data.get("password", "").strip()
    if not usuario or not password:
        return jsonify({"error": "Usuario y contraseña requeridos"}), 400

    users = get_users()
    if usuario in users and users[usuario] == password:
        _registrar_login_exitoso(ip)
        session["autenticado"] = True
        session["usuario"]     = usuario
        session.permanent      = True
        log(f"Login exitoso: {usuario} desde {ip}", "info")
        return jsonify({"success": True, "usuario": usuario})

    # Login fallido
    _registrar_intento_fallido(ip)
    datos_ip = _login_intentos.get(ip, {})
    intentos  = datos_ip.get("intentos", 1)
    restantes = max(0, _LOGIN_MAX_INTENTOS - intentos)
    log(f"Login fallido: usuario='{usuario}' IP={ip} (intento {intentos}/{_LOGIN_MAX_INTENTOS})", "warn")

    if restantes > 0:
        return jsonify({"error": f"Usuario o contraseña incorrectos. Te quedan {restantes} intento(s)."}), 401
    else:
        return jsonify({"error": f"Demasiados intentos fallidos. IP bloqueada por {_LOGIN_BLOQUEO_SEG//60} minutos."}), 429

@app.route("/api/auth/logout", methods=["POST"])
def auth_logout():
    session.clear()
    return jsonify({"success": True})

@app.route("/api/auth/me")
def auth_me():
    if session.get("autenticado"):
        return jsonify({"autenticado": True, "usuario": session.get("usuario", "")})
    return jsonify({"autenticado": False}), 401

@app.route("/")
@login_requerido
def index():
    return render_template("index.html")

@app.route("/api/servidores")
@login_requerido
def get_servidores():
    """Devuelve lista de servidores desde la API Esculapio."""
    try:
        servidores = _get_servidores_api()
        # Transformar al formato que espera el frontend: {Descripcion: {id, ip, puerto}}
        data = {}
        for srv in servidores:
            nombre = srv.get("Descripcion", "").strip()
            if nombre:
                data[nombre] = {
                    "id":     srv.get("$id", ""),
                    "ip":     srv.get("Ip_Conexion", ""),
                    "puerto": str(srv.get("Puerto", "3306")),
                    "usuario_default": srv.get("Usuario", ""),
                }
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/login", methods=["POST"])
@login_requerido
def login_usuario():
    """
    Valida credenciales contra la API y devuelve las empresas disponibles.
    Body: {servidor: 'Campbell', usuario: '...', password: '...'}
    """
    try:
        data     = request.get_json()
        nombre   = data.get("servidor", "")
        usuario  = data.get("usuario", "")
        password = data.get("password", "")

        if not all([nombre, usuario, password]):
            return jsonify({"error": "Faltan datos"}), 400

        # Buscar el servidor en la API
        servidores = _get_servidores_api()
        srv = next((s for s in servidores if s.get("Descripcion", "").strip() == nombre), None)
        if not srv:
            return jsonify({"error": f"Servidor '{nombre}' no encontrado"}), 404

        ip     = srv.get("Ip_Conexion", "")
        puerto = str(srv.get("Puerto", "3306"))

        # Validar credenciales y obtener empresas
        empresas = _get_empresas_api(ip, puerto, usuario, password)
        if not empresas:
            return jsonify({"error": "Credenciales inválidas o usuario sin empresas asignadas en este servidor"}), 401

        # Formatear empresas para el frontend
        lista = [{"cod": e.get("Empresa", ""), "nombre": e.get("Nombre_Empresa", "")}
                 for e in empresas if e.get("Empresa")]

        if not lista:
            return jsonify({"error": f"No se encontraron empresas (campos recibidos: {list(empresas[0].keys()) if empresas else []})"}), 401

        return jsonify({
            "success":  True,
            "empresas": lista,
            "ip":       ip,
            "puerto":   puerto,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/generar-plantilla")
@login_requerido
def gen_plantilla():
    if not EXCEL_OK:
        return jsonify({"error": "openpyxl no instalado"}), 400
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Casos"
    ws.append(["NoCuenta", "NoFactura"])
    ws.append(["715874", "FAC-001"])
    ws.append(["897544", "FAC-002"])
    for cell in ws[1]:
        cell.font      = Font(bold=True, color="FFFFFF", size=12)
        cell.fill      = PatternFill(start_color="1a3a52", end_color="1a3a52", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 20
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name="plantilla_casos.xlsx")

@app.route("/api/cargar-datos", methods=["POST"])
@login_requerido
def cargar_datos():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No se envió archivo"}), 400
        f = request.files["file"]
        if not f.filename:
            return jsonify({"error": "Archivo vacío"}), 400
        datos = []
        if f.filename.endswith(".csv"):
            raw = f.stream.read().decode("utf-8")
            reader = csv.DictReader(raw.splitlines())
            for row in reader:
                nc  = row.get("NoCuenta")  or row.get("nocuenta")  or row.get("NOCUENTA")
                nf  = row.get("NoFactura") or row.get("nofactura") or row.get("NOFACTURA") or row.get("Factura") or ""
                if nc and str(nc).strip():
                    datos.append({"NoCuenta": str(nc).strip(), "NoFactura": str(nf).strip()})
        elif f.filename.endswith(".xlsx"):
            if not EXCEL_OK:
                return jsonify({"error": "openpyxl no instalado"}), 500
            wb = openpyxl.load_workbook(f)
            ws = wb.active
            headers = [str(cell.value).strip().lower() if cell.value else "" for cell in ws[1]]

            def find_col(names):
                for name in names:
                    for i, h in enumerate(headers):
                        if h == name:
                            return i
                return None

            col_nc = find_col(["nocuenta", "no_cuenta", "no cuenta", "cuenta", "caso"])
            col_nf = find_col(["nofactura", "no_factura", "no factura", "factura", "nfactura"])

            if col_nc is None:
                # Intentar primera columna
                col_nc = 0

            for row in ws.iter_rows(min_row=2, values_only=True):
                nc = str(row[col_nc]).strip() if col_nc is not None and col_nc < len(row) and row[col_nc] is not None else ""
                nf = str(row[col_nf]).strip() if col_nf is not None and col_nf < len(row) and row[col_nf] is not None else ""
                if nc and nc != "None":
                    datos.append({"NoCuenta": nc, "NoFactura": nf})
        else:
            return jsonify({"error": "Use .xlsx o .csv"}), 400
        return jsonify({"success": True, "count": len(datos), "casos": datos})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/iniciar", methods=["POST"])
@login_requerido
def iniciar():
    with LK:
        if S["running"]:
            return jsonify({"error": "Proceso en curso"}), 400
    reset()

    data = request.get_json()
    casos    = data.get("casos", [])
    servidor = data.get("servidor", "")
    empresa  = data.get("empresa", "")
    usuario  = data.get("usuario", "")
    password = data.get("password", "")
    reanudar = data.get("reanudar", False)

    if not casos:
        return jsonify({"error": "No hay casos para procesar"}), 400
    if not servidor:
        return jsonify({"error": "Servidor no seleccionado"}), 400
    if not empresa:
        return jsonify({"error": "Empresa no seleccionada"}), 400
    if not usuario or not password:
        return jsonify({"error": "Credenciales incompletas"}), 400

    # Buscar servidor en la API
    servidores = _get_servidores_api()
    srv_data   = next((s for s in servidores if s.get("Descripcion", "").strip() == servidor), None)
    if not srv_data:
        return jsonify({"error": f"Servidor '{servidor}' no encontrado en la API"}), 400

    config = {
        "servidor":    servidor,
        "ip":          srv_data.get("Ip_Conexion", ""),
        "puerto":      str(srv_data.get("Puerto", "3306")),
        "cod_empresa": empresa,
        "usuario":     usuario,
        "password":    password,
    }

    # Buscar sesión pendiente o crear nueva
    if reanudar:
        session_dir = buscar_sesion_pendiente()
        if session_dir:
            casos_sesion, config_sesion = cargar_casos_sesion(session_dir)
            if casos_sesion:
                casos = casos_sesion
            completadas = cargar_progreso(session_dir)
        else:
            session_dir  = DOWNLOAD_DIR / f"sesion_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            session_dir.mkdir(exist_ok=True)
            completadas  = set()
            guardar_casos_sesion(session_dir, casos, config)
    else:
        session_dir = DOWNLOAD_DIR / f"sesion_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        session_dir.mkdir(exist_ok=True)
        completadas = set()
        guardar_casos_sesion(session_dir, casos, config)

    with LK:
        S["running"]  = True
        S["stopping"] = False
        S["stats"]["total"] = len(casos)

    threading.Thread(
        target=ejecutar_proceso,
        args=(casos, config, session_dir, completadas),
        daemon=True
    ).start()

    return jsonify({"success": True, "total": len(casos)})

@app.route("/api/detener", methods=["POST"])
@login_requerido
def detener():
    with LK:
        S["stopping"] = True
    log("Señal de parada enviada", "warn")
    return jsonify({"success": True})

@app.route("/api/estado")
@login_requerido
def estado():
    offset = int(request.args.get("offset", 0))
    with LK:
        nuevos = S["logs"][offset:]
        return jsonify({
            "running":     S["running"],
            "finished":    S["finished"],
            "logs":        nuevos,
            "log_offset":  offset + len(nuevos),
            "stats":       S["stats"],
            "error":       S["error"],
            "zip_path":    S["zip_path"],
            "completadas": list(S["completadas"]),
        })

@app.route("/api/sesion-pendiente")
@login_requerido
def sesion_pendiente_info():
    try:
        sesion = buscar_sesion_pendiente()
        if not sesion:
            return jsonify({"existe": False})
        completadas = cargar_progreso(sesion)
        casos, config = cargar_casos_sesion(sesion)
        return jsonify({
            "existe":      True,
            "nombre":      sesion.name,
            "completadas": len(completadas),
            "total":       len(casos),
            "pendientes":  len(casos) - len(completadas),
            "servidor":    config.get("servidor", "?"),
            "empresa":     config.get("cod_empresa", "?"),
        })
    except Exception as e:
        return jsonify({"existe": False, "error": str(e)})

@app.route("/api/limpiar-progreso", methods=["POST"])
@login_requerido
def limpiar_progreso():
    try:
        eliminadas = []
        for d in DOWNLOAD_DIR.iterdir():
            if d.is_dir() and d.name.startswith("sesion_"):
                if not (d / "finalizado.txt").exists():
                    shutil.rmtree(d)
                    eliminadas.append(d.name)
        reset()
        return jsonify({"success": True, "eliminadas": eliminadas})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/descargar-zip")
@login_requerido
def descargar_zip():
    with LK:
        zp          = S["zip_path"]
        session_dir = S["session_dir"]
    if not zp:
        return jsonify({"error": "ZIP no disponible"}), 404
    zip_file = (session_dir / zp) if session_dir else (DOWNLOAD_DIR / zp)
    if not zip_file.exists():
        return jsonify({"error": "ZIP no encontrado"}), 404
    return send_file(zip_file, mimetype="application/zip",
                     as_attachment=True, download_name=zp)

if __name__ == "__main__":
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    port = int(os.environ.get("PORT", 5050))
    app.run(debug=False, host="0.0.0.0", port=port)