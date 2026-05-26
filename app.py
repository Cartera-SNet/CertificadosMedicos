"""
Bot Certificaciones Médicas SOAT — v2.0 (Sin autenticación)
Generación masiva de certificados médicos con Flask.
"""

import os, csv, time, threading, zipfile, logging, json, shutil, re, base64, urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
from datetime import datetime, timedelta
from pathlib import Path
from io import BytesIO
from copy import deepcopy
from concurrent.futures import ThreadPoolExecutor, as_completed

from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
from functools import wraps

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

# ReportLab para PDF directo
try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ─────────────────────────────────────────────────────────────
# APIs Esculapio
# ─────────────────────────────────────────────────────────────
API_SERVIDORES = "https://appsintranet.esculapiosis.com/ApiCampbell/api/Usuarios/obtener-servidores"
API_EMPRESAS   = "https://appsintranet.esculapiosis.com/ApiCampbell/api/Usuarios/obtener-empresas"
API_BASE       = "https://appsintranet.esculapiosis.com/ApiCampbell/api/Reportes/obtener-certificacionmedica"

_cache_servidores = None
_cache_servidores_ts = 0

def _get_servidores_api():
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
    except Exception as e:
        log(f"API empresas error: {e}", "warn")
        return []

# ─────────────────────────────────────────────────────────────
# FLASK APP (sin autenticación)
# ─────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(message)s")
L = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "sis-soat-2025-key")
from datetime import timedelta
app.permanent_session_lifetime = timedelta(hours=8)

# ── HTTPS forzado en Railway ────────────────────────────────
@app.before_request
def forzar_https():
    proto = request.headers.get("X-Forwarded-Proto", "")
    if proto == "http":
        return redirect(request.url.replace("http://", "https://", 1), code=301)

# ── Rate limiting ───────────────────────────────────────────
_login_intentos = {}
_LOGIN_MAX = 5; _LOGIN_VENTANA = 600; _LOGIN_BLOQUEO = 900

def _check_rate(ip):
    ahora = time.time()
    d = _login_intentos.get(ip, {})
    if d.get("bloqueado_hasta", 0) > ahora:
        r = int(d["bloqueado_hasta"] - ahora)
        return False, f"IP bloqueada. Intenta en {r//60}m {r%60}s."
    if ahora - d.get("primer_intento", 0) > _LOGIN_VENTANA:
        _login_intentos.pop(ip, None)
    return True, ""

def _fail(ip):
    ahora = time.time()
    d = _login_intentos.get(ip, {})
    if not d or ahora - d.get("primer_intento", 0) > _LOGIN_VENTANA:
        d = {"intentos": 0, "primer_intento": ahora}
    d["intentos"] += 1
    if d["intentos"] >= _LOGIN_MAX:
        d["bloqueado_hasta"] = ahora + _LOGIN_BLOQUEO
    _login_intentos[ip] = d

def _ok(ip): _login_intentos.pop(ip, None)

def login_requerido(f):
    @wraps(f)
    def dec(*a, **kw):
        if not session.get("autenticado"):
            if request.is_json or request.path.startswith("/api/"):
                return jsonify({"error": "No autorizado", "login_required": True}), 401
            return redirect(url_for("login_page"))
        return f(*a, **kw)
    return dec
BASE_DIR = Path(__file__).parent
FONTS_DIR = BASE_DIR / "fonts"
FONTS_DIR.mkdir(exist_ok=True)
DOWNLOAD_DIR = BASE_DIR / "downloads"
DOWNLOAD_DIR.mkdir(exist_ok=True)

# Descargar fuentes (opcional, para ReportLab)
def _descargar_fuentes():
    import urllib.request
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
            continue
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=20) as r:
                dest.write_bytes(r.read())
        except Exception:
            pass

if REPORTLAB_OK:
    try:
        _descargar_fuentes()
    except:
        pass

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

# ── Persistencia y funciones auxiliares (igual que antes) ──
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

# Formateo de fechas
MESES_ES = {
    1:"ene", 2:"feb", 3:"mar", 4:"abr", 5:"may", 6:"jun",
    7:"jul", 8:"ago", 9:"sep", 10:"oct", 11:"nov", 12:"dic"
}

def formatear_fecha_nac(fecha_str):
    try:
        dt = datetime.fromisoformat(fecha_str.replace("T", " ").split(" ")[0])
        mes = MESES_ES[dt.month]
        return f"{dt.day:02d}-{mes}.-{dt.year}"
    except Exception:
        return fecha_str

# Procesamiento de firma
def procesar_firma_base64(base64_string, ruta_salida=None):
    try:
        if not base64_string or len(base64_string) < 100:
            log(f"  → Firma demasiado corta ({len(base64_string)} chars), ignorada", "warn")
            return None
        missing_padding = len(base64_string) % 4
        if missing_padding:
            base64_string += '=' * (4 - missing_padding)
        image_data = base64.b64decode(base64_string)
        img = Image.open(BytesIO(image_data))
        img = img.convert("RGBA")
        data = img.getdata()
        new_data = []
        for r, g, b, a in data:
            if r > 240 and g > 240 and b > 240:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append((r, g, b, a))
        img.putdata(new_data)
        try:
            alpha = img.split()[3]
            bbox = alpha.getbbox()
            if bbox:
                PAD = 8
                left   = max(0, bbox[0] - PAD)
                top    = max(0, bbox[1] - PAD)
                right  = min(img.width,  bbox[2] + PAD)
                bottom = min(img.height, bbox[3] + PAD)
                img = img.crop((left, top, right, bottom))
        except Exception:
            pass
        if ruta_salida:
            img.save(str(ruta_salida), format="PNG")
            return ruta_salida
        return img
    except Exception as e:
        log(f"Error al procesar la firma: {e}", "warn")
        return None

# ------------------------------------------------------------
# GENERACIÓN DE PDF CON REPORTLAB (RÁPIDO)
# ------------------------------------------------------------
def generar_pdf_directo(datos_api, no_cuenta, ruta_png_firma, session_dir) -> Path:
    if not REPORTLAB_OK:
        raise Exception("ReportLab no instalado. Ejecuta: pip install reportlab")

    W, H = LETTER  # 612 x 792

    # Coordenadas exactas medidas del MUESTRA.pdf con pdfminer
    TITULO1_Y   = 24.21
    TITULO2_Y   = 38.73
    CUADRO_Y    = 54.72
    CUADRO_H    = 118.2
    INTRO_Y     = 66.07
    PACIENTE_Y  = 78.07
    ID_Y        = 117.91
    SEXO_Y      = 136.63
    SERVICIO_Y  = 156.07
    EDAD_Y      = 116.35
    FECHA_NAC_Y = 136.63
    FECHA_ING_Y = 156.07
    SUSCRITO1_Y = 183.63
    SUSCRITO2_Y = 195.63
    FIRMA_X     = 322.2
    FIRMA_XR    = 565.2
    FIRMA_ZW    = FIRMA_XR - FIRMA_X
    LINEA_Y     = 282.24
    MEDICO_Y    = 293.07
    REG_Y       = 305.11

    # Registrar fuentes (Carlito/Caladea)
    def _reg(name, path):
        try:
            pdfmetrics.registerFont(TTFont(name, str(path)))
            return True
        except:
            return False

    if _reg("F1", FONTS_DIR/"Carlito-Bold.ttf") and _reg("F2", FONTS_DIR/"Carlito-Regular.ttf"):
        F1 = "F1"; F2 = "F2"
    else:
        F1 = "Helvetica-Bold"; F2 = "Helvetica"

    def sf(text, bold=False):
        return F1 if bold else F2

    nombre_empresa = str(datos_api.get("nombre_empresa", ""))
    paciente = str(datos_api.get("Paciente", ""))
    tipo_documento = str(datos_api.get("TipoDocumento", ""))
    sexo = str(datos_api.get("sexo", ""))
    cod_servicio = str(datos_api.get("CodServicio", ""))
    nom_servicio = str(datos_api.get("NomServicio", ""))
    edad = str(int(float(datos_api.get("Edad", 0))))
    medida_edad = str(datos_api.get("Medida_Edad", ""))
    fecha_nac = formatear_fecha_nac(str(datos_api.get("Fecha_Nac", "")))
    fecha_ingreso = str(datos_api.get("Fecha_Ingreso", ""))
    medico = str(datos_api.get("Medico", ""))
    tarjeta_profesional = str(datos_api.get("Tarjeta_Profesional", ""))

    pdf_path = session_dir / f"CM_{no_cuenta}.pdf"
    c = canvas.Canvas(str(pdf_path), pagesize=LETTER)
    def yt(y_top): return H - y_top

    # Títulos
    c.setFont(sf("CERTIFICADO DE ATENCION MEDICA PARA VICTIMAS DE", bold=True), 13)
    c.drawCentredString(W/2, yt(TITULO1_Y - 3), "CERTIFICADO DE ATENCION MEDICA PARA VICTIMAS DE")
    c.setFont(sf("ACCIDENTES DE TRÁNSITO", bold=True), 13)
    c.drawCentredString(W/2, yt(TITULO2_Y - 3), "ACCIDENTES DE TRÁNSITO")

    # Cuadro
    c.setLineWidth(0.8)
    c.roundRect(27.84, yt(CUADRO_Y + CUADRO_H), 553.2, CUADRO_H, 4, stroke=1, fill=0)

    # Texto interno (intro y paciente)
    c.setFont(F2, 9.941)
    intro = f"El suscrito médico de la Institución Prestadora de Servicios  {nombre_empresa} certifica que atendió a:"
    MAX_W = 541.99 - 31.92
    if c.stringWidth(intro, F2, 9.941) <= MAX_W:
        c.setFont(sf(intro), 9.941)
        c.drawString(31.92, yt(INTRO_Y), intro)
        c.setFont(sf(f"Señor(a)   {paciente}"), 9.941)
        c.drawString(31.92, yt(PACIENTE_Y), f"Señor(a)   {paciente}")
    else:
        words = intro.split(); lines = []; line = ""
        for word in words:
            test = (line + " " + word).strip()
            if c.stringWidth(test, F2, 9.941) <= MAX_W:
                line = test
            else:
                if line: lines.append(line)
                line = word
        if line: lines.append(line)
        c.drawString(31.92, yt(INTRO_Y), lines[0])
        if len(lines) > 1:
            c.setFont(sf(lines[1]), 9.941)
            c.drawString(31.92, yt(INTRO_Y + 12), lines[1])
        c.setFont(sf(f"Señor(a)   {paciente}"), 9.941)
        c.drawString(31.92, yt(INTRO_Y + 24), f"Señor(a)   {paciente}")

    # Columnas izquierda y derecha
    c.setFont(sf(f"Identificación: {tipo_documento}"), 9.941)
    c.drawString(35.64, yt(ID_Y), f"Identificación: {tipo_documento}")
    c.setFont(sf(f"Sexo:  {sexo}"), 9.941)
    c.drawString(35.64, yt(SEXO_Y), f"Sexo:  {sexo}")
    c.setFont(sf(f"Servicio Ingreso:  {cod_servicio} - {nom_servicio}"), 9.941)
    c.drawString(35.64, yt(SERVICIO_Y), f"Servicio Ingreso:  {cod_servicio} - {nom_servicio}")

    c.setFont(sf(f"Edad: {edad} {medida_edad}"), 9.941)
    c.drawString(365.64, yt(EDAD_Y), f"Edad: {edad} {medida_edad}")
    c.setFont(sf(f"Fecha Nacimiento:  {fecha_nac}"), 9.941)
    c.drawString(365.64, yt(FECHA_NAC_Y), f"Fecha Nacimiento:  {fecha_nac}")
    c.setFont(sf(f"Fecha y hora ingreso:  {fecha_ingreso}"), 9.941)
    c.drawString(365.64, yt(FECHA_ING_Y), f"Fecha y hora ingreso:  {fecha_ingreso}")

    # EL SUSCRITO — justificado: palabras espaciadas para llenar el ancho sin deformar letras
    _sus1 = "EL SUSCRITO MEDICO CERTIFICA LA RELACION CAUSAL DIRECTA ENTRE LAS LESIONES QUE"
    _sus2 = "PRESENTA EL PACIENTE Y EL ACCIDENTE DE TRANSITO."
    _font_sus = sf(_sus1, bold=True)
    _fsize_sus = 9.941
    _target_w = 523.0
    _x0_sus = (W - _target_w) / 2

    def _draw_justified(text, y):
        words = text.split()
        if len(words) <= 1:
            c.setFont(_font_sus, _fsize_sus)
            c.drawCentredString(W/2, y, text)
            return
        total_words_w = sum(c.stringWidth(w, _font_sus, _fsize_sus) for w in words)
        space_w = (_target_w - total_words_w) / (len(words) - 1)
        c.setFont(_font_sus, _fsize_sus)
        x = _x0_sus
        for i, word in enumerate(words):
            c.drawString(x, y, word)
            x += c.stringWidth(word, _font_sus, _fsize_sus)
            if i < len(words) - 1:
                x += space_w

    c.setFont(_font_sus, _fsize_sus)
    c.drawCentredString(W/2, yt(SUSCRITO1_Y), _sus1)
    c.drawCentredString(W/2, yt(SUSCRITO2_Y), _sus2)

    # Firma
    if ruta_png_firma and ruta_png_firma.exists():
        try:
            from PIL import Image as _PI
            img = _PI.open(str(ruta_png_firma))
            iw, ih = img.size
            ratio = iw / ih if ih > 0 else 2.5
            FIRMA_MAX_H = 55.0
            FIRMA_MAX_W = FIRMA_ZW * 0.75
            fw = min(FIRMA_MAX_W, FIRMA_MAX_H * ratio)
            fh = fw / ratio
            if fh > FIRMA_MAX_H: fh = FIRMA_MAX_H; fw = fh * ratio
            fx = FIRMA_X + (FIRMA_ZW - fw) / 2
            fy = yt(LINEA_Y) + 4
            c.drawImage(str(ruta_png_firma), fx, fy, fw, fh, mask="auto")
        except Exception as e:
            log(f"  → Firma PDF error: {e}", "warn")

    # Línea y datos del médico
    c.setLineWidth(0.96)
    c.line(FIRMA_X, yt(LINEA_Y), FIRMA_XR, yt(LINEA_Y))
    c.setFont(sf(f"Dr. {medico}"), 9.941)
    c.drawString(FIRMA_X, yt(MEDICO_Y), f"Dr. {medico}")
    c.setFont(sf(f"Reg.M. {tarjeta_profesional}"), 8.983)
    c.drawString(FIRMA_X, yt(REG_Y), f"Reg.M. {tarjeta_profesional}")

    c.save()
    log(f"  → PDF generado directamente (reportlab)", "info")
    return pdf_path

# ------------------------------------------------------------
# REPORTE EXCEL
# ------------------------------------------------------------
def generar_excel_reporte(session_dir, exitosos, errores) -> Path:
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

# ------------------------------------------------------------
# PROCESO PRINCIPAL (PARALELO CON REPORTLAB)
# ------------------------------------------------------------
def ejecutar_proceso(casos, config, session_dir, completadas_previas, num_workers=3):
    zipf = None
    zip_path = None
    exitosos = []
    errores = []

    try:
        total = len(casos)
        completadas_set = set(completadas_previas)

        zip_path = session_dir / f"certificados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zipf = zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED)
        zip_lock = threading.Lock()

        with LK:
            S["session_dir"] = session_dir
            S["stats"]["total"] = total

        pendientes = [c for c in casos if (
            str(c.get("NoCuenta", c) if isinstance(c, dict) else c)
        ) not in completadas_set]

        ya_hechos = total - len(pendientes)
        if ya_hechos > 0:
            log(f"Reanudando — {ya_hechos} ya completados, {len(pendientes)} pendientes", "warn")
            with LK:
                S["stats"]["ok"] = ya_hechos

        workers = min(num_workers, len(pendientes)) if pendientes else 1
        log(f"🚀 Procesando {len(pendientes)} casos con {workers} workers paralelos", "info")

        # Pool de sesiones HTTP
        sesiones = [req_lib.Session() for _ in range(workers)]
        for s in sesiones:
            s.verify = False

        def worker(caso):
            if isinstance(caso, dict):
                no_cuenta = str(caso.get("NoCuenta", "")).strip()
                no_factura = str(caso.get("NoFactura", "")).strip()
            else:
                no_cuenta = str(caso).strip()
                no_factura = no_cuenta
            nombre_pdf = no_factura if no_factura and no_factura != "None" else no_cuenta

            sesion = sesiones[hash(no_cuenta) % len(sesiones)]
            for intento in range(2):
                if S.get("stopping", False):
                    raise StopBotException("Detenido por usuario")
                try:
                    params = {
                        "IpConexion":   config["ip"],
                        "BdConexion":   "bd",
                        "PortConexion": config["puerto"],
                        "Usuario":      config["usuario"],
                        "PasswordUsu":  config["password"],
                        "codigoEmp":    config["cod_empresa"],
                        "NoCuenta":     no_cuenta,
                    }
                    resp = sesion.get(API_BASE, params=params, timeout=120, verify=False)
                    if resp.status_code != 200:
                        raise Exception(f"HTTP {resp.status_code}")
                    data = resp.json()
                    datos_api = data[0] if isinstance(data, list) and data else data

                    ruta_firma_png = None
                    firma_b64 = datos_api.get("FirmaMedico", "")
                    if firma_b64 and PIL_OK:
                        ruta_firma = session_dir / f"firma_{no_cuenta}.png"
                        if procesar_firma_base64(firma_b64, str(ruta_firma)):
                            ruta_firma_png = ruta_firma

                    pdf_path = generar_pdf_directo(datos_api, no_cuenta, ruta_firma_png, session_dir)

                    if ruta_firma_png and ruta_firma_png.exists():
                        ruta_firma_png.unlink()

                    pdf_final = session_dir / f"CM_{nombre_pdf}.pdf"
                    if pdf_path != pdf_final:
                        if pdf_final.exists():
                            pdf_final.unlink()
                        pdf_path.rename(pdf_final)
                        pdf_path = pdf_final

                    return True, no_cuenta, nombre_pdf, pdf_path, ""
                except StopBotException:
                    raise
                except Exception as e:
                    if intento == 0:
                        time.sleep(1)
                    if intento == 1:
                        return False, no_cuenta, nombre_pdf, None, str(e)
            return False, no_cuenta, nombre_pdf, None, "Error inesperado"

        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(worker, caso): caso for caso in pendientes}
            for future in as_completed(futures):
                if S.get("stopping", False):
                    executor.shutdown(wait=False, cancel_futures=True)
                    raise StopBotException("Detenido por usuario")
                try:
                    success, no_cuenta, nombre_pdf, pdf_path, motivo = future.result()
                except StopBotException:
                    raise
                except Exception as e:
                    log(f"Error inesperado en worker: {e}", "error")
                    continue

                if success and pdf_path and pdf_path.exists():
                    with zip_lock:
                        zipf.write(pdf_path, arcname=f"CM_{nombre_pdf}.pdf")
                    with LK:
                        completadas_set.add(no_cuenta)
                        S["stats"]["ok"] += 1
                        S["completadas"].add(no_cuenta)
                        S["exitosos"].append(no_cuenta)
                    exitosos.append((no_cuenta, nombre_pdf))
                    guardar_progreso(session_dir, completadas_set)
                    log(f"✓ CM_{nombre_pdf}.pdf [{S['stats']['ok']}/{total}]", "ok")
                else:
                    with LK:
                        S["stats"]["err"] += 1
                        S["errores"].append((no_cuenta, motivo))
                    errores.append((no_cuenta, motivo))
                    log(f"✗ {no_cuenta}: {motivo}", "error")

        for s in sesiones:
            try: s.close()
            except: pass

        reporte = generar_excel_reporte(session_dir, exitosos, errores)
        if reporte and reporte.exists():
            zipf.write(reporte, arcname="reporte_resultados.xlsx")

        zipf.close()
        with LK:
            S["zip_path"] = zip_path.name

        if not S["stopping"] and len(completadas_set) >= total:
            marcar_sesion_finalizada(session_dir)
            log("=" * 50, "info")
            log(f"✅ Completado — Total:{total} OK:{len(exitosos)} Errores:{len(errores)}", "ok")
        else:
            log("Sesión guardada para reanudar", "warn")

        if errores:
            log("Casos con error:", "warn")
            for nc, mot in errores[:10]:
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
            except:
                pass
        with LK:
            S["running"] = False
            S["finished"] = True

# ─────────────────────────────────────────────────────────────
# AUTENTICACIÓN
# ─────────────────────────────────────────────────────────────
@app.route("/login")
def login_page():
    if session.get("autenticado"):
        return redirect(url_for("index"))
    return render_template("login.html")

@app.route("/api/auth/login", methods=["POST"])
def auth_login():
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "").split(",")[0].strip()
    ok, msg = _check_rate(ip)
    if not ok:
        return jsonify({"error": msg}), 429

    data     = request.get_json() or {}
    usuario  = data.get("usuario", "").strip()
    password = data.get("password", "").strip()
    servidor = data.get("servidor", "").strip()
    empresas = data.get("empresas", [])  # recibidas del navegador

    if not all([usuario, password, servidor]):
        return jsonify({"error": "Completa todos los campos"}), 400

    if not empresas:
        _fail(ip)
        d = _login_intentos.get(ip, {})
        r = max(0, _LOGIN_MAX - d.get("intentos", 1))
        return jsonify({"error": f"Credenciales inválidas. {r} intento(s) restantes."}), 401

    # Login válido
    _ok(ip)
    session.permanent    = True
    session["autenticado"] = True
    session["usuario"]     = usuario
    session["password"]    = password
    session["servidor"]    = servidor
    session["ip"]          = data.get("ip", "")
    session["puerto"]      = data.get("puerto", "3306")
    log(f"Login: {usuario} @ {servidor}", "info")
    return jsonify({"success": True})

@app.route("/api/auth/logout", methods=["POST"])
def auth_logout():
    session.clear()
    return jsonify({"success": True})

@app.route("/api/auth/me")
def auth_me():
    if session.get("autenticado"):
        return jsonify({"autenticado": True, "usuario": session.get("usuario"), "servidor": session.get("servidor")})
    return jsonify({"autenticado": False}), 401

# ------------------------------------------------------------
# RUTAS FLASK
# ------------------------------------------------------------
@app.route("/")
@login_requerido
def index():
    return render_template("index.html")

@app.route("/api/servidores")
def get_servidores():  # Pública — usada en login
    try:
        servidores = _get_servidores_api()
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
    try:
        data = request.get_json()
        nombre = data.get("servidor", "")
        usuario = data.get("usuario", "")
        password = data.get("password", "")
        if not all([nombre, usuario, password]):
            return jsonify({"error": "Faltan datos"}), 400
        servidores = _get_servidores_api()
        srv = next((s for s in servidores if s.get("Descripcion", "").strip() == nombre), None)
        if not srv:
            return jsonify({"error": f"Servidor '{nombre}' no encontrado"}), 404
        ip = srv.get("Ip_Conexion", "")
        puerto = str(srv.get("Puerto", "3306"))
        empresas = _get_empresas_api(ip, puerto, usuario, password)
        if not empresas:
            return jsonify({"error": "Credenciales inválidas o usuario sin empresas"}), 401
        lista = [{"cod": e.get("Empresa", ""), "nombre": e.get("Nombre_Empresa", "")}
                 for e in empresas if e.get("Empresa")]
        if not lista:
            return jsonify({"error": "No se encontraron empresas"}), 401
        return jsonify({"success": True, "empresas": lista, "ip": ip, "puerto": puerto})
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
        cell.font = Font(bold=True, color="FFFFFF", size=12)
        cell.fill = PatternFill(start_color="1a3a52", end_color="1a3a52", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 20
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name="plantilla_casos.xlsx")

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
                nc = row.get("NoCuenta") or row.get("nocuenta") or row.get("NOCUENTA")
                nf = row.get("NoFactura") or row.get("nofactura") or row.get("NOFACTURA") or row.get("Factura") or ""
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
    casos = data.get("casos", [])
    empresa  = data.get("empresa", "")
    reanudar = data.get("reanudar", False)

    if not casos:
        return jsonify({"error": "No hay casos para procesar"}), 400
    if not empresa:
        return jsonify({"error": "Empresa no seleccionada"}), 400

    config = {
        "servidor":    session.get("servidor", ""),
        "ip":          session.get("ip", ""),
        "puerto":      session.get("puerto", "3306"),
        "cod_empresa": empresa,
        "usuario":     session.get("usuario", ""),
        "password":    session.get("password", ""),
    }
    if not config["ip"]:
        return jsonify({"error": "Sesión inválida, inicia sesión de nuevo"}), 401

    if reanudar:
        session_dir = buscar_sesion_pendiente()
        if session_dir:
            casos_sesion, _ = cargar_casos_sesion(session_dir)
            if casos_sesion:
                casos = casos_sesion
            completadas = cargar_progreso(session_dir)
        else:
            session_dir = DOWNLOAD_DIR / f"sesion_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            session_dir.mkdir(exist_ok=True)
            completadas = set()
            guardar_casos_sesion(session_dir, casos, config)
    else:
        session_dir = DOWNLOAD_DIR / f"sesion_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        session_dir.mkdir(exist_ok=True)
        completadas = set()
        guardar_casos_sesion(session_dir, casos, config)

    with LK:
        S["running"] = True
        S["stopping"] = False
        S["stats"]["total"] = len(casos)

    threading.Thread(
        target=ejecutar_proceso,
        args=(casos, config, session_dir, completadas, 3),
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
        zp = S["zip_path"]
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