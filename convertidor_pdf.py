#!/usr/bin/env python3
"""
Módulo de conversión DOCX a PDF para Linux usando LibreOffice CLI.
Alternativa multiplataforma a docx2pdf que solo funciona en Windows.
"""

import subprocess
import tempfile
import shutil
from pathlib import Path

def docx_a_pdf_libreoffice(ruta_docx: Path, directorio_salida: Path) -> Path:
    if not ruta_docx.exists():
        raise FileNotFoundError(f"El archivo {ruta_docx} no existe")
    directorio_salida.mkdir(parents=True, exist_ok=True)
    temp_dir = tempfile.mkdtemp()
    try:
        docx_en_temp = Path(temp_dir) / ruta_docx.name
        shutil.copy(ruta_docx, docx_en_temp)
        resultado = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(directorio_salida),
                str(docx_en_temp)
            ],
            capture_output=True,
            text=True,
            timeout=120
        )
        if resultado.returncode != 0:
            raise Exception(f"Error en LibreOffice: {resultado.stderr}")
        pdf_path = directorio_salida / f"{ruta_docx.stem}.pdf"
        if not pdf_path.exists():
            raise Exception(f"El PDF no fue generado en {pdf_path}")
        return pdf_path
    except FileNotFoundError:
        raise Exception(
            "LibreOffice no está instalado. "
            "En Railway, añade esto al Railway.toml:\n"
            "[nixpacks]\n"
            "aptPackages = [\"libreoffice\", \"libreoffice-calc\"]"
        )
    except subprocess.TimeoutExpired:
        raise Exception("Tiempo de espera agotado en la conversión (120s)")
    except Exception as e:
        raise Exception(f"Error convirtiendo a PDF: {e}")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

def detectar_metodo_conversion():
    import platform
    import shutil
    if platform.system() == "Windows":
        try:
            from docx2pdf import convert
            return 'docx2pdf'
        except ImportError:
            return None
    else:
        if shutil.which("libreoffice"):
            return 'libreoffice'
        return None