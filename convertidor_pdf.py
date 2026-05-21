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
    """
    Convierte un archivo DOCX a PDF usando LibreOffice CLI.
    Esta función es la alternativa a docx2pdf para servidores Linux (Railway, etc.)

    Args:
        ruta_docx: Ruta al archivo .docx de entrada
        directorio_salida: Directorio donde se guardará el PDF

    Returns:
        Path al archivo PDF generado

    Raises:
        Exception: Si la conversión falla o LibreOffice no está instalado
    """
    if not ruta_docx.exists():
        raise FileNotFoundError(f"El archivo {ruta_docx} no existe")

    # Crear directorio de salida si no existe
    directorio_salida.mkdir(parents=True, exist_ok=True)

    # Copiar el archivo a un directorio temporal para la conversión
    # LibreOffice guarda el PDF en el mismo directorio que el DOCX
    temp_dir = tempfile.mkdtemp()

    try:
        # Copiar DOCX al directorio temporal
        docx_en_temp = Path(temp_dir) / ruta_docx.name
        shutil.copy(ruta_docx, docx_en_temp)

        # Ejecutar LibreOffice en modo headless para convertir
        # --headless: sin interfaz gráfica
        # --convert-to pdf: convertir a PDF
        # --outdir: directorio de salida
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
            timeout=120  # 2 minutos máximo
        )

        if resultado.returncode != 0:
            raise Exception(f"Error en LibreOffice: {resultado.stderr}")

        # El PDF se genera en directorio_salida con el mismo nombre
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
        # Limpiar directorio temporal
        shutil.rmtree(temp_dir, ignore_errors=True)


def detectar_metodo_conversion():
    """
    Detecta qué método de conversión está disponible.
    Returns: 'libreoffice', 'docx2pdf', o None
    """
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