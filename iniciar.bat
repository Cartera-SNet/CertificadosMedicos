@echo off
setlocal
set "DIR=%~dp0"
set "VENV=%DIR%venv"

echo.
echo ============================================================
echo   Bot Esculapio Soft - Certificaciones SOAT v2.0
echo   Generacion masiva de certificados medicos
echo ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado. Instala Python 3.10 o superior.
    pause & exit /b 1
)
for /f "tokens=*" %%v in ('python --version') do echo [OK] %%v

if not exist "%VENV%\Scripts\python.exe" (
    echo Creando entorno virtual...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo [ERROR] No se pudo crear el entorno virtual.
        pause & exit /b 1
    )
    echo [OK] Entorno virtual creado.
)

echo Actualizando pip...
"%VENV%\Scripts\python.exe" -m pip install --upgrade pip --quiet

echo Instalando dependencias base...
"%VENV%\Scripts\python.exe" -m pip install flask==3.0.3 --quiet
"%VENV%\Scripts\python.exe" -m pip install requests==2.31.0 --quiet
"%VENV%\Scripts\python.exe" -m pip install openpyxl==3.1.2 --quiet
"%VENV%\Scripts\python.exe" -m pip install python-docx==1.1.2 --quiet
"%VENV%\Scripts\python.exe" -m pip install docx2pdf==0.1.8 --quiet
"%VENV%\Scripts\python.exe" -m pip install urllib3 --quiet

echo Instalando Pillow (imagenes de firma)...
"%VENV%\Scripts\python.exe" -m pip install Pillow --quiet
if errorlevel 1 (
    echo [WARN] Pillow no se instalo. Las firmas se omitiran en los PDFs.
)

if not exist "%DIR%downloads" mkdir "%DIR%downloads"
if not exist "%DIR%templates" mkdir "%DIR%templates"
if not exist "%DIR%static" mkdir "%DIR%static"

echo.
echo [OK] Todo listo.
echo.
echo Iniciando servidor en: http://localhost:5050
echo Abriendo navegador...
echo.
echo Presiona Ctrl+C para detener el servidor.
echo.

"%VENV%\Scripts\python.exe" "%DIR%app.py"

pause
