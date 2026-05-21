# Esculapio Soft - Despliegue en Railway

## Archivos añadidos para Railway

| Archivo | Propósito |
|----------|-----------|
| `Procfile` | Indica a Railway cómo ejecutar la aplicación |
| `railway.json` | Configuración del despliegue |
| `nixpacks.toml` | Paquetes del sistema (LibreOffice para PDF) |
| `convertidor_pdf.py` | Conversión DOCX → PDF usando LibreOffice (Linux) |

## No necesita venv

**Railway installa las dependencias automáticamente desde `requirements.txt`.** No necesitas incluir un `venv` en el proyecto — es innecesario y会增加 el tamaño.

## Pasos para desplegar

1. **Sube el proyecto a GitHub** (o conecta directamente Railway a tu repositorio)

2. **Conecta Railway a tu repositorio**:
   - Ve a [railway.app](https://railway.app)
   - New Project → Deploy from GitHub
   - Selecciona el repositorio

3. **Configura las variables de entorno** (opcional si Railway las detecta):
   - Railway detecta automáticamente `requirements.txt`
   - Asegúrate de que el buildpack Nixpacks esté seleccionado

4. **Deploy automático**:
   - Railway detectará el `Procfile` y ejecutará `python app.py`

## Nota importante sobre la conversión PDF

El proyecto usa **LibreOffice** para convertir DOCX a PDF en el servidor (no tiene Word como en Windows). Esto está configurado en `nixpacks.toml` con los paquetes:
- `libreoffice`
- `libreoffice-calc`
- `fonts-dejavu`

## Estructura del proyecto

```
Esculapio Soft/
├── app.py              # Aplicación Flask principal
├── FORMATO.docx        # Plantilla del certificado
├── requirements.txt    # Dependencias Python
├── Procfile            # Comando de inicio
├── railway.json        # Configuración Railway
├── nixpacks.toml       # Paquetes del sistema
├── convertidor_pdf.py  # Conversión DOCX→PDF (LibreOffice)
├── static/             # Archivos estáticos
│   ├── favicon.ico
│   ├── LOGO.png
│   └── SIS.png
└── templates/          # Templates HTML
    └── index.html
```

## Solución de problemas

### Error: "LibreOffice no está instalado"
Verifica que `nixpacks.toml` tenga los paquetes `libreoffice` y `libreoffice-calc`.

### Error: "Puerto en uso"
Railway asigna el puerto automáticamente. La app usa `PORT` del entorno.

### Build fallido
Revisa los logs de Railway para ver errores de dependencias o build.