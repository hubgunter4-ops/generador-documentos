# Combina Word

Aplicación de escritorio en Python/Tkinter para rellenar una plantilla Word de varias páginas con los datos de cada fila de Excel y generar **un único documento `.docx`**.

## Funcionalidades

- Flujo guiado: primero se carga la plantilla Word y después Excel.
- Al cargar Excel se detecta automáticamente el número de registros y se propone esa cantidad de copias/páginas.
- La cantidad detectada puede reducirse antes de generar el documento.
- Unión de todas las copias en un documento final.
- Marcadores en párrafos, tablas, encabezados y pies de página.
- Procesamiento en segundo plano para que la interfaz no se congele.
- Empaquetado para Windows (`.exe`) y Debian/Ubuntu (`.deb`).

## Plantilla y Excel

En Word usa marcadores con doble llave, por ejemplo:

```text
Nombre: {{NOMBRE}}
Identificación: {{IDENTIFICACION}}
Domicilio: {{DOMICILIO}}
Observaciones: {{OBSERVACIONES}}
```

Las columnas de Excel deben corresponder a esos nombres. Los encabezados se normalizan a mayúsculas y se recortan espacios; la correspondencia con los marcadores Word no distingue mayúsculas de minúsculas. Cada fila es un registro independiente.

En la interfaz, la cantidad detectada determina cuántas filas iniciales del Excel se procesan. Por ejemplo, con 50 registros se proponen 50 copias; si se cambia a `5`, se generan las cinco primeras copias completas de la plantilla.

Si la plantilla tiene varias páginas, conserva los saltos de página. Para que cada registro empiece en una nueva página, añade un salto de página al final de la plantilla (`Ctrl + Enter`).

> Los campos dentro de cuadros de texto o formas flotantes de Word no son compatibles con `python-docx`; usa párrafos o tablas normales.

## Ejecutar desde código fuente

Requiere Python 3.11 o posterior.

```bash
python -m venv .venv
# Linux/macOS
source .venv/bin/activate
# Windows PowerShell
.venv\Scripts\Activate.ps1

python -m pip install -r requirements.txt
python main.py
```

## Pruebas

```bash
python -m pip install pytest
pytest -q
```

## Construir el `.exe` en Windows

El ejecutable debe construirse en Windows para obtener un binario Windows confiable:

```powershell
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
pyinstaller --clean --noconfirm combina-word.spec
```

El resultado queda en `dist/CombinaWord.exe` o en la carpeta `dist/` según la versión de PyInstaller.

También se puede usar:

```powershell
powershell -ExecutionPolicy Bypass -File packaging/build-windows.ps1
```

## Construir el `.deb` en Debian/Ubuntu

En una máquina Debian/Ubuntu, instala Tkinter antes de construir:

```bash
sudo apt-get update && sudo apt-get install -y python3-tk dpkg-dev
```

```bash
chmod +x packaging/build-deb.sh
./packaging/build-deb.sh
```

El script crea `dist/combina-word_VERSION_amd64.deb`. El paquete instala una aplicación PyInstaller autónoma en `/opt/combina-word` y un lanzador en `/usr/bin/combina-word`.

Instalación:

```bash
sudo apt install ./dist/combina-word_*.deb
combina-word
```

El paquete `.deb` requiere un entorno gráfico Linux. El `.exe` es el formato recomendado para Windows.

## GitHub Actions

El workflow `.github/workflows/build.yml` se ejecuta automáticamente:

- En cada `push` a cualquier rama o etiqueta.
- En cada `pull request`.
- Manualmente desde **Actions → Build Combina Word → Run workflow**.

Después de una ejecución correcta, en la sección **Artifacts** del run estarán disponibles:

- `CombinaWord-windows.zip` con el `.exe`.
- `combina-word_amd64.deb`.

El workflow ejecuta las pruebas antes de compilar. Si las pruebas fallan, no se generan los paquetes.

Para publicar artefactos, sube este repositorio a GitHub y ejecuta el workflow desde la pestaña **Actions**.

## Limitaciones conocidas

- Se requiere una plantilla `.docx`, no `.doc`.
- Excel debe ser `.xlsx` o `.xls` compatible con pandas.
- El formato de un párrafo puede conservar principalmente el formato del primer fragmento de texto si un marcador está dividido en varios runs.
- La conversión a PDF no está incluida; puede añadirse posteriormente con LibreOffice o Word.
