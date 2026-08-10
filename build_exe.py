import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
ICON_FILE = ROOT / "icon.ico"
VENV_PYTHON = ROOT / ".venv" / "Scripts" / "python.exe"
VERSION_INFO_FILE = ROOT / "build" / "version_info.txt"


def get_app_version() -> str:
    """Lee __version__ desde app.py sin importar el módulo (evita depender de sus paquetes)."""
    match = re.search(r'^__version__\s*=\s*["\']([^"\']+)["\']', (ROOT / "app.py").read_text(encoding="utf-8"), re.M)
    if not match:
        raise RuntimeError("No se encontró __version__ en app.py")
    return match.group(1)


def write_version_info(version: str) -> Path:
    """Genera el archivo de recurso de versión que PyInstaller incrusta en el .exe."""
    parts = [int(p) for p in version.split(".")]
    while len(parts) < 4:
        parts.append(0)
    version_tuple = tuple(parts[:4])

    content = f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={version_tuple},
    prodvers={version_tuple},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0),
  ),
  kids=[
    StringFileInfo(
      [
        StringTable(
          '040a04b0',
          [StringStruct('CompanyName', 'voipers'),
          StringStruct('FileDescription', 'Preparador de archivos para IA'),
          StringStruct('FileVersion', '{version}'),
          StringStruct('InternalName', 'preparador_de_archivos_para_ia'),
          StringStruct('OriginalFilename', 'preparador_de_archivos_para_ia.exe'),
          StringStruct('ProductName', 'Preparador de archivos para IA'),
          StringStruct('ProductVersion', '{version}')])
      ]),
    VarFileInfo([VarStruct('Translation', [1034, 1200])])
  ]
)
"""
    VERSION_INFO_FILE.parent.mkdir(parents=True, exist_ok=True)
    VERSION_INFO_FILE.write_text(content, encoding="utf-8")
    return VERSION_INFO_FILE


def get_python_executable():
    """Usa Python de la virtualenv si está disponible, sino el actual."""
    return str(VENV_PYTHON) if VENV_PYTHON.exists() else sys.executable


def install_pillow():
    """Instala Pillow, necesaria para generar el icono por defecto y procesar el logo."""
    subprocess.check_call([get_python_executable(), "-m", "pip", "install", "pillow"])


def ensure_icon():
    """Garantiza que exista un icono válido para la aplicación."""
    if ICON_FILE.exists() and ICON_FILE.stat().st_size > 0:
        try:
            from PIL import Image

            with Image.open(ICON_FILE) as img:
                if img.format == "ICO":
                    return
        except Exception:
            pass

    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        install_pillow()
        from PIL import Image, ImageDraw, ImageFont

    size = 256
    image = Image.new("RGBA", (size, size), "#1f4e79")
    draw = ImageDraw.Draw(image)

    try:
        font = ImageFont.truetype("arial.ttf", 160)
    except Exception:
        font = ImageFont.load_default()

    text = "IA"
    try:
        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
    except AttributeError:
        text_width, text_height = font.getsize(text)

    text_position = ((size - text_width) / 2, (size - text_height) / 2)
    draw.text(text_position, text, fill="white", font=font)
    image.save(ICON_FILE, format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])


def ensure_dependency(module_name: str, package_name: str | None = None):
    """Comprueba que un módulo esté instalado y, si falta, lo instala con pip."""
    try:
        __import__(module_name)
    except ImportError:
        subprocess.check_call([get_python_executable(), "-m", "pip", "install", package_name or module_name])


def find_rapidocr_package_dir() -> Path:
    """Localiza la carpeta de instalación de rapidocr_onnxruntime en el intérprete que compila."""
    output = subprocess.check_output(
        [get_python_executable(), "-c", "import rapidocr_onnxruntime, os; print(os.path.dirname(rapidocr_onnxruntime.__file__))"],
        cwd=str(ROOT),
        text=True,
    )
    return Path(output.strip())


def build_exe():
    """Genera un ejecutable de Windows usando PyInstaller."""
    ensure_dependency("PyInstaller", "pyinstaller")
    ensure_dependency("docx", "python-docx")

    ensure_icon()

    version = get_app_version()
    version_info_file = write_version_info(version)
    print(f"Compilando versión {version}...")

    hidden_imports = [
        "docx",
        "pptx",
        "openpyxl",
        "pymupdf",
        "rapidocr_onnxruntime",
        "striprtf",
        "tkinterdnd2",
        "requests",
        "markdownify",
        "bs4",
        # Usados solo dentro de los submódulos de rapidocr_onnxruntime que se cargan en
        # tiempo de ejecución (ver comentario sobre '--add-data' más abajo); al no ser
        # analizados estáticamente, PyInstaller no los detecta por sí solo.
        "pyclipper",
        "shapely",
        "shapely.geometry",
        "six",
    ]
    rapidocr_dir = find_rapidocr_package_dir()
    command = [
        get_python_executable(),
        "-m",
        "PyInstaller",
        "--onefile",
        "--windowed",
        "--icon",
        str(ICON_FILE),
        "--name",
        "preparador_de_archivos_para_ia",
        "--add-data",
        f"{ROOT / 'assets' / 'logo.jpg'};assets",
        "--add-data",
        f"{ROOT / 'LICENSE'};.",
        "--version-file",
        str(version_info_file),
        # rapidocr_onnxruntime carga sus submódulos (detector/reconocedor/clasificador) en
        # tiempo de ejecución con importlib, añadiendo su propia carpeta a sys.path; ese truco
        # necesita que el paquete exista como carpeta real en disco, no solo empaquetado dentro
        # del .exe, así que se copia entero como datos sueltos (no basta con --collect-data,
        # que solo copiaría sus .yaml/.onnx y dejaría los .py fuera).
        "--add-data",
        f"{rapidocr_dir};rapidocr_onnxruntime",
        "app_gui.py",
    ]
    for hidden in hidden_imports:
        command.extend(["--hidden-import", hidden])

    subprocess.check_call(command, cwd=str(ROOT))


if __name__ == "__main__":
    build_exe()
