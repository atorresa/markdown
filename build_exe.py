import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
ICON_FILE = ROOT / "icon.ico"
VENV_PYTHON = ROOT / ".venv" / "Scripts" / "python.exe"


def get_python_executable():
    """Usa Python de la virtualenv si está disponible, sino el actual."""
    return str(VENV_PYTHON) if VENV_PYTHON.exists() else sys.executable


def install_pillow():
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
    try:
        __import__(module_name)
    except ImportError:
        subprocess.check_call([get_python_executable(), "-m", "pip", "install", package_name or module_name])


def build_exe():
    """Genera un ejecutable de Windows usando PyInstaller."""
    ensure_dependency("PyInstaller", "pyinstaller")
    ensure_dependency("docx", "python-docx")

    ensure_icon()

    hidden_imports = ["docx", "pptx", "openpyxl", "pypdf"]
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
        "app_gui.py",
    ]
    for hidden in hidden_imports:
        command.extend(["--hidden-import", hidden])

    subprocess.check_call(command, cwd=str(ROOT))


if __name__ == "__main__":
    build_exe()
