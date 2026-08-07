import html
import re
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from docx import Document as DocxDocument
from markdownify import markdownify as html_page_to_markdown
from openpyxl import load_workbook
from pptx import Presentation as PptxPresentation
from pypdf import PdfReader


# Extensiones de archivo local que la aplicación sabe leer (además de las páginas web por URL).
SUPPORTED_EXTENSIONS = {
    ".txt": "text",
    ".html": "html",
    ".htm": "html",
    ".docx": "docx",
    ".pdf": "pdf",
    ".xlsx": "xlsx",
    ".pptx": "pptx",
}

# Etiquetas que normalmente son ruido de maquetación (menús, cabeceras, anuncios...)
# y no forman parte del contenido real de una página web.
_BOILERPLATE_TAGS = ["script", "style", "noscript", "template", "nav", "header", "footer", "aside", "svg", "iframe", "form"]

# Algunos sitios web bloquean peticiones sin un User-Agent de navegador "real".
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) PreparadorDeArchivosParaIA/1.0",
}


def is_url(value: str) -> bool:
    """Indica si el valor recibido es una URL http(s) en vez de una ruta de archivo local."""
    parsed = urlparse(value)
    return parsed.scheme in ("http", "https") and bool(parsed.netloc)


def _fetch_url(url: str) -> str:
    """Descarga el HTML de una página web."""
    response = requests.get(url, headers=REQUEST_HEADERS, timeout=20)
    response.raise_for_status()
    if response.encoding is None:
        response.encoding = response.apparent_encoding
    return response.text


def _slugify(value: str) -> str:
    """Convierte un texto libre en un nombre de archivo simple y seguro."""
    value = re.sub(r"[^\w\-]+", "-", value.strip().lower())
    value = re.sub(r"-{2,}", "-", value).strip("-")
    return value[:80] or "pagina"


def _suggested_filename_from_url(url: str, soup: BeautifulSoup) -> str:
    """Sugiere un nombre de archivo a partir del <title> de la página o, si falta, de la URL."""
    title_tag = soup.find("title")
    if title_tag and title_tag.get_text(strip=True):
        return _slugify(title_tag.get_text(strip=True))
    parsed = urlparse(url)
    return _slugify(f"{parsed.netloc}{parsed.path}")


def convert_url_to_markdown(url: str) -> tuple[str, str]:
    """Descarga una página web y la convierte a Markdown. Devuelve (markdown, nombre_sugerido)."""
    html_content = _fetch_url(url)
    soup = BeautifulSoup(html_content, "html.parser")
    suggested_name = _suggested_filename_from_url(url, soup)

    for tag in soup(_BOILERPLATE_TAGS):
        tag.decompose()

    body = soup.body or soup
    markdown = html_page_to_markdown(str(body), heading_style="ATX")
    markdown = re.sub(r"\n{3,}", "\n\n", markdown)
    return markdown.strip(), suggested_name


def _read_text_from_docx(path: Path) -> str:
    """Extrae el texto de un archivo Word (.docx)."""
    doc = DocxDocument(path)
    return "\n".join(paragraph.text for paragraph in doc.paragraphs if paragraph.text)


def _read_text_from_pdf(path: Path) -> str:
    """Extrae el texto de un archivo PDF."""
    reader = PdfReader(str(path))
    pages = [page.extract_text() or "" for page in reader.pages]
    return "\n".join(page for page in pages if page)


def _read_text_from_excel(path: Path) -> str:
    """Extrae el texto de un archivo Excel (.xlsx)."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        texts = []
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows(values_only=True):
                values = [str(value) for value in row if value is not None]
                if values:
                    texts.append(" | ".join(values))
        return "\n".join(texts)
    finally:
        workbook.close()


def _read_text_from_pptx(path: Path) -> str:
    """Extrae el texto de una presentación PowerPoint (.pptx)."""
    presentation = PptxPresentation(str(path))
    slides_text = []
    for slide in presentation.slides:
        texts = [shape.text for shape in slide.shapes if hasattr(shape, "text") and shape.text]
        if texts:
            slides_text.append("\n".join(texts))
    return "\n\n".join(slides_text)


def read_input_text(input_path: str) -> str:
    """Lee el contenido de un archivo según su extensión."""
    input_file = Path(input_path)
    suffix = input_file.suffix.lower()

    if suffix == ".docx":
        return _read_text_from_docx(input_file)
    if suffix == ".pdf":
        return _read_text_from_pdf(input_file)
    if suffix == ".xlsx":
        return _read_text_from_excel(input_file)
    if suffix == ".pptx":
        return _read_text_from_pptx(input_file)

    return input_file.read_text(encoding="utf-8")


def convert_text_to_markdown(text: str) -> str:
    """Convierte texto plano o HTML simple a Markdown."""
    if not text:
        return ""

    cleaned = text.strip()

    # Si el contenido parece HTML, se transforma a una estructura Markdown básica.
    if "<" in cleaned and ">" in cleaned:
        cleaned = re.sub(r"<h1[^>]*>(.*?)</h1>", r"# \1\n", cleaned, flags=re.IGNORECASE | re.DOTALL)
        cleaned = re.sub(r"<h2[^>]*>(.*?)</h2>", r"## \1\n", cleaned, flags=re.IGNORECASE | re.DOTALL)
        cleaned = re.sub(r"<h3[^>]*>(.*?)</h3>", r"### \1\n", cleaned, flags=re.IGNORECASE | re.DOTALL)
        cleaned = re.sub(r"<p[^>]*>(.*?)</p>", r"\n\n\1", cleaned, flags=re.IGNORECASE | re.DOTALL)
        cleaned = re.sub(r"<br\s*/?>", "\n", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"<[^>]+>", "", cleaned)
        cleaned = html.unescape(cleaned)
        cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
        return cleaned.strip()

    lines = [line.rstrip() for line in cleaned.splitlines()]
    if lines and lines[0].strip():
        lines[0] = f"# {lines[0].strip()}"
    return "\n\n".join(part for part in lines if part.strip())


def convert_file(input_path: str, output_path: Optional[str] = None) -> str:
    """Convierte un archivo individual o una URL a Markdown y lo guarda en disco."""
    if is_url(input_path):
        markdown, suggested_name = convert_url_to_markdown(input_path)
        default_output = Path.cwd() / f"{suggested_name}.md"
    else:
        input_file = Path(input_path)
        content = read_input_text(input_path)
        markdown = convert_text_to_markdown(content)
        suggested_name = input_file.stem
        default_output = input_file.with_suffix(".md")

    if output_path is None:
        output_path = default_output
    else:
        output_path = Path(output_path)
        # Si nos dieron una carpeta (no un nombre de archivo concreto), generamos el nombre dentro de ella.
        if output_path.is_dir():
            output_path = output_path / f"{suggested_name}.md"

    # Crea la carpeta de salida si no existe.
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(markdown, encoding="utf-8")
    return str(output_path)


def convert_files(input_paths: list[str], output_dir: Optional[str] = None) -> list[str]:
    """Convierte varios archivos y/o URLs y devuelve las rutas de los archivos Markdown generados."""
    return [convert_file(input_path, output_dir) for input_path in input_paths]


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Convierte un archivo o una URL a Markdown")
    parser.add_argument("input", help="Ruta del archivo de entrada, o una URL (http/https)")
    parser.add_argument("output", nargs="?", help="Ruta de salida opcional (archivo o carpeta)")
    args = parser.parse_args()

    result = convert_file(args.input, args.output)
    print(f"Archivo convertido: {result}")
