import html
import re
from pathlib import Path
from typing import Optional

from docx import Document as DocxDocument
from openpyxl import load_workbook
from pptx import Presentation as PptxPresentation
from pypdf import PdfReader


SUPPORTED_EXTENSIONS = {
    ".txt": "text",
    ".html": "html",
    ".htm": "html",
    ".docx": "docx",
    ".pdf": "pdf",
    ".xlsx": "xlsx",
    ".pptx": "pptx",
}


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
    """Convierte un archivo individual a Markdown y lo guarda en disco."""
    input_file = Path(input_path)
    content = read_input_text(input_path)
    markdown = convert_text_to_markdown(content)

    if output_path is None:
        output_path = input_file.with_suffix(".md")
    else:
        output_path = Path(output_path)

    # Crea la carpeta de salida si no existe.
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(markdown, encoding="utf-8")
    return str(output_path)


def convert_files(input_paths: list[str], output_dir: Optional[str] = None) -> list[str]:
    """Convierte varios archivos y devuelve las rutas de los archivos Markdown generados."""
    outputs = []
    target_dir = Path(output_dir) if output_dir else None
    for input_path in input_paths:
        input_file = Path(input_path)
        output_path = None
        if target_dir is not None:
            output_path = target_dir / f"{input_file.stem}.md"
        outputs.append(convert_file(str(input_file), str(output_path) if output_path else None))
    return outputs


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Convert text or HTML to Markdown")
    parser.add_argument("input", help="Path to the input file")
    parser.add_argument("output", nargs="?", help="Optional output file path")
    args = parser.parse_args()

    result = convert_file(args.input, args.output)
    print(f"Archivo convertido: {result}")
