import tempfile
import unittest
from pathlib import Path

from docx import Document as DocxDocument
from openpyxl import Workbook as OpenPyXLWorkbook
from pptx import Presentation as PptxPresentation
from reportlab.pdfgen import canvas

from app import convert_file, convert_files, convert_text_to_markdown


class ConverterTests(unittest.TestCase):
    def test_plain_text_becomes_markdown(self):
        text = "Hola mundo\n\nEste es un ejemplo"
        expected = "# Hola mundo\n\nEste es un ejemplo"
        self.assertEqual(convert_text_to_markdown(text), expected)

    def test_html_is_converted_to_markdown(self):
        html = "<h1>Titulo</h1><p>Texto de prueba</p>"
        expected = "# Titulo\n\nTexto de prueba"
        self.assertEqual(convert_text_to_markdown(html), expected)

    def test_docx_is_converted_to_markdown(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "document.docx"
            doc = DocxDocument()
            doc.add_paragraph("Contenido de Word")
            doc.save(input_path)

            output_path = Path(tmpdir) / "document.md"
            convert_file(str(input_path), str(output_path))

            self.assertIn("Contenido de Word", output_path.read_text(encoding="utf-8"))

    def test_pdf_is_converted_to_markdown(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "document.pdf"
            pdf = canvas.Canvas(str(input_path))
            pdf.drawString(100, 750, "Contenido de PDF")
            pdf.save()

            output_path = Path(tmpdir) / "document.md"
            convert_file(str(input_path), str(output_path))

            self.assertIn("Contenido de PDF", output_path.read_text(encoding="utf-8"))

    def test_excel_is_converted_to_markdown(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "document.xlsx"
            workbook = OpenPyXLWorkbook()
            sheet = workbook.active
            sheet["A1"] = "Contenido de Excel"
            workbook.save(input_path)
            workbook.close()

            output_path = Path(tmpdir) / "document.md"
            convert_file(str(input_path), str(output_path))

            self.assertIn("Contenido de Excel", output_path.read_text(encoding="utf-8"))

    def test_pptx_is_converted_to_markdown(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "document.pptx"
            presentation = PptxPresentation()
            slide = presentation.slides.add_slide(presentation.slide_layouts[1])
            slide.shapes[1].text = "Contenido de PowerPoint"
            presentation.save(input_path)

            output_path = Path(tmpdir) / "document.md"
            convert_file(str(input_path), str(output_path))

            self.assertIn("Contenido de PowerPoint", output_path.read_text(encoding="utf-8"))

    def test_batch_conversion_creates_multiple_markdown_files(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            txt_path = tmp / "uno.txt"
            txt_path.write_text("Primer archivo", encoding="utf-8")
            html_path = tmp / "dos.html"
            html_path.write_text("<h1>Segundo</h1>", encoding="utf-8")

            outputs = convert_files([str(txt_path), str(html_path)], output_dir=str(tmp))

            self.assertEqual(len(outputs), 2)
            self.assertTrue(Path(outputs[0]).exists())
            self.assertTrue(Path(outputs[1]).exists())


if __name__ == "__main__":
    unittest.main()
