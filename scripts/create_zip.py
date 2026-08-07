"""Empaqueta el ejecutable ya compilado junto con el README en un .zip para distribuirlo."""
from pathlib import Path
import zipfile

root = Path(__file__).resolve().parents[1] / 'dist'
exe = root / 'preparador_de_archivos_para_ia.exe'
readme = Path(__file__).resolve().parents[1] / 'README.md'
zip_path = root / 'preparador_de_archivos_para_ia.zip'

with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
    if exe.exists():
        z.write(exe, exe.name)
    if readme.exists():
        z.write(readme, readme.name)

if zip_path.exists():
    print('Zip creado:', zip_path, '- tamaño:', zip_path.stat().st_size, 'bytes')
else:
    print('Error: no se pudo crear el zip')
