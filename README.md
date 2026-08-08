# Preparador de archivos para IA

Aplicación en Python para convertir archivos de texto, HTML, PDF y documentos de Office a Markdown para ser usados con IA.

## Requisitos

- Python 3.10 o superior
- Paquetes incluidos en requirements.txt

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

se puede usar para alimentar las IA de 3CX Y Yeastar

### Interfaz gráfica

```bash
python app_gui.py
```

### Línea de comandos

```bash
python app.py archivo.txt
```

## Instalador para Windows

1. Genera el ejecutable:

```powershell
python build_exe.py
```

2. Compila el instalador (requiere Inno Setup):

```powershell
build_installer.bat
```

3. El instalador resultante queda en `Output\Preparador_de_archivos_para_IA_Installer.exe`.

4. Al instalar en Windows, el instalador crea:
- un folder en el menú de Inicio llamado `voipers`
- un acceso directo en el Escritorio llamado `Preparador de archivos para IA`

