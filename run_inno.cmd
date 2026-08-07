@echo off
cd /d "%~dp0"
"C:\Users\Usuario\AppData\Local\Programs\Inno Setup 6\ISCC.exe" /O"Output" preparador_de_archivos_para_ia_installer.iss
echo EXITCODE=%ERRORLEVEL%
dir Output 2>nul || echo Output dir not found
