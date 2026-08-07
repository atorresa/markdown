@echo off
setlocal
if not exist "dist\preparador_de_archivos_para_ia.exe" goto noexe

set "ISCC_PATH=%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC_PATH%" (
  if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    set "ISCC_PATH=C:\Program Files\Inno Setup 6\ISCC.exe"
  )
)
if not exist "%ISCC_PATH%" (
  if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set "ISCC_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
  )
)
if not exist "%ISCC_PATH%" goto noinno

"%ISCC_PATH%" /O"Output" preparador_de_archivos_para_ia_installer.iss
if ERRORLEVEL 1 goto fail

echo Instalador generado correctamente.
if exist "Output\Preparador_de_archivos_para_IA_Installer.exe" echo Output\Preparador_de_archivos_para_IA_Installer.exe
goto end

:noexe
echo ERROR: no se encontro dist\preparador_de_archivos_para_ia.exe
echo Ejecuta build_exe.py primero para generar el ejecutable.
exit /b 1

:noinno
echo ERROR: No se encontro Inno Setup Compiler (ISCC.exe).
echo Instala Inno Setup y asegurate de que ISCC.exe exista en la ruta indicada o en PATH.
exit /b 1

:fail
echo ERROR: La compilacion del instalador fallo.
exit /b 1

:end
endlocal
