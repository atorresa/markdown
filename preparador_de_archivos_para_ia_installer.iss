[Setup]
AppName=Preparador de archivos para IA
AppVersion=2.0.0
DefaultDirName={autopf}\Preparador de archivos para IA
DefaultGroupName=voipers
DisableProgramGroupPage=no
AllowNoIcons=yes
OutputBaseFilename=Preparador_de_archivos_para_IA_Installer
OutputDir=Output
Compression=lzma2
SolidCompression=yes
SetupIconFile=icon.ico
ArchitecturesAllowed=x64compatible

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el Escritorio"; GroupDescription: "Tareas adicionales:"

[Files]
Source: "dist\preparador_de_archivos_para_ia.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Preparador de archivos para IA"; Filename: "{app}\preparador_de_archivos_para_ia.exe"; IconFilename: "{app}\icon.ico"
Name: "{commondesktop}\Preparador de archivos para IA"; Filename: "{app}\preparador_de_archivos_para_ia.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\preparador_de_archivos_para_ia.exe"; Description: "Abrir Preparador de archivos para IA"; Flags: nowait postinstall skipifsilent
