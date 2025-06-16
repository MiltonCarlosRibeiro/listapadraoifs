[Setup]
AppName=Lista Padrão IFS
AppVersion=1.0
DefaultDirName={pf}\ListaPadraoIFS
DefaultGroupName=ListaPadraoIFS
OutputDir=C:\listapadraoifs\installer
OutputBaseFilename=listapadraoifsInstaller
SetupIconFile=C:\listapadraoifs\listapadraoifs\out\artifacts\listapadraoifs_jar\logo.ico
Compression=lzma
SolidCompression=yes

[Files]
Source: "C:\listapadraoifs\listapadraoifs\out\artifacts\listapadraoifs_jar\listapadraoifs.jar"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\listapadraoifs\listapadraoifs\out\artifacts\listapadraoifs_jar\start.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\listapadraoifs\listapadraoifs\out\artifacts\listapadraoifs_jar\logo.ico"; DestDir: "{app}"

[Icons]
Name: "{group}\Lista Padrão IFS"; Filename: "{app}\start.bat"; WorkingDir: "{app}"; IconFilename: "{app}\logo.ico"
Name: "{commondesktop}\Lista Padrão IFS"; Filename: "{app}\start.bat"; WorkingDir: "{app}"; IconFilename: "{app}\logo.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\start.bat"; Description: "Iniciar Lista Padrão IFS"; Flags: nowait postinstall skipifsilent

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Área de Trabalho"
