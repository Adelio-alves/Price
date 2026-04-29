#define MyAppName "Price"
#define MyAppVersion "35.0"
#define MyAppPublisher "Adelio Alves"
#define MyAppURL "https://adelioalves.com"
#define MyAppExeName "price.exe"

[Setup]
AppId={{9F6A0A4A-7B9E-4A6F-8C2D-5E1B3F7A9C10}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
SourceDir=E:\python\price
OutputDir=output_installer
OutputBaseFilename=Instalador_Price_v34_0
SetupIconFile=app.ico
LicenseFile=EULA.rtf
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked

[Files]
Source: "dist\price\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "autorizacao.json"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "EULA.rtf"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Executar {#MyAppName} agora"; Flags: nowait postinstall skipifsilent