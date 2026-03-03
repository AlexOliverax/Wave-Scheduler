; ─────────────────────────────────────────────────────────────────────────────
; Waves Scheduler — Inno Setup Script v2.0
; ─────────────────────────────────────────────────────────────────────────────

#define AppName    "Waves Scheduler"
#define AppVersion "2.0.0"
#define AppPublisher "Antigravity"
#define AppExeName "Waves Scheduler.exe"

[Setup]
AppId={{A7F3D12B-4C8E-4F9A-B234-56789ABCDEF0}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=https://github.com/antigravity
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes

; Define o diretório base para os arquivos em [Files]
; Sem isso, o Inno Setup resolve caminhos relativos a partir do CWD do IDE
SourceDir=E:\Projetos - Antigravity\Project - Waves Scheduler

; Ícone que aparece no atalho de desinstalação — usa o ícone dentro do app instalado
UninstallDisplayIcon={app}\icon\WavesS.ico

; Pasta onde o instalador .exe será gerado (relativa ao .iss)
OutputDir=installer
OutputBaseFilename=WavesScheduler_Setup_v{#AppVersion}

; Compressão
Compression=lzma2/ultra64
SolidCompression=yes

; Requer Windows 10+ x64
MinVersion=10.0
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

; Interface
WizardStyle=modern

; Privilégios (permite instalar sem ser admin, com aviso)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; Desinstalar
Uninstallable=yes


[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"


[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked


[Files]
; Copiar TODA a pasta gerada pelo PyInstaller
; O caminho Source é relativo ao local do arquivo .iss
Source: "dist\Waves Scheduler\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Copiar ícones explicitamente para garantir atalhos com ícone correto
Source: "icon\*"; DestDir: "{app}\icon"; Flags: ignoreversion


[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\icon\WavesS.ico"
Name: "{group}\Desinstalar {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\icon\WavesS.ico"; Tasks: desktopicon


[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent


[UninstallDelete]
; Apaga a pasta inteira do aplicativo (incluindo logs, history.json, config.json etc)
Type: filesandordirs; Name: "{app}"
