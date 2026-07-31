; ============================================================
;  Inno Setup 6 Script — Delta Dyeing PO Converter
;  https://jrsoftware.org/isdl.php
;
;  HOW TO USE:
;    1. Run build.bat first to generate dist\DeltaDyeingConverter\
;    2. Open this file in Inno Setup Compiler (Ctrl+F9 to build)
;    3. Installer will appear in: installer_output\
; ============================================================

#define AppName      "Delta Dyeing PO Converter"
#define AppVersion   "1.0.0"
#define AppPublisher "Delta Dyeing S.A.E."
#define AppExeName   "DeltaDyeingConverter.exe"
#define SourceDir    "dist\DeltaDyeingConverter"
#define OutputDir    "installer_output"

[Setup]
; Unique app ID — do not change after first release
AppId={{6F3A2B1C-D4E5-4F60-A7B8-9C0D1E2F3A4B}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}

; Icon shown on the setup .exe itself, and used as the default for
; Add/Remove Programs when the installed app's own exe icon isn't used.
SetupIconFile=icon.ico

; Installation directory
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}

; Output
OutputDir={#OutputDir}
OutputBaseFilename=DeltaDyeingConverter_Setup_v{#AppVersion}

; Compression
Compression=lzma2/ultra64
SolidCompression=yes
CompressionThreads=auto

; Appearance
WizardStyle=modern
WizardResizable=no
DisableProgramGroupPage=yes

; Privileges — installs per-user (no admin required)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; Architecture
ArchitecturesInstallIn64BitMode=x64compatible

; No uninstaller needed? Set to "no" if you want to keep it simple
UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\{#AppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; \
    Description: "Create a desktop shortcut"; \
    GroupDescription: "Additional shortcuts:"

[Files]
; Copy the entire PyInstaller output folder into {app}
Source: "{#SourceDir}\*"; \
    DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Start Menu
Name: "{group}\{#AppName}"; \
    Filename: "{app}\{#AppExeName}"; \
    WorkingDir: "{app}"

Name: "{group}\Uninstall {#AppName}"; \
    Filename: "{uninstallexe}"

; Desktop shortcut (optional — user can deselect)
Name: "{commondesktop}\{#AppName}"; \
    Filename: "{app}\{#AppExeName}"; \
    WorkingDir: "{app}"; \
    Tasks: desktopicon

[Run]
; Offer to launch the app after install
Filename: "{app}\{#AppExeName}"; \
    Description: "Launch {#AppName}"; \
    Flags: nowait postinstall skipifsilent; \
    WorkingDir: "{app}"
