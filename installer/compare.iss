; Compare — Inno Setup installer script.
;
; Requirements:
;   - Inno Setup 6.x installed on the build machine.
;   - `cargo build --release` (CLI) and `cargo build -p compare-gui --release`
;     completed, so the target binaries exist.
;
; Usage:
;   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\compare.iss
;
; Produces: installer\Output\Compare-Setup-<Version>.exe

#define AppName       "Compare"
#define AppVersion    "0.1.0"
#define AppPublisher  "Compare Project"
#define AppExeGui     "compare-gui.exe"
#define AppExeCli     "compare.exe"

[Setup]
AppId={{8F1F7B91-2C2E-4E8D-9A31-1F1B5A7A4D23}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=Compare-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#AppExeGui}
MinVersion=10.0
LicenseFile=
SetupIconFile=..\crates\gui\icons\icon.ico

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면 바로가기 만들기"; GroupDescription: "추가 작업:"; Flags: unchecked
Name: "addtopath"; Description: "명령 프롬프트에서 compare 명령 사용 (PATH 추가)"; GroupDescription: "추가 작업:"; Flags: unchecked

[Files]
Source: "..\target\release\{#AppExeGui}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\target\release\{#AppExeCli}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\crates\gui\icons\icon.ico";   DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}";       Filename: "{app}\{#AppExeGui}"
Name: "{group}\제거";             Filename: "{uninstallexe}"
Name: "{commondesktop}\{#AppName}"; Filename: "{app}\{#AppExeGui}"; Tasks: desktopicon

[Registry]
; Optional: add install dir to user PATH when "addtopath" task is selected.
Root: HKCU; Subkey: "Environment"; ValueType: expandsz; ValueName: "Path"; \
  ValueData: "{olddata};{app}"; Check: NeedsAddPath('{app}'); Tasks: addtopath

[Run]
Filename: "{app}\{#AppExeGui}"; Description: "지금 {#AppName} 실행"; Flags: nowait postinstall skipifsilent

[Code]
function NeedsAddPath(Param: string): boolean;
var
  OrigPath: string;
begin
  if not RegQueryStringValue(HKEY_CURRENT_USER, 'Environment', 'Path', OrigPath)
  then begin
    Result := True;
    exit;
  end;
  Result := Pos(';' + Param + ';', ';' + OrigPath + ';') = 0;
end;
