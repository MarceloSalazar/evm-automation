; Instalador de Windows del Exportador EVM (Inno Setup 6).
; Se compila con instalador\construir.py, que define AppVersion (versión del programa) y Origen
; (carpeta generada por PyInstaller con evm_export.exe y generar_ejemplo.exe).

#ifndef AppVersion
  #error Falta /DAppVersion=X.Y.Z: el instalador se genera con instalador\construir.py
#endif
#ifndef Origen
  #error Falta /DOrigen=<carpeta generada por PyInstaller>: usar instalador\construir.py
#endif
#define Raiz AddBackslash(SourcePath) + ".."
#define Repositorio "https://github.com/MarceloSalazar/evm-automation"

[Setup]
AppId={{BC60A1FE-1670-457C-B2B0-F72DD6364777}
AppName=Exportador EVM
AppVersion={#AppVersion}
AppVerName=Exportador EVM {#AppVersion}
AppPublisher=Marcelo Salazar
AppPublisherURL={#Repositorio}
AppSupportURL={#Repositorio}/issues
AppUpdatesURL={#Repositorio}/releases
AppCopyright=Copyright (c) 2026 Marcelo Salazar. Licencia MIT.
VersionInfoVersion={#AppVersion}.0
VersionInfoDescription=Instalador del Exportador EVM
; Instalación para el usuario actual, sin permisos de administrador (%LOCALAPPDATA%\Programs)
PrivilegesRequired=lowest
DefaultDirName={autopf}\Exportador EVM
DefaultGroupName=Exportador EVM
DisableProgramGroupPage=yes
AllowNoIcons=yes
LicenseFile={#Raiz}\LICENSE
OutputBaseFilename=ExportadorEVM-{#AppVersion}-instalador
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
ChangesEnvironment=yes
UninstallDisplayName=Exportador EVM {#AppVersion}
UninstallDisplayIcon={app}\evm_export.exe

[Languages]
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "path"; Description: "Agregar evm_export al PATH del usuario, para usarlo desde cualquier consola"

[InstallDelete]
; Al actualizar, se eliminan las bibliotecas de la versión anterior antes de copiar las nuevas
Type: filesandordirs; Name: "{app}\_internal"

[Files]
Source: "{#Origen}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "consola.cmd"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#Raiz}\config.ejemplo.toml"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#Raiz}\ejemplos\programa_ejemplo.xml"; DestDir: "{app}\ejemplos"; Flags: ignoreversion
Source: "{#Raiz}\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion
Source: "{#Raiz}\CHANGELOG.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Exportador EVM (consola)"; Filename: "{app}\consola.cmd"; WorkingDir: "{userdocs}"; IconFilename: "{app}\evm_export.exe"
Name: "{group}\Documentación"; Filename: "{#Repositorio}/blob/main/docs/README.md"
Name: "{group}\Desinstalar Exportador EVM"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\consola.cmd"; WorkingDir: "{userdocs}"; Description: "Abrir la consola del Exportador EVM"; Flags: postinstall nowait skipifsilent shellexec

[Code]
const
  EnvKey = 'Environment';

{ ';' + lista + ';' en mayúsculas, para buscar una carpeta completa dentro del PATH }
function Wrapped(Paths: string): string;
begin
  Result := ';' + Uppercase(Paths) + ';';
end;

procedure AddToPath(Dir: string);
var
  Paths: string;
begin
  if not RegQueryStringValue(HKCU, EnvKey, 'Path', Paths) then
    Paths := '';
  if Pos(Wrapped(Dir), Wrapped(Paths)) > 0 then
    exit;
  { Si el PATH termina en ';', se conserva al final para que la desinstalación lo deje como estaba }
  if Paths = '' then
    Paths := Dir
  else if Copy(Paths, Length(Paths), 1) = ';' then
    Paths := Paths + Dir + ';'
  else
    Paths := Paths + ';' + Dir;
  RegWriteExpandStringValue(HKCU, EnvKey, 'Path', Paths);
end;

procedure RemoveFromPath(Dir: string);
var
  Paths, S: string;
  P: Integer;
begin
  if not RegQueryStringValue(HKCU, EnvKey, 'Path', Paths) then
    exit;
  S := ';' + Paths + ';';
  P := Pos(Wrapped(Dir), Uppercase(S));
  if P = 0 then
    exit;
  Delete(S, P, Length(Dir) + 1);
  RegWriteExpandStringValue(HKCU, EnvKey, 'Path', Copy(S, 2, Length(S) - 2));
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep = ssPostInstall) and WizardIsTaskSelected('path') then
    AddToPath(ExpandConstant('{app}'));
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
    RemoveFromPath(ExpandConstant('{app}'));
end;
