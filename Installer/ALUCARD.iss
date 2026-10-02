#define AddinGuid "{7E8D4B8A-3D52-4B74-9B7A-5F4E8C2D1A61}"

[Setup]
AppId={#AddinGuid}
AppName=ALUCARD
AppVersion=1.0.0
AppPublisher=ALUCARD
DefaultDirName={autopf}\ALUCARD
DefaultGroupName=ALUCARD
UninstallDisplayName=ALUCARD
OutputDir=..\Release
OutputBaseFilename=ALUCARD-Setup
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Files]
Source: "..\Addin\bin\Release\net48\ALUCARD.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\SolidWorks\VBA\ALUCARD.bas"; DestDir: "{app}\SolidWorks\VBA"; Flags: ignoreversion
Source: "..\SolidWorks\VBA\ALUCARD.swp"; DestDir: "{app}\SolidWorks\VBA"; Flags: ignoreversion

[Registry]
Root: HKLM64; Subkey: "Software\SolidWorks\AddIns\{#AddinGuid}"; ValueType: dword; ValueName: "LoadAtStartup"; ValueData: "0"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\SolidWorks\AddIns\{#AddinGuid}"; ValueType: string; ValueName: "Title"; ValueData: "ALUCARD"
Root: HKLM64; Subkey: "Software\SolidWorks\AddIns\{#AddinGuid}"; ValueType: string; ValueName: "Description"; ValueData: "ALUCARD SolidWorks DXF Automation"

Root: HKLM64; Subkey: "Software\Classes\CLSID\{#AddinGuid}"; ValueType: string; ValueName: ""; ValueData: "ALUCARD.SwAddin"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Classes\CLSID\{#AddinGuid}\ProgID"; ValueType: string; ValueName: ""; ValueData: "ALUCARD.SwAddin"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Classes\CLSID\{#AddinGuid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{#AddinGuid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{#AddinGuid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "ALUCARD.SwAddin"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{#AddinGuid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "ALUCARD, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{#AddinGuid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{#AddinGuid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{app}\ALUCARD.dll"

Root: HKLM64; Subkey: "Software\Classes\ALUCARD.SwAddin\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#AddinGuid}"; Flags: uninsdeletekey