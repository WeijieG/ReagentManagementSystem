; -- 基础设置 --
[Setup]
AppName=药剂仓库
AppVersion=1.0
DefaultDirName={autopf}\药剂仓库
OutputDir=.\Output
OutputBaseFilename=药剂仓库安装(正式开源版)
SetupIconFile=.\reagent.ico
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
UninstallFilesDir={app}\Uninstall
VersionInfoVersion=1.0.0.0
VersionInfoTextVersion=1.0.0

[Languages]
Name: "ChineseSimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

; -- 系统检测 --
; [Code]
// function IsX64: Boolean;
// begin
  // Result := Is64BitInstallMode and (ProcessorArchitecture = paX64);
// end;

[Files]
; 根据系统架构安装对应exe

Source: ".\Dist\x64\ReagentManagementSystem_x64.exe"; DestDir: "{app}"; DestName: "ReagentManagementSystem.exe"; Flags: ignoreversion


; Source: ".\Dist\x64\ReagentManagementSystem_x64.exe"; DestDir: "{app}"; DestName: "ReagentManagementSystem.exe"; Check: IsX64; Flags: ignoreversion
; Source: ".\Dist\x86\ReagentManagementSystem_x86.exe"; DestDir: "{app}"; DestName: "ReagentManagementSystem.exe"; Check: not IsX64; Flags: ignoreversion
Source: ".\reagent.ico"; DestName: "reagent.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: ".\config.ini"; DestDir: "{app}"; Flags: ignoreversion
Source: ".\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: ".\SOURCE_CODE"; DestDir: "{app}"; Flags: ignoreversion
Source: ".\images\*"; DestDir: "{app}\images"; Flags: ignoreversion recursesubdirs createallsubdirs

;数据库升级脚本打包
; Source: "database_migration.exe"; DestDir: "{app}"; Flags: ignoreversion

; 卸载备份脚本
Source: "uninstall_backup_64.exe"; DestDir: "{app}"; DestName: "uninstall_backup.exe"; Flags: ignoreversion

; Source: "uninstall_backup_64.exe"; DestDir: "{app}"; DestName: "uninstall_backup.exe"; Check: IsX64; Flags: ignoreversion
; Source: "uninstall_backup_86.exe"; DestDir: "{app}"; DestName: "uninstall_backup.exe"; Check: not IsX64; Flags: ignoreversion


; 依赖文件 - 确保包含所有运行库
Source: ".\Dependencies\VC_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall;

; Source: ".\Dependencies\VC_redist.x86.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: not IsX64
; Source: ".\Dependencies\VC_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: IsX64

; 公共文件（如有）
; Source: ".\Dist\Common\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

; -- 快捷方式 --
[Icons]
Name: "{commondesktop}\ReagentManagementSystem"; Filename: "{app}\ReagentManagementSystem.exe"
Name: "{commonprograms}\ReagentManagementSystem"; Filename: "{app}\ReagentManagementSystem.exe"

; -- 依赖安装（示例：VC++ 2015-2022运行库）--
[Run]
Filename: "{tmp}\VC_redist.x64.exe"; Parameters: "/install /quiet /norestart"; \
    StatusMsg: "正在安装运行库..."; Check: FileExists(ExpandConstant('{tmp}\VC_redist.x64.exe'))
    
; Filename: "{tmp}\VC_redist.x86.exe"; Parameters: "/install /quiet /norestart"; \
    ; StatusMsg: "正在安装运行库..."; Check: (not IsX64) and FileExists(ExpandConstant('{tmp}\VC_redist.x86.exe'))
; Filename: "{tmp}\VC_redist.x64.exe"; Parameters: "/install /quiet /norestart"; \
    ; StatusMsg: "正在安装运行库..."; Check: (IsX64) and FileExists(ExpandConstant('{tmp}\VC_redist.x64.exe'))
    
;执行数据库结构变更脚本
; Filename: "{app}\database_migration.exe"; Description: "升级数据库结构"; Flags: postinstall runhidden
    
[UninstallRun]
; 在卸载前执行备份
Filename: "{app}\uninstall_backup.exe"; 

[UninstallDelete]
; 删除应用程序目录
Type: filesandordirs; Name: "{app}"