; ===========================================================================
;  药剂仓库管理系统 —— Inno Setup 安装脚本
;
;  ⚠️ 不要直接用 ISCC 编译本文件。走构建脚本：
;         python tools\build_installer.py
;     它会依次做：PyInstaller 打包 → 补 GPL 文件 → 重算完整性哈希 →
;     生成 version.iss → 调用 ISCC。少了前面任何一步，装出来的包都是坏的
;     （最典型的是 config.ini 里没有哈希，完整性校验形同虚设）。
;
;  ── 为什么是整目录安装（onedir）而不是装单个 exe ──
;  QtWebEngine 的辅助进程 QtWebEngineProcess.exe 必须相对主程序有可预测的
;  路径。onefile 每次启动解压到 _MEIxxxx 临时目录，辅助进程会找不到，
;  报 "Could not find QtWebEngineProcess.exe"。所以产物是一整个目录。
;
;  ── 关于 AppId ──
;  这里**故意不写** AppId。不写时 Inno 用 AppName 作为 AppId，而 AppName
;  与旧版安装包保持一致，于是新版能识别出"已安装旧版"并原地升级，
;  而不是并排装两份。写了 AppId 反而会把它变成另一个应用。
;  同理，DefaultDirName 也和旧版保持一致。
;
;  ── 用户数据 ──
;  数据库在 %APPDATA%\ReagentManagementSystem\，安装、覆盖安装、卸载都
;  不碰它。卸载时会额外复制一份带时间戳的备份到同目录的 Backups\，
;  但**不删除原库** —— 卸载后重装仍能看到原有库存。
; ===========================================================================

#include "version.iss"

; ── 中文语言包 ────────────────────────────────────────────────────────────
; Inno Setup 自带 28 个语言包，**不含简体中文** —— 中文属于非官方翻译，
; 要单独下载。文件缺失时 MessagesFile 指向不存在的文件，ISCC 是**直接
; 编译失败**而不是给个警告，所以这里必须能探测、能退化。
;
;   1. 仓库里的 packaging\ChineseSimplified.isl  ← 放这里就会进 git
;   2. ISCC 安装目录的 Languages\
;   3. 都没有 → 不定义该宏 → 退回英文界面，构建照常完成
;
; 想换别语言：把 .isl 放进 packaging\ 改下面两处文件名即可。
#if FileExists(AddBackslash(SourcePath) + "packaging\ChineseSimplified.isl")
  #define ChineseMessagesFile "packaging\ChineseSimplified.isl"
#elif FileExists(AddBackslash(CompilerPath) + "Languages\ChineseSimplified.isl")
  #define ChineseMessagesFile "compiler:Languages\ChineseSimplified.isl"
#endif

; ── Windows 7 上的硬性前提：KB2533623 ────────────────────────────────────
;
; Python 3.8 在 Win7 上**必须**有 KB2533623（或其替代更新 KB3063858）。
;
; 原因：CPython 3.8 出于安全改了扩展模块的加载方式，内部用
; LoadLibraryExW 带 LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR 这个标志去加载 .pyd。
; 这组标志正是 KB2533623 引入的。缺了它，LoadLibraryExW 返回 0、
; GetLastError = ERROR_INVALID_PARAMETER(87)，于是**所有 .pyd 都加载不了**：
;
;     ImportError: DLL load failed while importing _ctypes: 参数错误。
;
; 然后 ctypes 没了、_socket/_sqlite3/PyQt5.* 全没了，程序死在启动第一行。
;
; ⚠️ 这不是能靠代码绕过去的东西 —— 出问题的代码在 python38.dll 内部，
;    而 app/compat.py 想兜底，自己却要用 ctypes。别再去写兼容层了。
;
; 好消息是这个更新 2011 年就发布了，任何正常打过 Windows Update 的 Win7
; 都有；只有全新装完 SP1 就没再更新过的机器才缺。
#define KbPackageName "Windows6.1-KB3063858-x64.msu"
#if FileExists(AddBackslash(SourcePath) + "Dependencies\" + KbPackageName)
  #define HasKbPackage
#endif

#define MyAppName "ReagentManagementSystem"
#define MyAppDisplayName "药剂仓库管理系统"
#define MyAppExeName "ReagentManagementSystem.exe"
#define MyDistDir "Dist\ReagentManagementSystem"
#define DataDirName "ReagentManagementSystem"


[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppDisplayName} {#MyAppVersion}
UninstallDisplayName={#MyAppDisplayName} {#MyAppVersion}

; 这四个是 Windows「程序和功能」里显示的属性
VersionInfoVersion={#MyAppVersionQuad}
VersionInfoTextVersion={#MyAppVersion}
VersionInfoProductName={#MyAppDisplayName}
VersionInfoProductVersion={#MyAppVersion}
VersionInfoDescription={#MyAppDisplayName} 安装程序
VersionInfoCompany=GWJ

DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppDisplayName}
DisableProgramGroupPage=yes
DisableDirPage=auto
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallFilesDir={app}\Uninstall

OutputDir=Output
OutputBaseFilename=药剂仓库{#MyAppVersion}安装器
SetupIconFile=reagent.ico

; lzma2 而不是 lzma2/max：这一包是 150MB 的 Chromium DLL，已经压得很紧，
; max 只能再省几个百分点，却要多花好几倍时间和内存。不划算。
Compression=lzma2
SolidCompression=yes
LZMAUseSeparateProcess=yes

; 需要管理员：装到 Program Files，且要静默安装 VC++ 运行库
PrivilegesRequired=admin
; 只支持 64 位。PyQtWebEngine-Qt5 在 Windows 上只有 64 位轮子，
; 出不了 32 位版本，别去开这个口子。
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
; 6.1sp1 = Windows 7 SP1。低于这个版本的机器直接拒绝安装，
; 好过装完打不开再来排查。
MinVersion=6.1sp1

WizardStyle=modern
; 升级时必须强制关掉正在运行的程序：QtWebEngineProcess.exe 会占住
; Qt5WebEngineCore.dll，不关掉就会留下新旧混装的目录。
CloseApplications=force
RestartApplications=no
; 防止更新程序调起的安装和自己手动双击的撞在一起
SetupMutex=ReagentManagementSystemSetup
SetupLogging=yes


[Languages]
; 语言包怎么找到的见文件开头的 ChineseMessagesFile 说明。
#ifdef ChineseMessagesFile
Name: "ChineseSimplified"; MessagesFile: "{#ChineseMessagesFile}"
#else
Name: "English"; MessagesFile: "compiler:Default.isl"
#endif


[Files]
; onedir 产物整目录分发。
;
; ⚠️ 这里**不要**再单独 Source 一次仓库根目录的 config.ini —— 那会把产物里
;    已经写好完整性哈希的那份覆盖掉，校验就永久失效了。产物里的 config.ini
;    已经在这一条里。
Source: "{#MyDistDir}\*"; DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

; GPL v3 要求随二进制一起提供许可证与源码线索。
; 由构建脚本放进产物目录，所以上面那条已经覆盖；这两行是保险，
; 万一有人手工删了产物里的副本，至少不会漏掉合规文件。
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "SOURCE_CODE"; DestDir: "{app}"; Flags: ignoreversion

; VC++ 运行库。主程序自己带了 app-local 的运行库副本，所以这个不是硬需求，
; 装上更稳妥（尤其是系统里没有 UCRT 的老 Win7）。
Source: "Dependencies\VC_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

#ifdef HasKbPackage
; Win7 缺 KB2533623 时的救命更新。装完**必须重启**才生效 ——
; kernel32.dll 正在被使用，Windows 换不掉它，新起的进程照样拿到旧的。
Source: "Dependencies\{#KbPackageName}"; DestDir: "{tmp}"; Flags: deleteafterinstall
#endif


[Icons]
; 用 {common*} 而不是 {auto*}：前者在 Inno 5/6 上都能编译，
; 且 PrivilegesRequired=admin 时两者等价。
Name: "{commondesktop}\{#MyAppDisplayName}"; Filename: "{app}\{#MyAppExeName}"; \
    WorkingDir: "{app}"; Comment: "{#MyAppDisplayName} {#MyAppVersion}"
Name: "{commonprograms}\{#MyAppDisplayName}"; Filename: "{app}\{#MyAppExeName}"; \
    WorkingDir: "{app}"


[Run]
#ifdef HasKbPackage
; Windows 7 缺 KB2533623 时自动补上。放在 VC_redist 之前 ——
; 缺了它什么都跑不了，先解决这个。
;
; wusa.exe 是 Windows 自带的独立更新安装器。已经装过时它会返回
; "已安装"错误码，[Run] 不看子进程退出码，无副作用。
Filename: "{sys}\wusa.exe"; \
    Parameters: """{tmp}\{#KbPackageName}"" /quiet /norestart"; \
    StatusMsg: "正在安装 Windows 7 必需更新（KB3063858）..."; \
    Flags: waituntilterminated skipifdoesntexist; \
    Check: KbUpdateNeeded
#endif

; 装运行库。
;
;   waituntilterminated —— 等它装完再往下走。不等的话用户可能在运行库还在
;                          装的时候就点了启动程序。
;   skipifdoesntexist   —— 解出来的文件不在就跳过，别弹报错。
;
; ⚠️ [Run] **没有** ignoreerrors 这个 flag（那是 [Files]/[Registry] 的），
;    写上会直接编译失败。好在 [Run] 本来就不检查子进程的退出码 ——
;    只有"启动不起来"才会报错，而那种情况由 skipifdoesntexist 挡掉。
;    所以 VC_redist 静默安装失败不会中断整个安装流程，
;    这对本项目是好事：主程序带了自己的运行库副本，实测不装也能跑。
Filename: "{tmp}\VC_redist.x64.exe"; Parameters: "/install /quiet /norestart"; \
    StatusMsg: "正在安装 Visual C++ 运行库..."; \
    Flags: waituntilterminated skipifdoesntexist; \
    Check: VCRedistNeeded


[InstallDelete]
; ⚠️ 升级时必须清掉旧版本留下的 MSVC 运行库副本。
;
; 新版本刻意不再分发这些文件（原因见 build.spec：打包机会带进过新的版本，
; 而新版 MSVCP140.dll 静态导入 Win8 才有的 GetSystemTimePreciseAsFileTime，
; 在 Win7 上加载即失败）。
;
; 但 Inno **不会自动删除**已存在、而新版本不再提供的文件 —— 覆盖安装之后，
; 那个坏的副本还留在程序目录里，照样盖住系统里那份好的，等于没修。
;
; [InstallDelete] 在 [Files] **之前**处理，所以这里是先删后拷，顺序正确。
; 删掉之后主程序会去 System32 找运行库，那里由 VC_redist 装的是正确的 14.44。
Type: files; Name: "{app}\MSVCP140.dll"
Type: files; Name: "{app}\MSVCP140_1.dll"
Type: files; Name: "{app}\MSVCP140_2.dll"
Type: files; Name: "{app}\MSVCP140_ATOMIC_WAIT.dll"
Type: files; Name: "{app}\MSVCP140_CODECVT_IDS.dll"
Type: files; Name: "{app}\VCRUNTIME140.dll"
Type: files; Name: "{app}\VCRUNTIME140_1.dll"
Type: files; Name: "{app}\CONCRT140.dll"
Type: files; Name: "{app}\VCCORLIB140.dll"


[UninstallDelete]
; 清掉整个安装目录，包括程序写出的 startup-error.txt 之类的散落文件
Type: filesandordirs; Name: "{app}"


[Code]
{ --------------------------------------------------------------------------
  卸载前备份用户数据库。

  只复制、不删除原库 —— 卸载后重装仍然能看到原有库存。
  （旧版用的 uninstall_backup.exe 会在备份后删掉原库，那个行为已废弃。）
  -------------------------------------------------------------------------- }

{ --------------------------------------------------------------------------
  Windows 7 前提检查：LoadLibraryExW 认不认那组搜索标志。

  背景见文件开头 KB2533623 的说明。这里用 A 版本而不是 W 版本，是因为
  Pascal Script 的 String 参数按 PAnsiChar 传递；A/W 对 dwFlags 的校验
  逻辑完全一样，用 A 版本省掉宽字符转换的坑。
  -------------------------------------------------------------------------- }

function LoadLibraryExA(lpLibFileName: String; hFile: THandle; dwFlags: LongWord): THandle;
  external 'LoadLibraryExA@kernel32.dll stdcall';
function FreeLibrary(hLibModule: THandle): Boolean;
  external 'FreeLibrary@kernel32.dll stdcall';

const
  { LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS }
  LIBRARY_SEARCH_FLAGS = $00001100;

function HasLibrarySearchFlags(): Boolean;
var
  H: THandle;
begin
  { 拿 kernel32.dll 自己当探针：必定存在，必定可加载。
    缺 KB2533623 时这两个标志位不被识别，调用直接失败返回 0。 }
  H := LoadLibraryExA(ExpandConstant('{sys}\kernel32.dll'), 0, LIBRARY_SEARCH_FLAGS);
  Result := H <> 0;
  if Result then
    FreeLibrary(H);
end;

function KbPackageCompiled(): Boolean;
begin
#ifdef HasKbPackage
  Result := True;
#else
  Result := False;
#endif
end;

function KbUpdateNeeded(): Boolean;
begin
  Result := (not HasLibrarySearchFlags()) and KbPackageCompiled();
end;

function InitializeSetup(): Boolean;
begin
  Result := True;

  if HasLibrarySearchFlags() then
    Exit;   { 前提满足，正常安装 }

  if KbPackageCompiled() then
    Result := MsgBox(
      '本机缺少 Windows 7 的必需更新（KB2533623 或 KB3063858）。' + #13#10 + #13#10 +
      '没有它，Python 3.8 加载不了任何扩展模块，本程序装上也启动不了。' + #13#10 + #13#10 +
      '安装程序会自动补上这个更新，但装完必须重启电脑才生效' + #13#10 +
      '（Windows 换不掉正在使用中的 kernel32.dll）。' + #13#10 + #13#10 +
      '现在继续安装吗？',
      mbConfirmation, MB_YESNO) = IDYES
  else
    Result := MsgBox(
      '本机缺少 Windows 7 的必需更新（KB2533623 或 KB3063858）。' + #13#10 + #13#10 +
      '没有它，Python 3.8 加载不了任何扩展模块，本程序装上也无法启动。' + #13#10 + #13#10 +
      '这个安装包没有附带该更新，需要你自己补：' + #13#10 +
      '    1. 让 Windows Update 把系统更新到最新（推荐）' + #13#10 +
      '    2. 或者单独安装 KB3063858' + #13#10 + #13#10 +
      '装好并重启之后再运行本程序即可。' + #13#10 + #13#10 +
      '仍然继续安装吗？',
      mbConfirmation, MB_YESNO) = IDYES;

  if not Result then
    MsgBox('已取消安装。', mbInformation, MB_OK);
end;

procedure WarnIfKbNeedsReboot();
begin
  if not KbPackageCompiled() then
    Exit;

  { 装完还是探测不到 —— 说明更新只是被"暂存"，要重启才会生效。 }
  if HasLibrarySearchFlags() then
    Exit;

  { 静默安装（程序内的自动更新走这条路）时不要弹窗卡住流程 }
  if WizardSilent() then
    Exit;

  MsgBox('Windows 更新已经装好，但需要重启电脑才能生效。' + #13#10 + #13#10 +
         '重启之前启动本程序，仍然会报' + #13#10 +
         '    DLL load failed while importing _ctypes' + #13#10 + #13#10 +
         '请重启后再运行。',
         mbInformation, MB_OK);
end;

{ 主程序不再自带 MSVC 运行库了（原因见 build.spec），四个文件全依赖系统。
  缺一个都起不来，所以这两个函数要按同一套条件判断。

  ⚠️ 定义顺序有讲究：Pascal Script 要求**先声明后使用**，所以这两个函数
  必须放在 CurStepChanged / WarnIfRuntimeMissing 之前，否则编译报
  "Unknown identifier"。 }

function SystemRuntimeMissing(): Boolean;
begin
  Result := (not FileExists(ExpandConstant('{sys}\vcruntime140.dll')))
         or (not FileExists(ExpandConstant('{sys}\vcruntime140_1.dll')))
         or (not FileExists(ExpandConstant('{sys}\msvcp140.dll')))
         or (not FileExists(ExpandConstant('{sys}\msvcp140_1.dll')));
end;

function VCRedistNeeded(): Boolean;
begin
  { 系统缺运行库、且包里带了 VC_redist，就去装。
    以前只查 vcruntime140.dll 一个是不够的 —— 比如系统里只有老版运行库、
    没有 msvcp140_1.dll 时，检查会通过，但程序照样启动不了。

    另一条 FileExists 和 [Run] 的 skipifdoesntexist 重复，是故意的：
    Check 为假时整个条目都不执行，比依赖 flag 更直观。 }
  Result := SystemRuntimeMissing()
            and FileExists(ExpandConstant('{tmp}\VC_redist.x64.exe'));
end;

procedure WarnIfRuntimeMissing();
begin
  if not SystemRuntimeMissing() then
    Exit;
  if WizardSilent() then
    Exit;

  { 走到这里说明 VC_redist 要么没跑要么没装上，而主程序现在不自带运行库，
    结果会是双击没反应或弹「找不到 MSVCP140.dll」—— 用户完全不知道怎么办。 }
  MsgBox('Visual C++ 运行库没能装上，本程序无法启动。' + #13#10 + #13#10 +
         '请手动安装随包的运行库后重试：' + #13#10 +
         '    Dependencies\VC_redist.x64.exe' + #13#10 + #13#10 +
         '如果它报错装不上，请在系统里运行 Windows Update 更新到最新。',
         mbError, MB_OK);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  { ssDone = 全部 [Run] 条目都跑完之后。放在这里而不是 ssPostInstall，
    是为了不依赖 [Run] 与那个步骤的先后关系。 }
  if CurStep = ssDone then
  begin
    WarnIfKbNeedsReboot();
    WarnIfRuntimeMissing();
  end;
end;


function GetUserDataDir(): String;
begin
  Result := ExpandConstant('{userappdata}') + '\' + '{#DataDirName}';
end;

procedure BackupUserData();
var
  DbFile, BackupDir, Target, Stamp: String;
begin
  DbFile := GetUserDataDir() + '\ReagentWarehouseData.db';

  if not FileExists(DbFile) then
  begin
    Log('未找到用户数据库，跳过备份: ' + DbFile);
    Exit;
  end;

  BackupDir := GetUserDataDir() + '\Backups';
  if not ForceDirectories(BackupDir) then
  begin
    MsgBox('无法创建备份目录，数据库未备份：' + #13#10 + BackupDir + #13#10 + #13#10 +
           '如果需要保留数据，请在卸载后手动复制：' + #13#10 + DbFile,
           mbError, MB_OK);
    Exit;
  end;

  Stamp := GetDateTimeString('yyyymmdd_hhnnss', #0, #0);
  Target := BackupDir + '\ReagentWarehouseData_' + Stamp + '.db';

  if FileCopy(DbFile, Target, False) then
    Log('数据库已备份到: ' + Target)
  else
    MsgBox('数据库备份失败，原文件未受影响：' + #13#10 + DbFile + #13#10 + #13#10 +
           '请在卸载后手动复制该文件。',
           mbError, MB_OK);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  { usUninstall = 开始真正删除文件之前。放在这里是为了保证
    用户数据在程序目录被清掉之前就已经备份好。 }
  if CurUninstallStep = usUninstall then
    BackupUserData();
end;
