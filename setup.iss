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
var
  { 安装**开始之前**的环境状态。

    用来区分「本机本来就有」和「本次装上了」这两种情况 —— 它们在安装
    结束时探测到的结果是一样的（都还探测不到，因为要重启才生效），
    只靠装完再探测一次分不出来，所以要在装之前先取一次基线。

    这个区分决定了要不要提示用户重启：只有"本次真的装了"才提示。
    见 CurStepChanged(ssInstall) 和 ShowFinalNotice()。 }
  KbWasMissing: Boolean;


{ --------------------------------------------------------------------------
  Windows 7 前提检查：LoadLibraryEx 认不认那组搜索标志。

  背景见文件开头 KB2533623 的说明。

  ⚠️ 必须用 **W** 版本，不能用 A。Inno Setup 6 的 String 是 Unicode，传给
     外部函数时按 **PWideChar** 走；调 A 版本的话对方按 ANSI 读，拿到的是
     乱码路径，**必然加载失败**。实测数据（同一台机器上四次调用）：

         LoadLibraryExA(无标志):  0            err=126
         LoadLibraryExW(无标志):  1975189504   err=126
         LoadLibraryExA(带标志):  0            err=87
         LoadLibraryExW(带标志):  1975189504   err=87

     A 连"不带标志"都返回 0，说明失败在路径编码，跟标志位无关。
     用 A 的后果非常隐蔽：探测永远报"缺环境"，于是**每次安装都会提示重启、
     每次都会白跑一遍 wusa**，而程序其实好好的 —— 用户看到的就是"第二次装
     还让我重启，可不重启也能打开"。

  判定看**返回值**，不看 GetLastError —— LoadLibraryEx 成功时不清 last error，
     句柄非零时 err 可能是上一次调用留下的（上面 err=87 就是这种情况）。
  -------------------------------------------------------------------------- }

function LoadLibraryExW(lpLibFileName: String; hFile: THandle; dwFlags: LongWord): THandle;
  external 'LoadLibraryExW@kernel32.dll stdcall';
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
    缺 KB2533623 时这两个标志位不被识别，调用失败返回 0。

    ⚠️ 这里必须是 W 版本。用 A 版本的话，Inno 传的是 UTF-16 字符串、
    A 函数按 ANSI 读 —— 路径成了乱码，**不管有没有装 KB 都返回 0**，
    探测就彻底失效了。原因见上面的说明。 }
  H := LoadLibraryExW(ExpandConstant('{sys}\kernel32.dll'), 0, LIBRARY_SEARCH_FLAGS);
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

{ 主程序不再自带 MSVC 运行库（原因见 build.spec），四个文件全依赖系统，
  缺任何一个都起不来。

  ⚠️ 定义顺序有讲究：Pascal Script 要求**先声明后使用**。这个函数必须放在
  ShowFinalNotice 之前，否则编译报 "Unknown identifier"。 }

function SystemRuntimeMissing(): Boolean;
begin
  Result := (not FileExists(ExpandConstant('{sys}\vcruntime140.dll')))
         or (not FileExists(ExpandConstant('{sys}\vcruntime140_1.dll')))
         or (not FileExists(ExpandConstant('{sys}\msvcp140.dll')))
         or (not FileExists(ExpandConstant('{sys}\msvcp140_1.dll')));
end;


{ --------------------------------------------------------------------------
  安装结束时的唯一一处提示。

  设计原则：**不问用户问题，也不讲系统补丁的术语；没做事就不出声。**

  需要补的东西在 [Run] 里已经静默装好了，这里只负责在"本次真的动了系统、
  且需要重启"时说一声。

  ⚠️ 判断依据是 **KbWasMissing（安装前的基线）**，不是装完再探测一次 ——
  因为更新装了但要重启才生效，装完探测到的结果和"本来就没装"是一样的，
  分不出来。只看装完的结果会把「本机本来就有、跳过安装」也误报成需要重启。

  各分支：
    · 本机本来就有环境      → 什么都没装      → 静默，不提示重启
    · 本来缺、装上且已生效  → 不需要重启      → 静默
    · 本来缺、装上但没生效  → 需要重启        → 提示重启
    · 本来缺、包里没带/装失败 → 走 Windows Update
    · 运行库缺失            → 硬故障，优先报
  -------------------------------------------------------------------------- }

procedure ShowFinalNotice();
begin
  { 静默安装（程序内的自动更新走这条路）不弹窗，否则会卡住更新流程 }
  if WizardSilent() then
    Exit;

  { 一、运行库缺了就是硬故障：主程序不再自带它，结果是双击没反应。
        无论安装前什么状态都要报出来。 }
  if SystemRuntimeMissing() then
  begin
    MsgBox('安装完成。' + #13#10 + #13#10 +
           '但 Visual C++ 运行库没能装上，程序暂时无法启动。' + #13#10 + #13#10 +
           '请把系统更新到最新（Windows Update），然后重启电脑再试。',
           mbError, MB_OK);
    Exit;
  end;

  { 二、本机本来就有这个环境 → 本次跳过了安装，什么都没动 → 不打扰 }
  if not KbWasMissing then
    Exit;

  { 三、本来缺、装完却已经生效 → 不需要重启 → 也不用提示 }
  if HasLibrarySearchFlags() then
    Exit;

  { 四、本来缺、装完仍缺：本次把它装上了，等重启才生效 }
  if KbPackageCompiled() then
    MsgBox('安装完成。' + #13#10 + #13#10 +
           '请重启电脑后再运行本程序。' + #13#10 + #13#10 +
           '本次安装补上了 Windows 7 的一个必需系统更新，' +
           '它需要重启才会生效。' + #13#10 +
           '（若重启后仍无法运行，请把系统更新到最新。）',
           mbInformation, MB_OK)
  else
    MsgBox('安装完成。' + #13#10 + #13#10 +
           '但本机缺少 Windows 7 的一个必需系统更新，程序暂时无法启动。' + #13#10 + #13#10 +
           '请把系统更新到最新（Windows Update），然后重启电脑再试。',
           mbError, MB_OK);
end;

{ VCRedistNeeded：系统缺运行库、且包里带了 VC_redist，就去装。
  主程序不再自带 MSVC 运行库（原因见 build.spec），四个文件全依赖系统。 }

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

procedure CurStepChanged(CurStep: TSetupStep);
begin
  { ssInstall = 开始复制文件、执行 [Run] **之前**。在这里取环境基线 ——
    必须早于 [Run]，否则装完了再取就分不清"本来就有"和"本次装上"了。 }
  if CurStep = ssInstall then
    KbWasMissing := not HasLibrarySearchFlags()
  else if CurStep = ssDone then
    { ssDone = 全部 [Run] 条目都跑完之后。放在这里而不是 ssPostInstall，
      是为了不依赖 [Run] 与那个步骤的先后关系。 }
    ShowFinalNotice();
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
