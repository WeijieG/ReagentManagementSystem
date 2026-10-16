# 运行与构建说明

---

## 0. 直接跑源码（开发测试推荐）

**开发阶段不要打包。** 源码模式下 `paths.IS_FROZEN` 为 False，打包相关的
DLL 搜索问题全都不存在，改完代码直接重跑即可。

### 装环境

```
Python 3.8.10（64 位，安装时勾选 Add Python 3.8 to PATH）
python -m pip install -r requirements.txt
```

### 跑起来

```
python main.py
```

数据库会建在源码目录下（`ReagentWarehouseData.db`），与正式环境隔离。

### 界面自检

不用手点，让它自己遍历所有页面并截图：

```
python tools/ui_smoke_test.py                  # 浅色
python tools/ui_smoke_test.py --theme dark     # 深色
```

截图写到 `tools/ui_smoke*.png`。它会检查五个页面与两个抽屉有没有正常渲染，
以及有没有掉进错误面板。

### 单元测试

```
python -m unittest discover -s tests           # 76 项
```

覆盖 GS1 解析（新旧实现逐键对照）、Excel 导出（黄金矩阵）、版本比较、
入库出库的边界行为、老库升级路径、完整性校验。

### 调试

```
set RMS_DEBUG=1        # 把页面里的 JS 报错打到控制台
python main.py
```

开发模式下窗口也支持 `RMS_SMOKE_SHOT` 环境变量：设定后启动完成即截图并退出，
方便自动化验证。

---

## 1. 出安装包（一条命令）

### 环境

| 项目 | 版本 | 说明 |
|---|---|---|
| Python | **3.8.10** 64 位 | 最后一个支持 Win7 的版本，不能更高 |
| PyInstaller | 4.10 | 见 `requirements-build.txt` |
| Inno Setup | 6 | 只需命令行编译器 ISCC，装在哪儿都行 |

```
python -m pip install -r requirements-build.txt
winget install -e --id JRSoftware.InnoSetup
```

### 构建

```
python tools\build_installer.py
```

产物：

```
Dist\ReagentManagementSystem\      ← 完整的 onedir 目录，自洽，可直接拷走
Output\药剂仓库2.1.0安装器.exe      ← 安装包
```

### 参数

| 参数 | 用途 |
|---|---|
| `--skip-build` | 跳过 PyInstaller，用现有 Dist 重打安装包（改过 `setup.iss` 后用） |
| `--no-installer` | 只出 Dist，不编译安装包（没装 Inno Setup 时也能用） |
| `--smoke` | 构建后启动一次产物并截图，确认界面**真的渲染出来**了 |
| `--version 2.2.0` | 改版本号（写回 `config.ini`）后构建 |

### 它做的事，以及为什么顺序不能乱

| 步骤 | 内容 | 乱了会怎样 |
|---|---|---|
| 1 | 检查 Python 3.8 / 64 位 / PyInstaller | 用高版本 Python 打的包在 Win7 上起不来 |
| 1 | 检查 `config.ini` 是生产配置 | 带着 `DataDirName=...Test` 发布，用户装上看到空库 |
| 1 | 检查 `server.py` 的 `CURRENT_VERSION` 是否同步 | 不同步则更新推送永远不生效 |
| 2 | 清理 `Dist` 与 `build\pyi` | 残留旧文件会混进新包 |
| 3 | PyInstaller 打包 | —— |
| 4 | 把 `LICENSE` / `SOURCE_CODE` 放进产物 + 产物自检 | 漏了是 GPL 合规问题 |
| 5 | 重算完整性哈希，写进**产物里**的 `config.ini` | 写错地方 → 校验静默失效 |
| 6 | （`--smoke`）启动产物截图 | —— |
| 7 | 生成 `version.iss` → 调用 ISCC | —— |

**第 5 步是这个脚本存在的主要理由。** `gen_hashes.py` 早期的 `--config`
默认写的是仓库根目录的 `config.ini`，而打进安装包的是产物目录里那份副本。
结果是：程序打得出来、安装包装得上、也能跑，但完整性校验永远是空的——
一个完全不报错的静默失效。现在默认写产物里那份，脚本还会复核哈希长度。

### 版本号

唯一来源是 `config.ini` 的 `[Application] version`。构建脚本读它，生成
`version.iss` 给 `setup.iss` 用。

**发布更新时记得同步 `server.py` 开头的 `CURRENT_VERSION`**，脚本会检查并提醒。
两处不同步的后果很隐蔽：安装包发出去了，但更新服务端认为「已经是最新」，
客户端永远收不到推送。

> `server.py` 里还有一处历史遗留的字符串版本比较（`'2.0.9' > '2.0.10'` 为真），
> 那会让推送永远失效。已改用 `app/core/version.py` 的语义化比较。

### 中文安装界面

Inno Setup 自带 28 个语言包，**里面没有简体中文**。文件缺失时 ISCC 是
**直接编译失败**（不是警告），所以脚本必须能探测、能退化。

查找顺序：

1. `packaging\ChineseSimplified.isl`（放这里就会进 git，换机器也能编出中文）
2. ISCC 安装目录的 `Languages\`
3. 都没有 → 安装界面退化成英文，构建照常完成，并打印下载地址

详见 `packaging\README.md`。

---

## 2. 产物的形状

```
Dist\ReagentManagementSystem\
├── ReagentManagementSystem.exe     ← 主程序（约 2.4 MB）
├── config.ini                      ← 含构建脚本写入的完整性哈希
├── LICENSE / SOURCE_CODE           ← GPL v3 合规
├── ui\                             ← Web 界面（21 个文件，明文散放）
├── images\ / reagent.ico
├── Qt5*.dll                        ← 40 个，**摊在根目录**
└── PyQt5\Qt5\
    ├── bin\QtWebEngineProcess.exe  ← Chromium 辅助进程
    └── translations\qtwebengine_locales\{en-US,zh-CN}.pak
```

注意 Qt5 的 DLL 在根目录，而辅助进程在 `PyQt5\Qt5\bin\` —— 这个不一致的
布局是 PyInstaller 的 PyQt5 hook 造成的，而且跨版本会变。所以
`build_installer.py` 的产物自检对这类文件是**按文件名搜**，不写死路径。

### 为什么是 onedir 而不是单文件

QtWebEngine 的辅助进程 `QtWebEngineProcess.exe` 必须相对主程序有可预测的
路径。onefile 每次启动解压到 `_MEIxxxx` 临时目录，辅助进程会找不到，
报 `Could not find QtWebEngineProcess.exe`。另外 Chromium 运行时约 150 MB，
onefile 每次启动都要解压一遍。

**但如果 onedir 在目标机上出问题，onefile 是一个有实证支持的备选** ——
旧版（37 MB，PyQt5 无 WebEngine）就是 onefile，在 Win7 上一直能跑。
改法：把 `build.spec` 末尾的 `COLLECT(...)` 去掉，`EXE(...)` 里删掉
`exclude_binaries=True`，并把第 3 个参数从 `[]` 改成 `a.binaries, a.zipfiles, a.datas`。

### ⚠️ onedir 会触发 PyQt5 里的一段代码

PyQt5 的 `find_qt()` 是这样写的：

```python
dll_dir = os.path.dirname(sys.executable)
if not os.path.isfile(dll_dir + '\\Qt5Core.dll'):
    ... 在 PATH 里找 ...
    else:
        return                      # 找不到就提前返回
try:
    os.add_dll_directory(dll_dir)   # 只有没提前返回才会执行
except AttributeError:
    pass
```

`os.add_dll_directory` 底层是 Windows API `AddDllDirectory`，**它需要 KB2533623**。
而 `except` 只捕获 `AttributeError` —— 缺补丁时抛的是 `OSError: WinError 127`，
捕获不到，异常直接冒出去，`import PyQt5` 整体失败。

| 布局 | `Qt5Core.dll` 位置 | `find_qt()` 走向 | 结果 |
|---|---|---|---|
| onefile | 解压到临时目录 `_MEIxxxx` | 三处都找不到 → **提前 return** | 从不调用该 API，能跑 |
| onedir | **就在 exe 旁边** | 第一个判断命中 → 调用 | 缺补丁时崩溃 |

`app/compat.py` 处理了这件事：启动时探测该 API 是否可用，不可用就先把
`os.add_dll_directory` 替换掉，让那次调用无害通过（Qt 的 DLL 本来就在 exe 旁边，
这个调用是多余的）。同时用绝对路径预加载运行库和 Qt 的 DLL 作为双保险。

**根治办法还是给系统装 KB2533623。** 下载时认准 `Windows6.1`（6.0 是 Vista 的）：

```
https://www.catalog.update.microsoft.com/Search.aspx?q=KB2533623
```

---

## 3. 完整性哈希

应用启动时比对主程序的 SHA-256 与 `config.ini` 里记录的值。

`config.ini` 里哈希**留空时校验自动跳过**。这是刻意的：每次重新打包 exe 的
哈希都会变，而补丁/升级只替换 exe、不替换 `config.ini`，旧哈希必然对不上。
如果这时判定为失败，用户看到的是「应用完整性异常」——一个指向完全错误方向的
报错，而且程序直接打不开。这个坑前后踩过两次。

所以：

* **正式发布必须重算**，否则校验形同虚设；构建脚本已经串进流程，还会复核。
* 仓库根目录的 `config.ini` 哈希**永远是空的**（它是构建模板），真正被打进
  安装包的是产物里那份。单独手动跑的话：

```
python tools/gen_hashes.py --dist Dist\ReagentManagementSystem
```

> 注意写入目标是**产物里**的 `config.ini`。早先默认写仓库那份，
> 结果哈希全写在了不会被分发的文件上。

---

## 4. 安装包脚本（setup.iss）

结构上几个容易踩的点，改之前先读：

**不写 `AppId`。** 不写时 Inno 用 `AppName` 作为 AppId，而 `AppName` 与旧版
安装包保持一致，于是新版能识别出「已安装旧版」并原地升级，而不是并排装两份。
同理 `DefaultDirName` 也和旧版一致。**加上 `AppId` 反而会把它变成另一个应用。**

**`[Files]` 只写一条整目录规则。** 不要再单独 `Source` 一次仓库根目录的
`config.ini` —— 那会把产物里已经写好哈希的那份覆盖掉，校验就永久失效了。

**`CloseApplications=force`** —— QtWebEngine 的辅助进程会占住
`Qt5WebEngineCore.dll`，不强制关闭会导致升级时留下新旧混版的目录。

**`SetupMutex`** —— 程序内的更新流程会调起安装包，加上防重入，
免得和用户手动双击的那个撞在一起。

**用户数据完全不碰。** 数据库在 `%APPDATA%\ReagentManagementSystem\`，
安装、覆盖安装都不动它。卸载时会额外复制一份带时间戳的备份到同目录的
`Backups\`，**但不删除原库** —— 卸载后重装仍能看到原有库存。

> 旧版用 `uninstall_backup.exe`，它在备份后会**删掉原库**。那个行为已废弃，
> 现在备份逻辑直接写在 `setup.iss` 的 `[Code]` 里，不到 20 行，可见可控，
> 也省掉一个 5 MB 的预编译二进制。
>
> ⚠️ 用 `{userappdata}` 定位数据目录，意味着**由谁卸载就备份谁的数据**。
> 多用户机器上用管理员账户卸载，可能备份的不是实际使用者的库。单机单用户
> 场景无影响。

### 升级与自动更新

程序内的「检查更新」下载的是**安装包**，然后以 `/SILENT /RESTARTAPPLICATIONS`
静默运行 —— 这个流程没有变，onedir 不影响它。

发布新版本时：

1. 把 `Output\药剂仓库X.Y.Z安装器.exe` 拷到 `server.py` 同级目录的 `updates\`
2. 把 `server.py` 的 `CURRENT_VERSION` 改成新版本号
3. 需要的话在 `updates\changelog.txt` 写更新说明（没有该文件时用通用文案）

---

## 5. 给测试机发补丁包

测试机已经装好、只是要换一版程序时，不用传 247 MB 的整个目录。

**补丁必须自带 `config.ini`**，且里面的哈希要**清空** —— 补丁替换了 exe，
旧哈希必然对不上，走「留空即跳过」的分支。不给 `config.ini` 的话，
补丁装上去会直接弹「应用完整性异常」。

补丁内容：

```
ReagentManagementSystem.exe     ← 主程序
config.ini                      ← 哈希清空，必须带
reagent.ico
_ctypes.pyd / libffi-7.dll      ← app/compat.py 的预加载列表要读它们
VCRUNTIME140.dll / _1.dll
ui\                             ← 整个目录，别只拷其中几个文件
```

不需要带 `Qt5*.dll`、`PyQt5\`、`QtWebEngineProcess.exe` —— 测试机目录里本来
就有，且不随代码改动。**但如果测试机目录不是本项目的产物**（比如拿旧的 2.0
目录改的），补丁会缺东西，直接传整个 `Dist\ReagentManagementSystem\` 更省事——
那个是自洽的，不依赖原目录里有什么。

**测试隔离**：补丁的 `config.ini` 里加一行
`[Application] DataDirName = ReagentManagementSystemTest`，数据库就落到
`%APPDATA%\ReagentManagementSystemTest\`，和正式库完全隔离。确认稳定后删掉
这行即可回到正式目录。**正式包绝不能带这行**，构建脚本会检查并拦下来。

---

## 6. 自测

```
python -m unittest discover -s tests          # 单元测试，76 项
python tools/ui_smoke_test.py                 # 界面冒烟测试，截图到 tools/
python tools/ui_smoke_test.py --theme dark    # 深色主题
```

冒烟测试会启动真实主窗口，遍历五个页面与两个抽屉，检查渲染是否正常，
最后把截图写到 `tools/ui_smoke*.png`。打包之后也可以直接用 exe 验证：

```
set RMS_SMOKE_SHOT=%CD%\shot.png
set RMS_SMOKE_DELAY_MS=15000
ReagentManagementSystem.exe
```

十几秒后会在指定路径生成一张界面截图，据此判断「真的渲染出来了」而不是
仅仅「进程没崩」。`python tools\build_installer.py --smoke` 就是自动做这件事。

---

## 7. 纯净 Win7 上能不能跑

### 随包自带、不需要系统预装

| 依赖 | 文件 |
|---|---|
| Python 3.8 运行时 | `python38.dll` / `python3.dll` |
| UCRT（通用 C 运行库） | `ucrtbase.dll` + 14 个 `api-ms-win-crt-*.dll` |
| Qt 5.15 + Chromium 83 | 40 个 `Qt5*.dll` + `PyQt5\Qt5\bin\QtWebEngineProcess.exe` + 语言包 |
| 图形 / OpenGL | `opengl32sw.dll`（软件渲染兜底）、`d3dcompiler_47.dll`、`libEGL.dll`、`libGLESv2.dll` |
| SQLite / OpenSSL | `sqlite3.dll`、`libssl-1_1-x64.dll`、`libcrypto-1_1-x64.dll` |
| 业务依赖 | `openpyxl`、`requests`（已封进 `base_library.zip`） |

**关键在 app-local 部署**：这些 DLL 就放在 exe 旁边，Windows 加载器会优先从
程序目录取，不看系统装了什么。`app/compat.py` 还会用绝对路径主动预加载一遍
作为双保险。所以「系统里没装什么」基本不影响。

### 需要系统提供的部分

| 项 | 必需？ | 缺了会怎样 |
|---|---|---|
| Windows 7 **SP1** 64 位 | **必需** | 安装包 `MinVersion=6.1sp1` 直接拒绝安装 |
| **KB2533623**（或替代更新 **KB3063858**） | **必需** | 全部 `.pyd` 加载失败，程序起不来。见下节 |
| VC++ 运行库 | **必需** | 程序**不自带**（见下方 MSVC 那节），靠安装包的 `VC_redist.x64.exe` 装进系统 |
| KB2999226 (UCRT) | 不必需 | 程序自带 app-local 副本 |

### ⚠️ KB2533623 是硬性前提，不是"最好装一下"

这一条曾经被本文档写错（"compat.py 能绕过"），实测证伪，记在这里免得再犯。

**症状**：干净 Win7 SP1 上装完启动，报

```
ImportError: DLL load failed while importing _ctypes: 参数错误。
...
ImportError: DLL load failed while importing QtCore: 参数错误。
```

注意是「**参数错误**」而不是「找不到模块」——`参数错误` = `ERROR_INVALID_PARAMETER(87)`。

**原因**：CPython 3.8 出于安全改了扩展模块的加载方式，`dynload_win.c` 里用
`LoadLibraryExW` 带 `LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS`
去加载 `.pyd`。这组标志正是 **KB2533623** 引入的；缺了它，`LoadLibraryExW`
因无法识别标志位直接失败。Python 官方为此专门更新过文档（bpo-42339）。

**为什么 compat.py 救不了**：出问题的代码在 `python38.dll` 内部，
发生在任何 Python 层代码之前；而 `app/compat.py` 要兜底就得先 `import ctypes`，
`ctypes` 恰恰是加载不了的那个模块。**这不是能靠写兼容层绕过的问题。**

**怎么判断是不是这个原因**：看 `startup-error.txt` 里的两处特征 ——

* `预加载运行库 0/18、Qt 0/15`（全零，说明 compat 连 ctypes 都没拿到）
* 逐模块导入自检里 `json`/`csv`/`zlib` 成功但 `ctypes`/`_socket`/`_sqlite3` 全失败
  （`zlib` 是 PyInstaller 引导器内置的，不是真扩展模块，所以它能"成功"）

**处理**：安装包会**静默探测并自动安装**，全程不打扰用户。这个更新发布于
2011 年，任何正常打过 Windows Update 的 Win7 都有；只有装完 SP1 就没再更新过
的机器才缺。

安装脚本的交互是这样设计的（`setup.iss` 的 `ShowFinalNotice`）：

| 情况 | 安装时的行为 |
|---|---|
| 本机已有该更新 | **不问、不提示**，安装完就是安装完 |
| 缺更新、包里带了 | 静默装，结束时提示一次「请重启电脑后再运行本程序」 |
| 缺更新、包里没带 | 结束时报错，让用户走 Windows Update |
| 运行库也缺 | 同上（这个更严重，优先报） |

**判定依据是安装前的基线，不是装完再探测一次。** 这一点容易写错：
更新装了但要重启才生效，所以**装完探测到的结果和"本来就没装"完全一样**
（都是"探测不到"）。只看装完的结果，就会把「本机本来就有、跳过安装」
也误报成需要重启。

所以 `setup.iss` 里用 `KbWasMissing` 在 **`CurStepChanged(ssInstall)`**
（即 `[Run]` 执行之前）记一次基线，结束时拿它和当前状态比：

```pascal
if not KbWasMissing then Exit;          { 本来就有 → 什么都没装 → 不提示 }
if HasLibrarySearchFlags() then Exit;   { 装了且已生效 → 不需要重启 }
                                        { 剩下才是：装了，等重启 }
```

**刻意不做**安装前的「本机缺少 KB2533623，是否继续安装？」那种盘问 ——
用户既不知道那是什么，也没有别的选择，而且大部分机器根本不缺，
问了纯属噪音。

> KB2533623 已被微软从下载中心下架，现在要装的是它的替代更新
> **KB3063858**。`Dependencies\Windows6.1-KB3063858-x64.msu` 已经放在仓库里了。
> 装完**必须重启**才生效 —— Windows 换不掉正在使用中的 `kernel32.dll`。

> ⚠️ **VC_redist 的版本别乱升。** 包里这份是 **14.44.35208**，属于最后一组
> 还能在 Windows 7 上安装的版本。从 **14.50** 起（随 Visual Studio 2026 发布）
> 已放弃 Win7，装不上。升级依赖时认准 `aka.ms/vs/17` 这个源。

### ⚠️ 别让 PyInstaller 自带 MSVC 运行库

同一个"缺东西"家族的另一半，但表现完全不同 —— 这次是**加载期的弹窗**，
不是 Python 异常：

```
ReagentManagementSystem.exe - 无法找到入口
无法定位程序输入点 GetSystemTimePreciseAsFileTime 于 KERNEL32.dll 上。
```

`GetSystemTimePreciseAsFileTime` 是 **Windows 8 才有**的 API，Win7 上任何补丁
都不提供。这个弹窗说明某个 DLL **静态导入**了它，加载时解析不到符号。

**根因**：PyInstaller 会从**打包机的 System32** 抓 MSVCP140 / VCRUNTIME140
这些运行库，打包机上装什么版本就带什么版本；而开发机的 Visual Studio
永远比目标机新。

实测踩过一次：开发机 System32 的 `MSVCP140.dll` 是 **14.51**（Visual Studio
2026），**这一版已经放弃 Windows 7**。用 `tools/pe_imports.py` 一查就现形：

```
MSVCP140.dll (14.51)     -> 静态导入 GetSystemTimePreciseAsFileTime × 1   ← 坏的
MSVCP140.dll (14.44.35211) -> × 0                                        ← 好的
```

**最阴的地方**：这个 app-local 副本会**盖住** System32 里那份好的（应用目录在
DLL 搜索顺序里最靠前），所以目标机就算装了正确的 VC++ 运行库也救不回来。
症状因此非常反直觉 —— **装了完整安装包反而不如只发一个 exe 的补丁包能跑**
（补丁包里没带这个 DLL，用的是系统那份）。

**处理**（已落实）：

1. `build.spec` 里排除 `msvcp140*` / `vcruntime140*` / `concrt140*` / `vccorlib140*`，
   不随包分发。顺带解决了版本混搭问题（以前产物里是 MSVCP140 14.51 配
   MSVCP140_1 14.26，本身就不该出现）。
2. 改由系统提供，由安装包里的 `VC_redist.x64.exe` 负责 —— 那份**刻意钉死在
   **14.44**，最后一个还支持 Win7 的版本。别盲目升成最新版。
3. `tools/build_installer.py` 的产物自检里加了一道 **API 面检查**：扫全产物有没有
   静态导入 Win8+ 的 API，有就直接中止构建。这类问题在开发机上**永远复现不了**
   （开发机是 Win10，那些 API 都在），只能在构建期拦。

排查这类问题时用：

```
python tools\pe_imports.py --check-dir Dist\ReagentManagementSystem ^
    --find-symbol GetSystemTimePreciseAsFileTime
```

### 结论

理论上纯净的 Win7 SP1 x64 装完就能跑 —— 重依赖里除了 MSVC 运行库，其余都是
app-local 的。
**但这件事没有在真机上验证过。** 唯一的确认方式是在干净虚拟机里装一遍。

### 验证步骤

1. 新建（或从快照还原）一台 **Windows 7 SP1 x64**，不要预装任何运行库或补丁
2. 拍个快照，后面还要反复用
3. 装 `Output\药剂仓库2.1.0安装器.exe`
4. 确认：桌面快捷方式出现 → 程序能启动 → 能入库一条试剂 → 能导出 Excel
5. 如果起不来：
   - 先看**桌面**上的 `startup-error.txt`（程序目录写不进去时会落到桌面）
   - 再把 `Output\win7diag.exe` 拷进去，放到程序目录的**旁边**双击运行
   - 它会写出 `win7diag-result.txt`，把这两个文件发回来

`win7diag.exe` 是独立的控制台程序，只依赖 Python 运行时，**不碰 Qt**。
它和主程序共用同一套运行时，所以能把故障切成两半：

* 它**也**起不来 → 问题在 Python 运行时层面（Python 本身、UCRT、加载器）
* 它**能**跑而主程序不行 → 问题在 Qt / Chromium，诊断输出会指出断在哪一步

`win7diag.exe` 不进安装包，只在排查时用。重新构建：

```
python -m PyInstaller --noconfirm --clean --onefile --console ^
    --name win7diag --icon "%CD%\reagent.ico" ^
    --distpath Output --workpath build/pyi-diag --specpath build/pyi-diag ^
    tools\win7_diag.py
```

> `--icon` 必须给**绝对路径**：`--specpath` 会把 spec 写到 `build\pyi-diag\`，
> 相对路径的 `reagent.ico` 会按 spec 所在目录去找，然后报
> "Unable to open icon file"，构建静默失败只留下一个 `.notanexecutable`。

---

## 8. Win7 上可能遇到的问题

| 现象 | 原因 | 处理 |
|---|---|---|
| 启动即退出，无任何提示 | 缺运行库 | 装 `Dependencies\VC_redist.x64.exe` |
| `DLL load failed while importing QtCore` 等 | 加载器找不到同目录的依赖 | 见下方「依赖都在，但加载失败」 |
| 界面全白 | WebEngine 渲染失败 | 见下方渲染模式 |
| 提示「应用完整性异常」 | exe 换了但 `config.ini` 里的哈希没跟着变 | 重跑构建脚本，或清空哈希 |

### 依赖都在，但加载失败

`DLL load failed while importing XXX` 这个报错**不告诉你是哪个依赖出的问题**，
而"某个机器能跑、另一台不行"最容易让人往错的方向查（重装运行库、换机器）。

先用 `tools/pe_imports.py` 把事实确定下来 —— 是**真的缺文件**，还是文件都在、
只是加载器找不到：

```
python tools/pe_imports.py --check-dir Dist\ReagentManagementSystem ^
    Dist\ReagentManagementSystem\PyQt5\QtCore.pyd
```

它会把每个依赖分成「产物里有 / 靠系统」两类，`[未知]` 那几行才是真问题。

整包扫一遍也行（对每个 PE 文件都跑一次）：

```
python tools\pe_imports.py --check-dir Dist\ReagentManagementSystem <任一文件>
```

**如果扫下来一个不缺**，那问题就是加载器**搜索路径**，不是缺文件。这种情况
`app/compat.py` 的预加载（用绝对路径把 Qt 与运行库先装进进程）就是解药 ——
加载器按**基名**匹配已加载模块，模块一旦在进程里，谁再要它都直接复用。

> 这里踩过的坑：`compat.apply_all()` 原本写成「探测到 `add_dll_directory`
> 可用就直接 return，跳过预加载」。但探测成功只说明那个 API 存在，
> 并不保证之后 `QtCore.pyd` 找得到 `Qt5Core.dll`。现在预加载无条件执行。

确诊要看 `startup-error.txt` 里 `compat.describe()` 的那几行：

```
--- DLL 预加载与 add_dll_directory ---
  [原生可用] 本机支持 AddDllDirectory（已装 KB2533623 或系统为 Win8+）
  预加载运行库 18/18、Qt 15/15
```

打包后的正常值就是 **18/18 和 15/15**。数字不够，后面会跟着
`！Qt 预加载失败 N 个` 以及具体的 DLL 和系统错误。

### 渲染模式

老显卡上 QtWebEngine 可能渲染出纯白界面且不抛异常。程序会在
`%APPDATA%\ReagentManagementSystem\userData.ini` 里读 `[UI] RenderMode`：
`hardware`（默认）或 `software`。改后者会禁用 GPU 加速，走软件渲染。

### 故障报告

打包后的程序没有控制台。启动失败时会在**桌面**和程序目录下写出
`startup-error.txt`（网络组件故障则写 `network-error.txt`），
里面包含完整堆栈和环境自检结果。
