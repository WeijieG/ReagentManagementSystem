# packaging\

安装包相关的第三方素材，由 `tools\build_installer.py` 自动取用。

## ChineseSimplified.isl

Inno Setup 自带 28 个语言包，**里面没有简体中文** —— 中文属于非官方翻译，
要单独下载。缺了它，`setup.iss` 里 `MessagesFile` 指向不存在的文件，
ISCC 是**直接编译失败**，不是给个警告。

构建脚本按这个顺序找：

1. `packaging\ChineseSimplified.isl`（就是这里）
2. ISCC 安装目录的 `Languages\ChineseSimplified.isl`
3. 都没有 → 安装界面退化成英文，构建照常完成

**想让安装界面是中文，把文件放到这个目录下即可**，不用改任何脚本。

下载地址（文件约 30KB）：

```
https://github.com/jrsoftware/issrc/raw/main/Files/Languages/Unofficial/ChineseSimplified.isl
```

GitHub 在国内经常连不上，用你自己的方式拿。文件头应该有类似这样的内容：

```
; *** Inno Setup version 6.x+ Chinese Simplified messages ***
```

拿到后确认里面有 `ButtonNext`、`WelcomeLabel1` 这些键，才是完整文件。

放好之后重跑 `python tools\build_installer.py`，输出里会有一行
`安装界面语言包: packaging\ChineseSimplified.isl`。
