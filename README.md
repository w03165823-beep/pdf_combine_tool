# PDF / RTF 合并工具

Windows 桌面工具：合并 PDF、将 RTF 转为 PDF，并按文件名或 Metadata 生成书签和可点击目录（TOC）。

文档更新：2026-10-09。Filename + TOC 和 Metadata + TOC 均支持自动转换 RTF；前者无需 Metadata，后者使用 Metadata 控制顺序与标题。

## 运行条件与启动

- Windows；当前在 Python 3.11 环境中验证。
- 处理 RTF 需要 Microsoft Word 桌面版；合并现有 PDF 不需要 Word。
- 保持项目的源代码、图片和模板目录结构完整。

首次使用，在 PowerShell 中执行：

```powershell
cd D:\AI\pdf_combine_tool
python -m pip install -r requirements.txt
```

之后双击 **start.bat**。它优先使用 Python 3.11，找不到时尝试 PATH 中的 Python；启动失败会保留错误。运行期间保持启动窗口打开。

手动启动：

```powershell
cd D:\AI\pdf_combine_tool
python -m src.main
```

从项目根目录启动，不要直接双击 src/main.py。如果提示缺少 fpdf：

```powershell
python -m pip install fpdf2==2.8.1
```

## 基本操作

1. **Choose Folder**：选择源文件夹。
2. **Merge mode**：选择模式。
3. **Output As**：填写输出名称，例如 combined，不填 .pdf 后缀。
4. Metadata 模式需要选择元数据文件。
5. 点击 **GO!**，通过日志查看顺序和结果。

结果保存在所选文件夹，例如 combined.pdf。同名结果会被替换。只扫描该文件夹的直接文件，不递归扫描子文件夹。

普通模式和文件名模式排除当前同名输出文件，但其他旧合并结果仍可能成为输入，建议源文件夹只放本次需要处理的文件。

## 三种模式

| 界面选项 | 输入 | 顺序 | 书签标题 | 前置目录 |
| --- | --- | --- | --- | --- |
| Combine PDFs without TOC | PDF / RTF | 文件名字母顺序 | 不生成 | 不生成 |
| Filename + TOC | PDF / RTF | 文件名数字自然排序 | 完整 PDF 文件名 | 可点击 |
| Metadata + TOC | PDF / RTF | Order 升序 | 直接使用 Title | 可点击 |

### 1. 普通合并

无需 Metadata，不生成目录和书签。RTF 先转换到 _PDF 子文件夹，再与 PDF 合并。同名 RTF 和 PDF 同时存在时采用 RTF 转换结果。

按字母顺序排序，可能出现 1.pdf、10.pdf、2.pdf。此模式沿用原有转换实现，可能复用 _PDF 中已有结果；更新 RTF 后，应检查对应缓存 PDF 是否需要移走后重新转换。

### 2. Filename + TOC

自动收集 PDF 和 RTF。即使文件夹中没有 PDF，也会先通过 Word 将 RTF 转换到 _PDF，再按数字自然排序，例如 1.pdf → 2.pdf → 10.pdf。同名 RTF 和 PDF 同时存在时优先采用 RTF 转换结果，只合并一次。

操作步骤：

1. 选择包含 RTF、PDF 或两者混合的源文件夹。
2. 选择 **Filename + TOC**，不需要准备或选择 Metadata。
3. 填写输出名称，点击 **GO!**。
4. 有 RTF 时自动转换；全部输入按文件名数字顺序合并，并在结果前面生成可点击目录。

例如源文件夹中只有 `1.rtf、2.rtf、10.rtf`，会先在 `_PDF` 中生成 `1.pdf、2.pdf、10.pdf`，再依次合并。目录和书签标题分别为这三个 PDF 文件名。

每个 PDF 的完整名称（包含 .pdf）作为目录标题和一级书签；RTF 使用转换后的同名 PDF 名称。目录放在正文前面，点击条目或书签跳到该文件在最终 PDF 中的第一页。

### 3. Metadata + TOC

选择模式后，在 **TLF Metadata** 旁点击 Browse，选择 CSV 或 XLSX。

只处理 Metadata 列出的文件，按 Order 排序，Title 直接作为目录和书签标题，不拼接其他列。TFL 只校验类型，不分组或排序。

两个 TOC 模式均使用独立 Word 实例将 RTF 转换到 _PDF，每次重新转换。转换失败停止合并。

这两种模式的主要区别是：**Filename + TOC 自动收集文件，用文件名决定顺序和标题；Metadata + TOC 只处理表中列出的文件，用 Order 决定顺序、Title 决定标题。** 两者都会在正文前添加可点击目录并生成书签。

## Metadata 格式

当前使用四列，表头大小写不限；旧的八列格式不再用于第三种模式。

| 列 | 含义 | 要求 |
| --- | --- | --- |
| Order | 合并顺序 | 数值，必须唯一，不必连续 |
| TFL | 文件类型 | T=表、F=图、L=列表；大小写不限 |
| FileName | 文件名 | 不含路径；可带 .rtf / .pdf，也可不带后缀 |
| Title | 完整标题 | 非空，直接用于书签和目录 |

```csv
Order,TFL,FileName,Title
1,T,t-14-01-03-01-01-demo,表1：受试者基本信息
2,F,f-15-01-03-01-tte-forest,图1：森林图
3,L,listing01.pdf,列表1：详细数据
```

文件名规则：

- 不带后缀时，优先匹配同名 RTF，再匹配 PDF；带后缀时匹配指定类型。
- 两边统一转为大写比较，转换时保留实际文件名。
- 去除 Metadata 文件名前后空格；数字、内部空格和分隔符仍需一致。
- 不把连字符改成下划线，也不改写点号。
- 例如 F-15-01-03-01-TTE-FOREST 可匹配 f-15-01-03-01-tte-forest.rtf。

**Create Metadata Template** 在所选文件夹创建并打开 metadata_template_v2.csv；已存在时直接打开，不覆盖。项目示例：[metadata_simple.csv](examples/metadata_simple.csv)。

Excel 编辑后可保存为 XLSX，或选择 **CSV UTF-8**。标题含英文逗号时需按 CSV 规则加引号，Excel 保存可自动处理。

以下问题会停止处理：缺失文件、重复或非数值 Order、非法 TFL、空标题、FileName 含路径、输入与输出同名。找不到文件时提示文件夹和相近文件名。

## 目录与密码

第二、第三种共用目录流程：确定顺序和标题 → 按文字宽度换行 → 计算目录页数 → 合并正文 → 添加目录链接和书签 → 保存。

支持中文和英文，嵌入字体，长标题自动换行，目录可跨页。目录页码是最终 PDF 的页序号，并非源文件印刷页码。无需旧版字体选择、标题分隔符或人群拼接选项。

**Set password** 当前设置所有者密码（owner password），没有设置打开密码，不代表打开文件时必须输入密码。

## Python 文件职责

| 文件 | 做什么 | 修改场景 |
| --- | --- | --- |
| src/main.py | 入口：创建界面、工具和调度对象，连接按钮动作，启动窗口事件循环，显示失败弹窗 | 启动或按钮连接 |
| src/gui.py | 窗口、标签、按钮、输入框、模式切换、模板按钮、密码控件及资源路径 | 界面文字、位置、显示隐藏 |
| src/pdf_compiler.py | 调度：根据模式收集、排序、转换文件，调用合并并记录日志 | 模式规则与处理顺序 |
| src/toc_merge.py | 新版 Metadata 解析、校验、文件匹配，以及目录分页、书签、链接和最终 PDF 保存 | Metadata 格式、目录样式、页码 |
| src/pdf_util.py | 文件选择、文件夹操作、普通 PDF 合并、Word 转换、界面日志；保留旧版辅助方法 | 转换和基础文件操作 |
| build.py | 检查资源，清理旧打包目录，调用 PyInstaller | 制作 EXE |
| run_tests.py | 调用 pytest，输出覆盖率，支持性能测试参数 | 原有测试套件 |
| tests/test_pdf_combine.py | 原有工具、旧 Metadata / TOC、转换及性能测试 | 历史兼容检查 |
| tests/test_filename_merge.py | 实际 PDF 测试：数字排序、多页书签、点击链接、重复合并、中文名称，以及纯 RTF / 混合输入的模拟转换 | 文件名模式回归 |
| tests/check_toc_modes.py | 无需 pytest 的回归脚本：两种 TOC、CSV/XLSX、RTF 匹配与模拟转换、长标题、跨页和报错 | 当前主要功能验证 |
| src/__init__.py、tests/__init__.py | 标识 Python 包，不是用户启动入口 | 通常无需修改 |

pdf_compiler.py 和 pdf_util.py 仍保留部分旧版八列 Metadata 和 TOC 辅助方法。当前第二、第三种使用 toc_merge.py 的新流程，维护时优先看新流程。

```text
start.bat → src/main.py
              ├─ gui.py：界面与选项
              └─ pdf_compiler.py：流程调度
                   ├─ pdf_util.py：普通合并、RTF 转换、日志
                   └─ toc_merge.py：Metadata、目录、书签、最终 PDF
```

其他文件：

- start.bat：双击启动器。
- requirements.txt：依赖清单。
- main.spec：PyInstaller 打包配置。
- assets/images/：图片与图标。
- examples/metadata_simple.csv：当前模板；metadata_example.csv 是历史模板。

## 测试与打包

快速回归，在项目根目录执行：

```powershell
python tests/check_toc_modes.py
```

使用临时文件，Word 转换采用模拟调用，不能替代实际 Word 与自己的 RTF 验证。脚本会在 tmp/pdfs/ 生成布局检查用 PDF。

原有 pytest 套件（需要测试依赖）：

```powershell
python run_tests.py
```

打包 EXE 可选，日常双击 start.bat 即可：

```powershell
python -m pip install pyinstaller
python build.py
```

打包会清理旧 build/、dist/，按 main.spec 生成 EXE。EXE 打包流程尚未完整验证，RTF 转换仍需要本机 Word。

## 常见问题

| 问题 | 排查 |
| --- | --- |
| 无法启动 | 查看启动窗口报错；检查 Python 与依赖是否在同一环境 |
| 找不到源文件 | 核对 Choose Folder、FileName、数字及分隔符；大小写不限 |
| RTF 转换失败 | 确认 Word 桌面版能打开该文件，查看具体错误 |
| 中文乱码 | 使用 XLSX 或 CSV UTF-8 |
| 界面还是旧版本 | 关闭旧窗口，再双击 start.bat |
| 普通模式 RTF 内容没更新 | 检查 _PDF 缓存；Filename + TOC 和 Metadata + TOC 每次重新转换 |
| 文件名模式的源文件是 RTF | 自动转换，无需 Metadata；需要 Microsoft Word 桌面版 |

大批量文件可能占用较多内存。正式处理前建议用自己的小批量文件确认转换版式、顺序和跳转。

## 许可证

GNU GPL v3.0，详见 [LICENSE](LICENSE)。Word 转换需合法的 Microsoft Office 授权。

