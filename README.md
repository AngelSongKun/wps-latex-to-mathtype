# wps-latex-to-mathtype

> 一键把 WPS / Word 文档里的数学公式换成 MathType 公式对象 —— 全自动、可续转、带验证。
> Batch-convert LaTeX / OMML equations in `.docx` files into native MathType (OLE) equations on Windows.

[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue)](#环境要求)
[![Python](https://img.shields.io/badge/python-3.9%2B-3572A5)](scripts/requirements.txt)
[![MathType](https://img.shields.io/badge/requires-MathType%207%2B-8A2BE2)](#环境要求)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

---

## 这个 Skill 解决什么问题

写论文、出报告的人常遇到这种死局：正文里公式是 LaTeX 文本或者 Word 原生公式，但期刊 / 学校要求必须是 **MathType 公式对象**（双击能进 MathType 编辑器）。手工一个一个点要几小时，WPS 自带的"转换公式"又枚举不到 WPS 原生公式。

本 Skill 把整条链路自动化：

| 输入形态 | 处理路径 |
|---|---|
| LaTeX 文本公式 `$...$`、`$$...$$`、`\[...\]` | 先转成原生 OMML 公式，再走 MathType |
| 已是原生公式（OMML） | 直接批量转 MathType OLE 对象 |

转换覆盖**正文、表格、文本框、脚注、尾注、页眉、页脚**（不只正文），支持批量文件、断点续转、Word 崩溃自动恢复，并在转换后做计数校验、字形扫描和可视页渲染。

## 环境要求（重要）

本工具**只能在 Windows 上工作**，且依赖以下商业软件：

- Windows 10 / 11
- Microsoft Word（2016 及以上）
- MathType 7+（需含 Word 集成模板，通常在 `C:\Program Files\Microsoft Office\root\Office16\STARTUP\`）
- Python 3.9+，建议使用 `.venv`

如果是 Linux / macOS，本 Skill 不适用（MathType OLE 是 Windows COM 生态）。

```bash
pip install -r scripts/requirements.txt
```

可选依赖：`pymupdf` + 本机 Word（用于 `verify_output.py --word-pdf` 显示层终检）。

## 安装

把仓库放到你的 Agent Skills 目录即可，例如：

```bash
git clone https://github.com/<your-name>/wps-latex-to-mathtype.git
cp -r wps-latex-to-mathtype ~/.codeartsdoer/skills/     # 或 ~/.claude/skills/
```

目录结构需保持 `SKILL.md` 在根目录，`scripts/`、`references/` 与其同级。

## 使用

### 第 0 步 · 环境修复（只需一次）

```bash
python scripts/fix_environment.py
```

修复 Windows 11 24H2 起删除 MSXML 版本化 ProgID / TypeLib 注册的问题——不修的话 MathType 转换引擎会**全链路静默失败**（0 转换、无报错）。幂等，仅写用户级注册表，无需管理员。

### 第 1 步 · 预检

```bash
python scripts/precheck.py 输入.docx --json 报告.json
```

输出公式分布（各故事区域的 OMML 计数）、LaTeX 候选数、风险结构（矩阵 / 方程组转换后会被扁平化）、字符集嫌疑清单，并直接给出路由结论。

### 第 2 步 · LaTeX 文本 → 原生公式（仅当有 LaTeX 候选）

```bash
python scripts/latex_to_omml.py 输入.docx
```

产出 `输入_converted.docx`，失败公式保留原文并写入报告。

### 第 3 步 · 原生公式 → MathType 对象（核心）

```bash
python scripts/convert_via_word_vba.py 输入.docx [更多.docx ...] -o 输出.docx
```

- 宏以 Word Startup 插件（`PyMathTypeConv_vN.dotm`）形式安装，日常转换**零 VBE 访问**
- 逐公式：`WordOpenXML` 取 OMML → OMML2MML.XSL 转 MathML → 非 ASCII 转数字实体 → 插入空白 MathType 公式 → 引擎直调 `MTSetEqnFromLangStr` 写入
- 每 20 个公式自动存盘；Word 崩溃后自动重开并从检查点续转（每文件最多 3 次）
- 输入文件本身永不修改，输出默认 `输入_MathType.docx`

常用参数：`--limit N` 冒烟测试、`--kill-word` 强杀残留 Word、`--rebuild-addin` 重建宏插件。

### 第 4 步 · 验证

```bash
python scripts/verify_output.py 原始.docx 转换后.docx --render --word-pdf
```

自动检查：OMML 残留是否为 0、`Equation.DSMT4` 数与嵌入对象数是否等于原始公式总数、EMF 预览中的 `?` 缺字形、零文本预览；`--word-pdf` 用 Word 本体导 PDF 逐字形扫描（**这步不可用 LibreOffice 替代**，字体替换路径不同会掩盖问题）。

## 实战记录

- 2026-10：228 / 228 公式全量转换成功
- 2026-11-05：插件化架构重构后 257 / 257 验证通过

## 已知限制

- 矩阵（`m:m`）、方程组（`m:eqArr`）转 MathType 后会被扁平化，需按报告清单手工重建
- MathType 翻译器画不出部分预组合重音符，脚本已自动将其重写为 `m:acc` 结构；表外组合符（如 cedilla）安全跳过
- 转换过程中若已有 `WINWORD.EXE` 在运行，脚本会拒绝启动（防进程劫持）；请先关闭或使用 `--kill-word`
- 仅支持 `.docx`，`.doc` / `.wps` 请先另存为 `.docx`

## 目录结构

```
wps-latex-to-mathtype/
├── SKILL.md                              # Agent Skill 主文件
├── scripts/
│   ├── fix_environment.py                # 一次性环境修复
│   ├── precheck.py                       # 转换前预检
│   ├── latex_to_omml.py                  # LaTeX 文本 → OMML
│   ├── convert_via_word_vba.py           # OMML → MathType OLE（核心）
│   ├── verify_output.py                  # 转换后验证
│   ├── docx_scan.py                      # 共享检查库
│   ├── symbol_audit.py                   # 符号对照审计
│   └── requirements.txt
└── references/
    ├── latex-mathtype-map.md             # 符号映射审计表
    ├── mathtype-conversion.md            # 手动方案 / Toggle TeX
    └── troubleshooting.md                # 环境坑与失败路径全记录
```

排障请先读 `references/troubleshooting.md`，里面记录了所有已实测失败的弯路与根因。

## 免责声明

本工具会调用 Word COM 与 MathType 引擎修改文档副本。**使用前请自行备份原始文件**。作者不对因进程强杀、软件兼容性或文档损坏造成的任何损失负责。

## License

[MIT](LICENSE)