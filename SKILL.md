---
name: wps-latex-to-mathtype
description: 将 WPS/Word 文档中的数学公式替换为 MathType 公式对象。覆盖两种输入：① LaTeX 文本公式（$...$、$$...$$、\[...\]）先用脚本转为原生 OMML 公式；② 已是原生公式（OMML）的文档直接进入转换。核心是自动化的"原生公式 → MathType OLE 对象"批量转换（Word COM + 注入 VBA + MathType 引擎直调），全程无需用户手动操作对话框。凡是用户提到 WPS/Word/docx 里的公式要转 MathType、LaTeX 公式替换、公式批量转换、论文公式格式统一等场景，都应使用本 skill。
---

# WPS/Word 公式 → MathType 公式（全自动）

把 .docx 中的公式替换为 MathType 对象，产出可直接交付的新文件。流程分四步，除标注外全部由脚本自动完成：

## 第 0 步：前置检查

- 输入必须是 `.docx`（`.doc`/`.wps` 先在 WPS 中另存为 .docx）。
- **先物理备份原文件**。
- 依赖：`pip install pywin32 python-docx lxml`（comtypes 仅备用路径需要；
  `--render` 视觉验证另需 LibreOffice + pymupdf）。
- 本机需已安装：Microsoft Word（含 MathType 的 Word 集成模板，通常在
  `C:\Program Files\Microsoft Office\root\Office16\STARTUP\`）和 MathType 本体
  （`C:\Program Files (x86)\MathType\`）。
- **先运行 `scripts/fix_environment.py`（幂等，用户级注册表，无需管理员）**。
  它修复 2026-10 实测发现的四大环境坑，详见 references/troubleshooting.md：
  Win11 24H2+ 删除了版本化 MSXML ProgID/TypeLib 注册，导致 MathType 转换引擎
  全链路静默失败。

## 第 1 步：预检（一条命令出全部判断）

```bash
python scripts/precheck.py 输入.docx [--json 报告.json]
```

输出：① 公式分布（正文/表格、文本框、脚注、尾注、页眉页脚的 OMML 计数，
即转换预期值）；② LaTeX 文本候选数；③ 风险结构（矩阵 m:m、方程组 m:eqArr 等，
转换后会被扁平化、需手工重建的清单）；④ 字符集嫌疑（含 GBK 外字符的公式及其
上下文定位线索——这些页转换后必检）。结论行直接给出路由：

- LaTeX 候选 > 0 → 先跑第 2 步；
- 全部为原生 OMML → 直接第 3 步；
- 两者都无 → 无需转换。

## 第 2 步（仅 LaTeX 文本输入）：LaTeX → 原生公式

```bash
python scripts/latex_to_omml.py 输入.docx      # 产出 输入_converted.docx + 报告
```

覆盖正文、表格、**文本框、脚注、尾注**、页眉页脚（脚注/尾注经文档关系部件的
blob 打补丁写入）。详见脚本头部注释（`--all`/`--only-high`/`--report`）。
转换后让用户在 WPS/Word 抽查渲染。失败公式保留原文，用
`references/mathtype-conversion.md` 方案 B 手工处理。

## 第 3 步：原生公式 → MathType 对象（核心，全自动）

```bash
python scripts/convert_via_word_vba.py 输入.docx [更多.docx ...] [-o 输出.docx] \
    [--kill-word] [--limit N] [--rebuild-addin]
```

- **宏插件架构（2026-11-05）**：转换宏一次性安装为 Word Startup 文件夹的
  `PyMathTypeConv_vN.dotm`（与 MathType 自身集成同机制、默认受信任）。
  日常转换**零 VBE 访问**（VBE 自动化是旧管线最不可靠环节）；VBA 变更时
  bump `ADDIN_VER` 自动重建，`--rebuild-addin` 强制重建，卸载=删 dotm。
- 支持**批量**：多个输入文件顺序转换，`-o` 仅限单输入。
- 输出默认 `输入_MathType.docx`；输入文件本身永不落盘修改。
- **重音分解（自动）**：复制到输出路径后、交给 Word 前，`decompose_accents()`
  基于 **NFD 规范分解**把公式里一切可分解的重音字符（Ŷ/ŷ/**Ã/Â** 等 Latin-1 与
  Latin-Ext 预组合字符，及 x+U+0302 等 11 种组合序列）重写为 m:acc 结构——
  MathType 翻译器画不出这些字符（缺字形菱形，或混入 GBK 字体记录变 '?'），
  m:acc 则在 Word/WPS/PDF 全路径完美渲染。表外组合符（如 cedilla）安全跳过。
- **断点续转 + 崩溃自动恢复**：工作文档直接以输出路径存在，每 20 个公式自动
  存盘（VBA `SAVE_EVERY`）；输出已完成时重跑提示"已完成"退出；输出与输入
  不匹配时拒绝启动。**Word 进程崩溃（RPC 断连）自动重开并从检查点续转**
  （每文件最多 3 次，续转前做 census 一致性校验）。
- `--limit N` 只转前 N 个公式（冒烟测试；经 Run 参数传入宏）。
- **脚本拒绝在任何 WINWORD.EXE 已运行时启动**（残留进程会劫持
  GetActiveObject 附着，user 自己开的 Word 同理）——按提示关闭后重跑，
  或加 `--kill-word` 强杀（会丢其他 Word 文档未保存内容）。
- 残留的 `~$` 锁文件自动清理。

原理（2026-10 实战 228/228；2026-11-05 插件架构 + 257/257 验证）：

1. 以自定义 PATH（含 `MathType\MathPage\64`、`System\64`）启动 WINWORD `/embedding`；
2. 转换宏从 Startup 插件直接 `Application.Run`（首次运行自动安装插件）；
3. 宏**遍历 `doc.StoryRanges`（全故事：正文/表格/文本框/脚注/尾注/页眉页脚，
   而非只覆盖正文的 `doc.OMaths`；跨节故事经 `NextStoryRange` 链接，先快照再处理）**，
   逐公式完成转换：
   - `Range.WordOpenXML` 取 OMML；
   - 晚期绑定 `CreateObject("MSXML2.DOMDocument30")` + OMML2MML.XSL 转 MathML；
   - **`EntityEncode` 把 MathML 中所有非 ASCII 字符转为数字实体（`&#8722;`）**，
     此前先做三步清洗：**NBSP→普通空格**（OMML2MML.XSL 把 m:t 空格转成 NBSP
     `<mi>`，预览按字节 0xA0 绘制而现代 Symbol 字体缺 F0A0，Word/WPS 会画成
     '?' 并与相邻字形叠印）；**中点 ·(U+00B7)→点积符 ⋅(U+22C5)**（公式含
     中文时 MathType 把 · 分进宋体记录，0xB7 落 GBK 前导区变 '?'）；
     EntityEncode 本身防非 ASCII 字符经 MathType 解析器的环境相关损坏；
   - `rng.InsertFile BlankEqn.doc, "MTBlankEqn"` 插入空白 MathType 公式
     （**必须用 `rng.InsertFile` + `rng.Expand wdWord` + `rng.InlineShapes(1)`，
     用 `Selection.InlineShapes` 找不到刚插入的对象**）；
   - `DoVerb(2)` 激活 OLE 服务器 → `MTSetEqnFromLangStr(obj, 2, bytes, len)`
     写入 MathML → `MTCloseOleObject` 关闭；
   - **循环结束后 1.5s DoEvents**：最后一个处理的公式（文档第一个）的 OLE
     预览缓存与最终存盘存在竞态，偶发整条空白（2026-11-03 实测并修复）。
4. 返回 `ok/fail/left/shapes` 统计，**ok 应等于 precheck 的公式合计**
   （断点续转时等于剩余数）。VBA 运行期错误以 `VBAERR <号>: <描述>` 返回；
   引擎级失败以 `stats=|-9999@<序号>` 报告（序号倒序处理，文档位置 =
   总数−序号+1）。

**不要走的弯路**（都已实测失败，见 references/troubleshooting.md）：
- WPS 的 MathType"转换公式"枚举不到 WPS 原生公式；
- Word 里静默调 MathType 宏 `DoConvertEquations`（引擎/MSXML 环境坑修好后
  仍静默 0 转换，其内部 InsertFile 逻辑另有问题）；
- 从 Python 跨进程传 IDispatch 给 WLL 引擎函数（返回 -9999）；
- 跨工程直接调用或经 `Application.Run` 传长字符串调用 MathType 宏内部函数
  （宏工程加密 + Run 长字符串参数失败）。

## 第 4 步：验证与交付

```bash
python scripts/verify_output.py 原始.docx 转换后.docx --render --word-pdf
```

**"原始"= 公式形态与转换时一致的文件**：纯 OMML 输入就是原文件；走过第 2 步
（LaTeX→OMML）时要用第 2 步的输出（`*_converted.docx`），否则公式总数对不上。

自动检查（全故事，mc:Fallback 死副本剥离）：各部分 OMML 残留 = 0；
`Equation.DSMT4` 数 == `word/embeddings/*.bin` 数 == 原始公式总数；
python-docx 可打开；**EMF 预览 '?' 字形扫描**（MathType 缺字形的确定性
检测）；**零文本预览告警**（空白/位图预览清单，供人工目检）。

**`--word-pdf`（显示层的最终验收，强烈建议）**：用 Word 本体导出 PDF 并逐字形
扫描——`?` 替换字形与同坐标叠字（如 NBSP 工件画出的 '1?'）应均为 0。
**LibreOffice 渲染不能替代此检查**（其字体替换路径不同，会隐藏这类工件；
实测视觉抽查也会漏小叠字）。`--pdf 已有.pdf` 可跳过导出直接扫描。

`--render` 额外把 **字符集嫌疑公式所在页**（由 precheck 的线索驱动，**不要按
固定间隔抽页**——实测固定抽样漏掉过 21 个问题公式）渲染为 PNG。

已知 MathType 缺口已清零（`\gg`/`\vdash` 经命名实体路径修复，2026-11-05，
详见 `references/latex-mathtype-map.md`）；precheck 仍会列出这两类字符作为
回归验证信号。

视觉验收：对渲染出的必检页逐页检查公式（空白框/乱码/错位）。**最终验收用
Word 自己的 `ExportAsFixedFormat` 导 PDF**（Word 活渲染 MathType OLE = 用户
所见；LibreOffice 只能渲染 WMF 预览缓存，Symbol 字体替换会产生假象）。
给 vision 模型的检查提示必须**中性**（"逐字报告公式内容"），不要写"是否
正确显示 X"——引导性提示会得到顺着说的误报（实测发生过）。Read 工具读 PNG
只返回 CDN 链接，视觉检查用 analyze_image 传 `%5C` 编码的 URL；**PNG 文件名
不可复用**（CDN 会命中旧缓存），verify_output 已自动加时间戳前缀。

最终提醒用户：在 Word/WPS 中打开抽查，双击公式应进入 MathType 编辑器；
公式字号与正文不匹配时用 MathType → Format Equations 统一。

## 文件清单

- `scripts/fix_environment.py` — 一次性环境修复（MSXML ProgID/TypeLib、AccessVBOM、PATH）
- `scripts/precheck.py` — 转换前预检（公式分布/LaTeX 候选/风险结构/字符集嫌疑）
- `scripts/latex_to_omml.py` — LaTeX 文本 → 原生 OMML（第 2 步；含 groupChr 修复与 align 降级）
- `scripts/convert_via_word_vba.py` — 原生 OMML → MathType OLE，全故事覆盖+断点续转（第 3 步核心）
- `scripts/verify_output.py` — 转换后自动验证（计数+EMF '?'+零文本预览+必检页渲染）
- `scripts/docx_scan.py` — 共享检查库（census/Fallback 剥离/EMF 扫描，仅标准库）
- `scripts/symbol_audit.py` — LaTeX↔MathType 符号审计（301 符号矩阵，产出对照表）
- `scripts/requirements.txt`
- `references/latex-mathtype-map.md` — 符号对照审计表（含已修复缺陷与 MathType 缺口清单）
- `references/mathtype-conversion.md` — 手动操作指引（备选方案、Toggle TeX）
- `references/troubleshooting.md` — 环境坑与失败路径全记录（排障先读）
