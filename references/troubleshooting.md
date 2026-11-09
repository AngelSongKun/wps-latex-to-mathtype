# 排障手册：MathType 转换环境坑全记录

2026-10-24 在 Windows 11（26200，Office 16 64位，WPS 32位，MathType 7）
的硕士论文实战中逐一验证。**症状与根因的对应关系**是排障最快的入口。

## 症状速查表

| 症状 | 根因 | 修复 |
|------|------|------|
| WPS 中 MathType"转换公式"提示"没有公式可以转换" | WPS 对 MathType 的 OMath 枚举接口不完整；或见下条 | 改走 Word 自动化路径 |
| 对话框"转换 OMML 到 MathML 出现问题：ommml2mml.xsl 未找到或已过期" | OMML2MML.XSL 缺失/版本校验不过 | Word 自带：`C:\Program Files\Microsoft Office\root\Office16\OMML2MML.XSL`（须含签名 `&#x2146`）；WPS 需手动复制到 office6 目录 |
| 弹窗"设置 MathType 在转换公式时出错. 没有要转换的公式"（转换错误） | **引擎未初始化**：MathPage.WLL 的 `MTXFormReset()` 返回 -9999（需先 `MTInitAPI(0,30)` 连接） | 注入 VBA 里先调 `MTInitAPI`（或 `UILib.IsDLLVersionOK`） |
| 宏静默返回、0 个公式被转换、无任何报错 | `New MSXML2.DOMDocument30` 编译失败被 `On Error` 吞掉 | `fix_environment.py` 注册 ProgID + TypeLib |
| COM 错误 0x800A9C68 (VBA 40040) | VBA 编译错误"Can't find project or library"（常为 Declare 的 Lib 解析失败或引用损坏） | 检查 Declare 别名拼写（Windows 版 WLL 导出名就是 `MTInitAPI`，无 Alias）、TypeLib 注册 |
| COM 错误 0x800A0403 (VBA 1027) / 438 "对象不支持该属性或方法" | 经 `Application.Run` 传递**长字符串/对象参数**失败 | 在注入 VBA 内部直调，不要跨 Run 边界传大参数 |
| `MTSetEqnFromLangStr` 返回 -9999（跨进程 Python ctypes 调用） | OLE 对象的 IDispatch 代理跨进程不可用 | 必须在 Word 进程内（注入 VBA）调用 |
| 插入空白公式后 `Selection.InlineShapes` 找不到对象 | `InsertFile` 后选择区不包含新对象 | 用 `rng.InsertFile` + `rng.Expand wdWord` + `rng.InlineShapes(1)`（MathType 原版顺序） |
| Word 的 STARTUP 没有加载 MathPage.wll | 现代 Word 对 WLL 自动加载不可靠；MathType 的自愈函数 `FixMathPageWllAvailability` 依赖错误 53 才触发 | 用户 PATH 加入 `MathType\MathPage\64` 与 `System\64`（Declare 按名解析会查 PATH） |

## 环境事实（本机验证）

- MathType 的转换引擎全在 `MathType\MathPage\{32,64}\MathPage.wll`，WPS 是 32 位进程、
  Office 是 64 位进程，位数不通用。
- dotm 模板（MathTypeCommands 工程）**加密**，无法注入；但 AccessVBOM=1 时可向
  **Normal 工程**注入自建模块。
- `Declare PtrSafe Function ... Lib "MathPage.WLL"` 按名解析顺序含 PATH；
  VBA `Application.Run` 无法可靠传递 71KB 级字符串参数。
- Win11 24H2+（build 26200）移除了 `MSXML2.DOMDocument30/40/60` 的版本化 ProgID
  与 TypeLib 注册（仅剩 `MSXML2.DOMDocument`/`Microsoft.XMLDOM`），这是
  MathType 在新系统上批量转换失效的总根源之一；CLSID 本体仍注册，补 ProgID
  映射即可。
- MathType 的 MathML 字节传递按 ANSI（代码页）处理：VBA `Dim b() As Byte: b = str`
  → Python 等价 `mml.encode('mbcs')`；实测中文环境下希腊字母/上下标正常。

## 验证产物的方法

1. 解包 docx：`word/embeddings/*.bin` 数量、`Equation.DSMT4` 计数、
   `<m:oMath` 残留（应为 0）。
2. bin 尺寸 ≥3072 且不全相同（全为同一尺寸=空白公式）。
3. LibreOffice headless → PDF → pymupdf 渲染公式密集页 → 视觉检查
   （空白框/乱码/上下标错位）。Read 读 PNG 只回 CDN 链接，用
   analyze_image 传 `%5C` 编码 URL。
4. 终验：Word/WPS 打开，双击公式进 MathType 编辑器。

## 2026-10-24 补丁：ANSI 传输丢字符（公式里出现 `?`）

**现象**：转换全部"成功"（ok=N fail=0），但公式中数学负号 `−`(U+2212)、`Â/Ã`、`ŷ/Ŷ`、组合帽子 `x̂`、`𝔼`、`∘`、`∉`、字面 `²/₂`、`R̄` 等 **GBK 之外的字符全部变成 `?`**（ZCode.docx 实测 21 个公式中招）。

**根因**：VBA `b = mml` 按 ANSI(GBK) 代码页把 String 转 Byte()，GBK 无对应字节的字符替换为 `?` 后才传给 `MTSetEqnFromLangStr`。

**修复（已并入 convert_via_word_vba.py）**：VBA 侧新增 `EntityEncode()`，在 transformNode 之后把 MathML 中所有 cp>127 的字符替换为十进制数字实体（如 `&#8722;`），纯 ASCII 无损通过 ANSI 通道，MathType 的 MathML 解析器还原为正确字形。实测 12 种字符类全部修复，`𝔼` 双线体也能正确渲染。

**教训（视觉验收）**：按固定间隔抽页会漏掉问题公式；应先扫描源文档 `m:t` 中非 cp936 可编码字符，把含这些字符的公式页列为**必检页**。

## 2026-10-24 补丁：Word COM 启动竞态

- attach 后立即 `word.Visible=True` 可能报 `RPC_E_CALL_REJECTED`(-2147418111) → 已加 `com_retry()`（60×0.5s）包裹 attach 后所有 COM 调用；
- 打开文档后立即访问 `word.VBE` 可能报 `DISP_E_EXCEPTION`/`E_FAIL`（MathType 加载项仍在初始化）→ 已加 6 次重试；
- 异常退出会遗留 WINWORD 进程并锁住工作副本 → 重跑前 `taskkill /IM WINWORD.EXE /F` 并删除 `~$*.docx` 锁文件。

## 2026-10-24 优化轮新发现（已全部修入脚本/SKILL.md）

- **`doc.Stories` 不存在**：Word OM 的正确属性是 `doc.StoryRanges`（配合 `Range.NextStoryRange` 链接跨节故事片段）。VBA 侧任何对象/属性写错都会以 COM 异常 -2147352567 (DISP_E_EXCEPTION) 形式抛回 Python 且无描述——**VBA 现已内置全局错误处理器，运行期错误以 `VBAERR <号>: <描述>` 字符串返回**，直接可读。Python 侧只对 `result.startswith("ok=")` 才保存产物。
- **僵尸 WINWORD 锁 Normal.dotm → 下一轮 VBE 全程 E_FAIL**：`word.Quit()` 在 `/embedding` 模式下有时杀不干净进程，残留进程持有 Normal.dotm 及 `~$Normal.dotm` 锁，导致下一个全新 Word 实例访问 `word.VBE` 持续报 -2147352567/E_FAIL（表象为"VBE busy"重试耗尽）。修复：① 脚本退出时轮询 Popen 句柄 10s，未退出则 `proc.kill()`（只杀自己启动的进程）；② 启动守卫发现无 WINWORD 时顺手删除陈旧 `~$Normal.dotm`。
- **`tasklist` 输出是 GBK**：zh-CN Windows 下 `subprocess.run(text=True)` 按 UTF-8 解码直接崩（UnicodeDecodeError→stdout=None→TypeError）。进程检测改按**字节**比较 `b"WINWORD.EXE"`。
- **CDN 上传按内容寻址**：Read 工具上传 PNG 的 CDN 对相同内容去重复用旧 URL。文件名唯一并不能防"命中旧图"，但内容相同则无害；不同内容（修复前后对比）必须换文件名才能拿到新 URL。
- **文档结构验证要覆盖全部故事**：只在 document.xml 数 `<m:oMath`/`Equation.DSMT4` 会漏掉脚注/尾注/页眉页脚/文本框部分；verify_output.py 已按 footnotes/endnotes/header/footer 全部 part 计数。

## 2026-10-24 断点续转设计（已实测：limit=5 → 续转 223 → 完成/已完成/校验全过）

- 工作文档打开后**立即 SaveAs2 到输出路径**（输入永不落盘），VBA 每 SAVE_EVERY=20
  个公式 `doc.Save` 一次 → 崩溃最多丢 20 个公式的进度。
- 续转判定：输出 census 满足 `剩余OMML + (输出DSMT4 - 输入DSMT4) == 输入OMML`
  即为一致的部分状态，直接把输出当工作文档继续（已转公式不再是 OMaths，循环天然
  跳过）。输出完成态（剩余=0 且总数吻合）→ 提示已完成退出；不匹配 → 拒绝启动。
- VBAERR/中途异常 → `doc.Close(False)` 丢弃检查点之后的改动，保持输出一致性，
  重跑同一命令即续转。
- 注意：python-docx 不建模脚注/尾注部件——latex_to_omml.py 经
  `doc.part.rels` 取 `RT.FOOTNOTES/ENDNOTES` 关系，解析 blob → lxml 改写 →
  回写 `part._blob`（2026-10-24 实测 2 条脚注 LaTeX 全转换）。

## 2026-11-03 VBE 抖动与自动恢复（实测：第 1 轮失败×15 → 重启第 2 轮成功）

- `word.VBE` 首次访问可能瞬时失败（探测实验：失败 1 次、约 4s 后自愈），但实测存在
  更顽固的形态：同一实例连续 15 次×8s 仍 E_FAIL，**重启 Word 后立即正常**。
- 修复（已入 convert_via_word_vba.py）：VBE 获取 15×8s 重试 → 仍失败抛 VbeUnavailable
  → main 外层第 2 轮：`stop_word`（Quit+确保进程死亡+清 ~$Normal.dotm）→ 重新
  Popen/attach/注入。同一命令无需人工干预自动恢复。
- 教训：VBE 故障有两种形态（瞬时 + 毒化实例），只靠加长重试不够，必须有进程级重启恢复。

## 2026-11-03/04 深挖轮：四个根因修复（全部 Word 实验证实）

1. **MathType 画不出预组合字符（Ŷ→❖ 菱形）**：MathType 7 的 MathML→MTEF
   翻译（WinAllBasicCodePages）对 U+0176（Ŷ）等 Latin-Ext-A 预组合字符缺项，
   Word 里渲染为缺字形菱形（❖）——注意此前"Ŷ 正确"的视觉判定是 vision 模型
   被引导性提示语诱导的误报。**修复**：convert_via_word_vba.py 转换前
   `decompose_accents()` 把预组合字符（Ŷ/ŷ/Ŵ/ŵ）与组合序列（x+U+0302 等 11 种）
   重写为 Word 原生 m:acc 重音结构；mini 实验证明 m:acc 的 hat/tilde/bar
   全部完美渲染（预组合 Ã/U+00C3 本来就正常，Latin-1 在表内）。
2. **latex_to_omml 重建段落摧毁非文本 run**：process_paragraph 旧实现按
   "收集全文→重建段落"工作，会丢掉脚注引用/图形/域代码 run——脚注引用被删后
   脚注孤儿化，Word 打开即丢弃孤儿脚注（表现为"公式凭空少 2 个"）。
   **修复**：外科手术式替换——只改写命中 LaTeX 的 run，其余元素原位保留；
   新增 is_container_para 跳过含嵌套 w:p 的锚段落（旧实现还会摧毁 VML 文本框）。
3. **最后处理公式的预览竞态**：见 2026-11-03 条目（VBA 循环后 1.5s DoEvents
   + Python 存盘前 1s）。
4. **Word 导 PDF vs LibreOffice**：LO 无法运行 MathType OLE 服务器，只能渲染
   WMF 预览缓存（Symbol 字体字节 + 字体替换 → 假象）；**最终验收必须用
   Word 自己的 ExportAsFixedFormat**（活 OLE 渲染 = 用户所见）。
   WMF 文本记录的字节值不能当作码位证据（Symbol 编码 + gbk/cp1252 解码双重失真）。

### 断点续转实战记录
2026-11-03 论文 257 公式转换中 Word 崩溃（RPC 断连），重跑同命令自动从
40/217 检查点继续，最终 257/257——崩溃恢复机制真实生效。

## 2026-11-05 WPS 显示"？"与数据层无关的最终定论（字节级证据）

用户在 WPS 打开转换件：部分符号显示"？"且与相邻字形（帽子/Θ）重叠，双击
MathType 编辑器却完全正常。**这是显示层（WMF 预览）问题，公式数据完好**。

机理（公式 2-6 的预览 image41.wmf 逐条解析实证）：
- MathType 预览 = 几十条「字体+单字节码」的 ExtTextOut 绘制指令。Θ 在
  Symbol 字体是字节 0x51、≤ 是 0xA3、帽子是 0x88、NBSP 是 0xA0，
  同一下标行打包成一条指令 `51 2d 2b 2d 2b 2d 2b a3`。
- **Word/GDI**：按字体的 Symbol 字符集翻译（字节 OR 0xF000 → U+F0A3=≤）→
  渲染正确（Word 导出 PDF 已验证）。
- **WPS（中文系统）**：图元回放时按 GBK 代码页解码文本字节——0x88/0xA3 是
  GBK 双字节序列的前导字节，落单即非法 → 画"？"替代，但坐标仍是原字形槽位，
  替代字形与紧排的邻居（帽子下的 Y、下标行的 Θ）宽度不同 → 视觉重叠。
- **分界线：字节 <0x80 的（Y/F/t/L/−/+/Θ(0x51)/α(0x61)…）GBK 透传正常
  显示；字节 ≥0x80 的（≤(A3) ≥(B3) ±(B1) ≠(B9) ∞(A5) ∈(CE) ⊂(C7)
  ∑(E5) 帽子(88)…）变成"？"**。同一公式两种混排 → "有的能显示有的不能"。
- 本机装齐 MathType 字体仍显示"？"→ 与缺字体无关，是 WPS 解码器行为；
  让对方装 MathType 也无法根治。
- 排查工具：fontTools 读 symbol.ttf 的 (3,0) cmap（F0xx PUA 键 = Adobe
  Symbol 字节位），可离线解码任何预览字节的真实字形。

处置：定稿/打印走 Word（或 Word 导出 PDF）；给 WPS 读者提供原生 OMML 版
（原始输入文件即是）；不要试图"重转一遍"——预览编码方式不变，结果不变。

## 2026-11-05（二）重大修正与修复：Word 导出的 PDF 也有"？"叠字——NBSP 工件

**修正上一条结论**："Word 渲染完全正确"不成立。用户在导出 PDF 中发现重叠内容，
经查（pymupdf rawdict 逐字形扫描）：论文 PDF 含 **47 个 U+003F '?' 字形、7 处
同坐标叠字**（'1?'、'?G'、'?t' 等，页 25/28/33/34/41）——之前视觉抽查全部漏检
（叠字小、引导性提示会掩盖），教训：**显示层验收必须程序化**。

**根因链（逐环实证）**：
1. 原始 OMML 的 m:t 含普通空格（如 "w = 1"）；
2. OMML2MML.XSL 把每个空格转成**含 U+00A0 NBSP 的 `<mi>` 元素**（lxml 离线
   复现确认，`w = 1` → `<mi>w</mi><mi> </mi><mo>=</mo>...` 两个 NBSP）；
3. MathType 把 NBSP 存为字符，预览 WMF 以单字节 `0xA0` 绘制；
4. **现代 symbol.ttf 的 (3,0) cmap 缺 F0A0**（0xF07F-0xF0A0 区段为空）→
   Word 渲染预览时发生字形替换，PDF 里画成字面 '?'，且**与相邻真字形同 origin
   叠印**——即用户看到的"重叠内容"。WPS 里同类字节叠加 GBK 解码问题更糟。

**修复（已入 convert_via_word_vba.py）**：VBA 在 EntityEncode 前执行
`mml = Replace(mml, ChrW(160), ChrW(32))`——普通空格 0x20 在所有预览字体中
都有（画出来不可见）或被 MathML 解析器丢弃。实测：`w = 1` 的预览记录从
`a0 | 31 a0 | 77 | 3d` 变为 `31 | 77 | 3d`，2-6 公式的 `2c a0` 变 `2c`；
修复后 Word-PDF 扫描 0 个 '?'、0 叠字。

**检测手段（已入 verify_output.py --word-pdf / docx_scan.pdf_artifact_scan）**：
用 Word 本体 ExportAsFixedFormat 导出 PDF（LibreOffice 的字体替换路径不同，
**会隐藏这类工件，不能用于显示层验收**），然后 rawdict 逐字形扫描：
- '?' 字形（U+003F、非 CJK 字体）＝替换产物，公式中不应存在；
- 同 origin 两个不同可见字形 ＝ 叠字（剔除目录点线 ' .' 的空白+点同位误报）。

**交付节奏教训**：显示层问题在 Word/WPS/PDF 三处表现不同且互不替代，
最终验收 = --word-pdf 扫描 0 工件 + 必检页人工抽查。

## 2026-11-05（三）叠字工件三轮根因全记录（论文 47'?'→0'?'）

Word-PDF 扫描器逐字形定位后，"?"工件共三个独立根因，全部修复：

1. **NBSP（U+00A0）**→ 预览单字节 0xA0，现代 symbol.ttf 缺 F0A0 字形 → Word
   画 '?' 与相邻字形同位叠印（占 47 个中的 ~40 个）。修复：MathML 中
   NBSP→普通空格。
2. **Latin-1 预组合重音字符（Ã Â）**：MathType 把它们画进 GBK/DBCS 字体
   记录，单字节落 GBK 前导区 → '?'（论文 2-7 的 Ã、2-17 的 Â）。
   **此前"Latin-1 预组合渲染正常"的 mini 实验结论是错的**（场景太简单）。
   修复：decompose_accents 改用 **NFD 通用分解**——任何可规范分解的重音
   字符都重写为 m:acc（多段处理防组合标记泄漏；表外组合符如 cedilla 安全
   跳过保持原文）。分解数从 9 涨到 19。
3. **U+00B7 中点（·）混入 CJK 文本风格**：公式内含中文（如 1[pos_j ~ 候选]）
   时，MathType 把相邻的 · 分进宋体记录，单字节 0xB7 落 GBK 前导区 → '?'。
   修复：MathML 中 ·(U+00B7)→⋅(U+22C5 点积符)——数学上本就应使用点积符，
   且 Symbol 字体路径渲染稳定（301 符号审计验证）。

**检测器精化**（docx_scan.pdf_artifact_scan）：
- 叠字判定加"同字体同字号"条件（撇号′+∈ 等窄字宽舍入同位是合法排版，曾致
  92 处误报）；
- "？"区分工件与正文原生（参考文献标题 "forecasting?[C]" 原文就有）——
  按**空间左侧最近字符是否字母**判定（按提取顺序判定会被 span 切分骗过）。

**最终结果**：论文 257 公式，Word-PDF 渲染工件 0 个 '?'、0 处叠字，
程序化终检 PASS。经验：显示层问题必须在 Word 本体导出的 PDF 上做逐字形
程序化扫描，视觉抽查（引导性提示 + 小叠字）已两次漏检。

## 2026-11-05（四）结构优化：宏插件架构 + 最后两个 MathType 缺口修复

### 1. 宏插件架构（替代 Normal.dotm 注入）

旧路径每次转换都要 VBE 注入 Normal.dotm→跑完删除——VBE 的 E_FAIL 竞态是
全管线最不可靠环节（实测 8 次运行 4 次触发，每次浪费 2-4 分钟重启循环），
且 Word 崩溃时注入的组件会残留在 Normal.dotm 里。

**新架构**：转换宏一次性安装为 Word Startup 文件夹的
`%APPDATA%\Microsoft\Word\STARTUP\PyMathTypeConv_vN.dotm`（与 MathType
自身集成同一机制，默认受信任位置）。此后每次转换**零 VBE 访问**：
attach → 打开文档 → `Application.Run` 直调宏。VBE 仅在首次安装/版本升级
时使用（fail-fast 探测 + 重启循环）。VBA 变更时 bump `ADDIN_VER`，
旧插件自动删除重建。卸载 = 删除该 dotm。
- `PyConvertAll` 签名改为 `Optional testLimit`——`--limit` 经
  Application.Run 参数传入（旧实现靠字符串替换注入常量，插件架构下不可行）。
- 实测：插件路径转换 3 公式 8 秒；批量两文件 45 秒；论文 257 公式全流程
  正常。

### 2. Word 崩溃自动续转

RPC 断连（-2147023170/-2147023174/-2147417848/-2147418107）不再中止脚本：
`run_once` 转为 `CrashDuringRun`，驱动层自动重开 Word 并从检查点续转
（每文件最多 3 次，续转前做 census 一致性校验）。

### 3. 批量模式

`convert_via_word_vba.py a.docx b.docx ...` 顺序转换多文件（每文件独立
输出/独立续转校验）；`-o` 仅限单输入。

### 4. MathType 最后两个数据缺口修复（命名实体路径）

`\gg`(≫) 与 `\vdash`(⊢) 作为原始码点/数字实体时被 MathType 的
WinAllBasicCodePages 翻译器拒绝（'?' 或整条空白，引擎 stat=-9999）。
**改写为 MathML 命名实体 `&Gg;`/`&vdash;` 后完美渲染**；孤立算符公式
还需 `<mo>`→`<mi>` 语义转换（否则 MathType 接受但写入空内容，MTEF 停留在
DSMT6 空白模板——RESULT 仍报 ok，只能靠 MTEF 长度或目检发现）。
修复已固化进 VBA（三行 Replace），mini6 四变体 + Word 导出目检全部通过。
301 符号审计的"2 个缺口"清零。

## 症状速查决策树（先看这里，再往下翻历史记录）

```
WPS/Word/PDF 里公式显示异常？
├─ 双击 MathType 编辑器内容正常？
│   ├─ 是 → 显示层问题（预览/渲染），数据完好
│   │   ├─ WPS 里 '?' + 与邻居重叠 → 预览 GBK 解码（2026-11-05 条），
│   │   │   定稿走 Word；WPS 读者看原生 OMML 版
│   │   └─ Word 导出 PDF 里 '?'/叠字 → 检查是否旧版转换件
│   │       （NBSP/ÃÂ/· 修复前的），用当前脚本重转 + --word-pdf 验证
│   └─ 否（编辑器里也坏）→ 数据层问题 → 看下方历史根因
├─ 转换 RESULT 里 fail>0 或 stats=-9999@k → 引擎拒绝 → 按序号定位公式
├─ 转换 RESULT ok=N 但 verify 计数不符 → 见"段落重建摧毁 run"（2026-11-05）
├─ VBE busy/插件构建失败 → 已有重启循环；持续失败查 AccessVBOM 与
│   ~$Normal.dotm 锁（本条 2026-10-24 章节范围内的 Normal 修复）
└─ 公式"凭空少了几个" → 孤儿脚注被 Word 丢弃（2026-11-05 段落重建条）
```
