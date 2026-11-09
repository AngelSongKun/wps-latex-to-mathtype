# LaTeX ↔ MathType 符号对照审计表

生成方式：`scripts/symbol_audit.py`——301 个符号先过 LaTeX→OMML 层
（latex2mathml + mathml2omml + 本 skill 的修复补丁），再把全部符号构成
审计文档跑真实 Word+MathType 转换，用 EMF 预览扫描逐符号判定。

**结论（2026-11-03，MathType 7 / Word 16 / Win11 26200 实测）**：

- LaTeX→OMML 层：**301/301 全部通过**（修复 groupChrPr 畸形 XML 与
  align 系环境之前为 291/301）。
- MathType 层：**301/301 全部正确**（2026-11-05 更新：此前两个缺口
  `\gg`/`\vdash` 已通过 **MathML 命名实体路径**修复——原始码点被 MathType
  的 WinAllBasicCodePages 翻译器拒绝/画'?'，改写为 `&Gg;`/`&vdash;`
  （并转 `<mi>` 语义避免孤立算符被丢弃）后完美渲染）：

| LaTeX | 字符 | 状态 |
|---|---|---|
| `\gg` | ≫ (U+226B) | ✅ 已修复（&Gg; 命名实体） |
| `\vdash` | ⊢ (U+22A2) | ✅ 已修复（&vdash; 命名实体 + mi 语义） |

- `\quad`/`\qquad` 输出空白属**正常**（纯空白符号）。
- 转换语义降级（内容不丢但形式变化）：

| 输入 | 行为 |
|---|---|
| `\begin{align}`/`aligned`/`split`/`flalign`/`eqnarray` | 自动降级为 `\begin{gather}`（多行居中）：对齐位 `&` 丢弃；align 的虚假编号 (1)(2) 消除 |
| `\mathbb{1}` | 降级为普通 `1`（Unicode 无对应空心数字 1 的映射被 latex2mathml 丢弃） |
| `\text{ }`/`\mbox{}` 内空格 | 变为 NBSP(U+00A0)，MathType 渲染正常 |

## 已修复的链条缺陷（本次审计发现并修复）

1. **MathType 画不出预组合重音字符**：Ŷ(U+0176) 等 Latin-Ext-A 字符经
   MathType 的 WinAllBasicCodePages 翻译后渲染为缺字形菱形（Word 实测），
   组合序列（x+U+0302）退化为分离的 ^。**修复**：转换前 decompose_accents()
   把它们重写为 m:acc 重音结构（hat/tilde/bar/dot/vec 等 11 种，Word 实测
   完美渲染）。预组合 Ã/Â（Latin-1）本就正常，无需处理。
2. **mathml2omml groupChrPr 畸形 XML**：`\bar` `\vec` `\widetilde`
   `\overbrace` `\underbrace` `\overrightarrow` `\check` 等所有组字符
   结构在库输出中闭合标签错写为 `</m:groupChr>`，导致整文档解析失败
   → `latex_to_omml.py` 中 `_GROUPCHR_BUG` 正则修复。
3. **latex2mathml align 环境缺陷**：`aligned` 产出未转义的 `&`（非法
   XML）、`align` 注入虚假编号 → 预处理降级为 `gather`。
4. **latex_to_omml 重建段落摧毁非文本 run**：脚注引用/图形/域代码在旧
   "全文重建"实现中被删除（脚注引用被删→孤儿脚注→Word 打开丢弃→公式
   凭空消失）；含嵌套 w:p 的 VML 文本框锚段落也被摧毁。**修复**：外科
   手术式替换（只改写命中的 run）+ 容器段落跳过。
5. **末公式预览竞态**：转换循环最后一个处理的公式（文档第一个）的 OLE
   预览缓存与最终存盘竞争，偶发整条空白 → VBA 循环后 1.5s DoEvents +
   Python 存盘前 1s 等待。

## 检测手段说明

- **EMF/WMF 文本记录扫描**（`docx_scan.question_scan`）：MathType 画不出
  的字符以字面 `?` 绘制，文本记录可检出（如 `\gg`）。**局限**：部分公式
  预览为位图绘制（StretchDIB），文本记录为空——问号藏进位图时扫不到，
  须配合渲染目检。
- **零文本预览告警**（`verify_output.py`）：正常公式预览应含文本记录；
  零文本 = 空白或位图绘制，输出清单供人工目检（不做硬失败，位图绘制
  可能合法）。
- **字符集嫌疑页**（`precheck.py` → `verify_output.py --render`）：含
  GBK 外字符的公式所在页必检。
