# MathType 批量转换指引（WPS 环境）

前置状态：文档中的 LaTeX 文本已被 `scripts/latex_to_omml.py` 转换为 WPS/Word 原生公式（OMML）。本步把这些原生公式批量转成 MathType OLE 对象。

> 菜单名称随 MathType 版本略有差异（英文 Convert Equations / 中文"转换公式"），以用户实际安装的版本为准。

## 方案 A：MathType"转换公式"批量转换（推荐）

前提：用户已安装 **MathType for WPS 版**（MathType 官方提供 WPS 集成；WPS 应用中心/加载项里能找到 MathType 即说明已装好）。

1. 在 WPS 中打开 `*_converted.docx`。
2. 菜单栏找到 **MathType** 选项卡 → **转换公式（Convert Equations）**。
3. 对话框中：
   - 范围（Range）：**整个文档（Whole document）**；
   - 转换方向勾选：**WPS/Word 公式（Word equations / OMML）→ MathType 公式（MathType equations, OLE objects）**；
   - 其余选项保持默认。
4. 点 **转换（Convert）**，等待完成（公式多时需几十秒）。
5. 抽查：双击任一公式应打开 MathType 编辑器。

转换后建议"另存为"最终文件，并提醒用户：
- 公式对象字号若与正文不匹配，可在 MathType 菜单 → 格式化（Format Equations）中统一按文档字号刷一遍；
- 文件体积会增大（每个公式是独立 OLE 对象），属正常现象。

## 方案 B：Toggle TeX 手动转换（公式极少时）

适合只有几个公式、或方案 A 转换后个别公式出问题的场景：

1. 选中一段 LaTeX 文本（不含 `$` 定界符）；
2. 按 **Alt + \**（MathType 的 Toggle TeX），选中文本即变为 MathType 公式；再按一次可切回 TeX。
3. 若无反应，在 MathType → 首选项（Preferences）→ 剪贴板（Cut and Copy Preferences）中启用 TeX 翻译器。

## 方案 C：无 MathType 时的替代

若用户本机没有 MathType，停在第 1 步的产出即可——原生 OMML 公式在 WPS/Word 中已能正常显示和编辑（双击用 WPS 自带公式编辑器），只是不是 MathType 对象。明确告知用户这一差异，由其决定是否安装 MathType 后再做第 3 步。

## 常见问题

| 现象 | 处理 |
|------|------|
| 报错"转换 OMML 到 MathML 出现问题：此操作所需样式表（OMML2MML.XSL）未找到或已过期" | 该 XSL 随 Microsoft Word 安装、WPS 目录里没有。从 `C:\Program Files\Microsoft Office\root\Office16\OMML2MML.XSL`（或 Office15）复制到 WPS 安装目录的 `office6` 文件夹（通常在 `C:\Users\<用户>\AppData\Local\Kingsoft\WPS Office\<版本号>\office6\`，有多个版本目录就每个都放一份），**完全退出 WPS（含托盘）后重启**再转换；机器上没有 Word 时可点报错框的"是"打开 MathType 官方获取说明，或改在装有 Word 的电脑上执行转换 |
| WPS 菜单里没有 MathType 选项卡 | 未装 MathType for WPS，或版本不兼容；去 MathType 官网装 WPS 版，或在 WPS 应用中心启用 |
| 转换对话框里没有"WPS/Word 公式"选项 | 文档里没有 OMML 公式——先确认第 1 步脚本已跑且 `converted > 0` |
| 转换后公式间距/字号异常 | 用 MathType → 格式化公式（Format Equations）按样式统一 |
| 个别公式转换失败 | 用方案 B 的 Toggle TeX 手工处理该个公式 |
