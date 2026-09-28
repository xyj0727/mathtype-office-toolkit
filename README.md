# MathType Office Toolkit

简体中文 | [English](README-en.md) | [繁體中文](README-zhTW.md)

一套 **MCP 服务器 + AI Agent 技能（skill）**，让 Claude Code、Claude Desktop、Codex、ChatGPT 等 AI Agent 在 Word 和 PowerPoint 里生成**真正可编辑的 MathType 7 公式**，并支持 MathType 原生公式编号 `(1)`、自动更新的交叉引用和结构化校验。

**默认公式格式为《岩土工程学报》（CJGE）MathType 规范**：五号 10.5 磅、变量 Times New Roman 斜体、小写希腊字母 Symbol 斜体、矢量矩阵黑斜体、编号右对齐、正文引用写作“式（n）”。

版本 **1.2.0**。基于 [felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word)（MIT）fork 并扩展，改进内容见[本 fork 的改进](#本-fork-的改进)。

---

## 目录

- [适用场景](#适用场景)
- [环境要求](#环境要求)
- [核心能力](#核心能力)
- [默认公式格式（CJGE）](#默认公式格式cjge)
- [MCP 工具与 Skill](#mcp-工具与-skill)
- [清单（manifest）写法](#清单manifest写法)
- [静默运行](#静默运行)
- [安装](#安装)
- [验证安装](#验证安装)
- [常见问题](#常见问题)
- [本 fork 的改进](#本-fork-的改进)
- [仓库结构](#仓库结构)
- [致谢与许可](#致谢与许可)

## 适用场景

当你需要让 AI 把**数学公式写进 Word / PowerPoint，而且之后还要在 MathType 里继续修改**时使用：

- 撰写或修改含公式的学位论文、期刊论文（尤其是投《岩土工程学报》等中文期刊）、报告、课程设计。
- 把 TeX（笔记、LaTeX 草稿或大模型输出）转成真正的 MathType 公式，而不是 Word 自带公式（OMath）、图片或纯文本。
- 给行间公式加 `(1)、(2)…` 编号，并在正文里插入“式（2）”这样的引用；增删、移动公式后编号和引用自动更新。
- 把已有文档里所有 MathType 公式**一次性统一成期刊要求的字号和字体样式**。
- 在答辩、组会 PPT 里放和论文同一样式、字号统一的可编辑公式。
- 检查文档里的公式、编号、引用和公式格式是否合规。

**不适用于**：Word 自带公式编辑器（OMath）、LaTeX/PDF 输出、macOS 或网页版 Office。

## 环境要求

| 项目 | 要求 |
|---|---|
| 操作系统 | Windows 10 或 11，需在有桌面的登录会话中运行（Office COM 自动化） |
| Office | 桌面版 Microsoft **Word** 和 **PowerPoint**（Microsoft 365 / Office 2016 及以上，已在 16.0 64 位测试） |
| MathType | 桌面版 **MathType for Windows** 7，可从 [MathType 下载页](https://mathtype.tw/download/) 获取。开发与测试版本为 **MathType-win-zh-7.11.1.462**（`ProductVersion 7.11.1.462`）。只装 **MathType Add-In for Microsoft 365**（任务窗格插件）**不够**，它没有本工具依赖的桌面 OLE 组件、MathType API（`MathPage.wll`）、Word 模板和 PowerPoint 插件 |
| PowerShell | **PowerShell 7 及以上**（`pwsh.exe`）；Office 桥接脚本不能在 Windows PowerShell 5.1 中运行 |
| Python | `PATH` 中有 Python 3（与 Office 位数一致，64 位 Office 用 64 位 Python），并安装 `pip install -r requirements.txt`（`pywin32`、`lxml`） |
| AI 客户端 | 支持 MCP 的 Agent：Claude Code、Claude Desktop、Codex；ChatGPT 需通过远程端点或 Secure MCP Tunnel |

### 终端兼容性

桥接脚本是 PowerShell 7 脚本，可在 PowerShell 7、Windows 上的 Bash（含 Git Bash）或 CMD 中调用，但脚本本身必须由 `pwsh.exe` 执行。

| 当前终端 | 做法 |
|---|---|
| PowerShell 7+ | 直接用 `pwsh.exe` 运行 |
| Windows 上的 Bash（含 Git Bash） | 调用 Windows 的 `pwsh.exe` |
| WSL Bash | 调用 Windows 的 `pwsh.exe`；Linux 版 `pwsh` 无法控制 Windows Office |
| CMD | 用相同参数调用 `pwsh.exe` |
| Windows PowerShell 5.1 | 不要在 PowerShell 5.1 中运行，改用 Git Bash 或 CMD 调用 `pwsh.exe` |
| 没有可用终端或没有 `pwsh.exe` | 安装 PowerShell 7，参见[微软 PowerShell 更新说明](https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6) |

## 核心能力

**Word（`.docx`）**

1. **真正的 MathType 公式**：每个公式都是可编辑的 `Equation.DSMT4` 对象，由 MathType 自身的 TeX 转换生成，双击即可在 MathType 中编辑。
2. **行内公式与行间公式**：行内公式留在句子里，行间公式单独成段。
3. **MathType 原生编号**：带编号的行间公式使用 MathType 的 `MACROBUTTON MTPlaceRef` + `SEQ MTEqn` 域，格式为 `(1)、(2)、(3)`，不含章节号，全文连续。
4. **自动更新的交叉引用**：正文“式（2）”是 MathType 引用（`GOTOBUTTON` + `REF` 书签），公式重新编号后自动跟着变。
5. **一键统一公式格式**：按 MathType 偏好文件（默认 CJGE）把全文每个公式重新排版，效果等同 MathType 的“设置公式格式”，并把偏好写进文档，之后手动插入的公式也沿用同一格式。
6. **表格式行间公式版式**：公式居中、编号靠右，编号和引用仍是 MathType 原生的。
7. **公式后标点**：公式与编号之间自动加“，”或“。”。
8. **结构化 + 格式校验**：检查公式对象、编号、引用、书签、占位符，以及每个公式是否符合 CJGE 字号和样式。
9. **全文自动分类**：Skill 通读全文，判断每个表达式是行内公式、无编号行间公式、带编号行间公式还是引用。
10. **扫描纯文本公式**：自动找出正文和表格里手打的公式和符号（如 `σ1f = (σ3 + Δσ3)·Kp`、`T_ult`），给出建议的 TeX 和处理方式，审阅后一键插入标记并生成清单。可选 CJGE 策略：表达式转 MathType，单个符号转 Times New Roman 斜体加 Word 下标。
11. **大文档分批渲染**：每批一个独立 Word 进程，批后存档；崩溃后可断点续跑，失败的批次自动拆半重试；长任务可在后台运行并随时查询进度。
12. **行内公式不被裁切**：固定行距段落里放了较高的行内公式时，自动改为“最小值”行距。
13. **正文排版与格式报告**：按 CJGE 配置排版正文（页面、标题、正文、表格三线表、参考文献），不动公式；并输出按段落类型汇总的格式报告，可与配置对比列出不符项。

**PowerPoint（`.pptx`）**

14. **可编辑的浮动 MathType 公式**，水平居中，内容逐一核对。
15. **与 Word 同一样式**：同样按 CJGE 偏好排版，整份演示文稿公式字号统一（默认跟随占位文字字号）。

**安全性**

16. 从不修改源文件，结果写入新路径并原子替换。
17. Word、PowerPoint、MathType 全程隐藏静默运行，不弹窗、不占用剪贴板改格式，**不会关闭你已经打开的 Word / PowerPoint**；超时时只结束本工具自己启动的 Office 进程。

## 默认公式格式（CJGE）

依据《岩土工程学报》刊出论文实测和官网《征稿简则》整理，完整规范见 [cjge-format.md](skills/mathtype-for-word/references/cjge-format.md)。

**尺寸与样式**（`config/cjge_equation_preferences.eqp`，自动套用）

| MathType 项 | 设定 |
|---|---|
| Full（常规） | **10.5 磅（五号）** |
| 上下标 / 次级上下标 | 58% / 42% |
| 大运算符 / 次级大运算符 | 150% / 100% |
| 文本、函数名（sin、tan、max…）、数字 | Times New Roman **正体** |
| 变量 | Times New Roman **斜体** |
| 小写希腊字母 | Symbol **斜体** |
| 大写希腊字母、运算符和括号 | Symbol 正体 |
| 矢量、矩阵 | Times New Roman **黑斜体** |

**版式**（`config/cjge_layout_profile.json`）

| 元素 | 格式 |
|---|---|
| 行间公式 | 单独成行、居中；1×3 无边框表格（两侧 72 磅），段前段后 0，行距“最小值 15.6 磅”（公式不会被裁切） |
| 编号 | 半角 `(1)、(2)…`，Times New Roman 五号正体，右对齐，全文连续 |
| 公式后标点 | 后接“式中：”用“，”，句末用“。”，宋体，位于公式与编号之间 |
| 正文引用 | “式（5）”“见式（16）”，全角括号，编号自动更新 |
| 联立方程组 | 右侧一个大括号 `}`，整组只编一个号 |

**写作规则**（由 Skill 在写稿和清单时执行）：描述性下标正体（*W*<sub>t</sub>、*β*<sub>d</sub>，TeX 写 `W_{\mathrm{t}}`）；序号下标 i、j 斜体；数字下标、单位正体；减号为“−”；公式下方另起一行顶格写“式中：”，各项用“；”分隔、以“。”结束；正文中的简单符号直接用 TNR 斜体输入，与公式中写法一致。

> 只想要 MathType 原来的样式？在清单里写 `"equation_preferences": "none"`、`"display_layout": "tab"`、`"reference_brackets": "halfwidth"` 即可。

## MCP 工具与 Skill

**Skill 名称：`mathtype-for-word`**（目录 `skills/mathtype-for-word/`，打包文件 `dist/mathtype-for-word.skill`）：告诉 Agent 怎么通读文档、分类公式、写清单、按 CJGE 规范写 TeX 和“式中”、调用工具并校验。

**MCP 服务器名称：`mathtype-for-word`**（入口 `scripts/run-mcp.ps1` → `scripts/mcp_server.py`，stdio 协议）：

| 工具 | 适用 | 只读 | 作用 |
|---|---|:---:|---|
| `probe_mathtype_word` | Word | ✓ | 检查 Windows、PowerShell、Word COM、MathType 7、Word 模板和 `Equation.DSMT4` 注册 |
| `probe_mathtype_powerpoint` | PowerPoint | ✓ | 同上，另检查 PowerPoint COM 和 MathType 的 PowerPoint 插件 |
| `configure_mathtype_word_defaults` | Word | | 保存默认编号格式 `(1)` 和 MathType 警告偏好；安装后、重装 Office 后各运行一次 |
| `scan_plain_text_math` | Word | | 扫描正文和表格中的纯文本公式与符号，写出可审阅的候选文件（建议 TeX、处理方式；策略 `cjge` / `all`） |
| `prepare_mathtype_markers` | Word | | 按审阅后的候选文件插入 `{{MATH:…}}` 标记并生成清单；简单符号可改为斜体文字加下标 |
| `render_mathtype_word_document` | Word | | 按清单分批生成 MathType 公式、原生编号和引用（可断点续跑、可后台运行），并默认完成 CJGE 排版、表格版式、行距修正、全角引用和校验 |
| `get_mathtype_render_status` | Word | | 查询分批或后台渲染任务的进度、日志和结果 |
| `fix_mathtype_line_spacing` | Word | | 把会裁切行内公式的固定行距段落改为“最小值” |
| `apply_mathtype_equation_preferences` | Word | | 用 MathType 偏好文件（默认 CJGE）统一全文公式的字号和样式 |
| `apply_mathtype_repo_layout` | Word | | 把行间公式改成 1×3 无边框表格版式（默认 CJGE 参数） |
| `validate_mathtype_word_document` | Word | ✓ | 校验公式对象、编号、引用、书签、占位符，并检查公式是否符合 CJGE 格式 |
| `update_mathtype_word_fields` | Word | | 编辑后刷新全部编号域和引用域 |
| `apply_cjge_body_format` | Word | | 按 `config/cjge_body_profile.json` 排版正文（不含公式）：页面、标题、正文、列表、图表题、三线表、参考文献 |
| `report_docx_formatting` | Word | ✓ | 按段落类型汇总字体、字号、行距、缩进、对齐和表格边框；可与配置对比列出不符项 |
| `render_mathtype_powerpoint_presentation` | PowerPoint | | 把占位文本框替换成居中、CJGE 样式、字号统一的 MathType 公式 |
| `validate_mathtype_powerpoint_presentation` | PowerPoint | ✓ | 校验对象命名、居中、内嵌 MathML、公式字号和残留占位符 |

## 清单（manifest）写法

```json
{
  "schema_version": 1,
  "equations": [
    { "id": "stress", "marker": "{{MATH:stress}}", "tex": "\\sigma =\\frac{M y}{I}",
      "layout": "display", "numbered": true, "punctuation": "，" },
    { "id": "group", "marker": "{{MATH:group}}",
      "tex": "\\left.\\begin{array}{l}W_{\\mathrm{t}}=L\\tan \\beta _{\\mathrm{d}}\\\\ V_{\\mathrm{L}}=W_{\\mathrm{b}}H\\end{array}\\right\\}",
      "layout": "display", "numbered": true, "punctuation": "。" }
  ],
  "references": [ { "marker": "{{EQREF:r1}}", "target": "stress" } ]
}
```

- 正文里写 `由式{{EQREF:r1}}可得`，生成后为“由式（1）可得”。
- 可选顶层字段：`reference_brackets`（`fullwidth` 默认 / `halfwidth`）、`equation_preferences`（`.eqp` 路径，默认 CJGE，或 `none`）、`display_layout`（`table` 默认 / `tab`）、`inline_line_spacing`（`at_least` 默认 / `keep`）。
- 渲染参数：`batch_size`（默认 40）、`resume`（默认 true）、`background`（默认 false，公式超过约 150 个时建议打开）、`allow_unresolved_markers`（部分渲染）。
- 文档里的公式是手打纯文本时，不必手写清单：先 `scan_plain_text_math`，审阅候选文件，再 `prepare_mathtype_markers`。
- 联立方程组的 `\\` 后要留一个空格；MathType 的 TeX 导入不支持 `aligned` 和 `\cr`。

## 静默运行

AI Agent 操作 Word、PowerPoint 或 MathType 时必须在后台静默进行：不显示或激活程序窗口、不抢键盘焦点、不弹对话框、不模拟鼠标键盘操作界面。某一步无法静默完成时，应停止并说明原因，而不是接管用户桌面。统一公式格式时直接调用 MathType API 改写文件，不经过剪贴板；只有向 PowerPoint 放入公式时会短暂使用剪贴板。

## 安装

### 让 AI Agent 帮你安装

把下面这段话粘贴给 Claude Code、Claude Desktop、Codex 或 ChatGPT Desktop：

```text
从 https://github.com/xyj0727/mathtype-office-toolkit 安装或升级 MathType Office Toolkit。检测我可用的终端，使用 PowerShell 7、Windows 上的 Bash（含 Git Bash）或 CMD；不要在 Windows PowerShell 5.1 中运行 Office 桥接脚本，如果当前是 5.1，改用 Git Bash 或 CMD 调用 Windows 的 pwsh.exe；在 WSL Bash 中调用 Windows 的 pwsh.exe 而不是 Linux 的 pwsh。如果没有可用终端或没有 pwsh.exe，停下来并提示我按 https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6 安装 PowerShell 7。确认已安装桌面版 MathType for Windows（ProductVersion 7.11.1.462）以及桌面版 Word 和 PowerPoint，运行 pip install -r requirements.txt，安装 mathtype-for-word skill，注册名为 mathtype-for-word 的本地 stdio MCP 服务器，运行 configure_mathtype_word_defaults、两个 MathType 检测和仓库自带测试，保留我现有的 Agent 配置，并列出所有改动过的文件。只有当输出包含可编辑的 Equation.DSMT4 对象且校验返回 ok: true 时才算成功。
```

各平台的具体路径见[安装对照表](skills/mathtype-for-word/references/installation-matrix.md)。

### 手动安装

```console
git clone https://github.com/xyj0727/mathtype-office-toolkit.git
cd mathtype-office-toolkit
pip install -r requirements.txt
```

记下仓库的绝对路径 `<REPO_ROOT>`，然后按你使用的 Agent 安装 skill 并注册 MCP 服务器（保留已有配置）。

### Claude Code

```console
xcopy /E /I "<REPO_ROOT>\skills\mathtype-for-word" "%USERPROFILE%\.claude\skills\mathtype-for-word"
claude mcp add --scope user mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
```

仓库也提供 `.claude-plugin/plugin.json`、`.mcp.json` 和 `dist/mathtype-for-word-plugin.zip`，可按插件方式安装。

### Claude Desktop

在 **Customize > Skills** 上传 `dist/mathtype-for-word.skill`，再把下面内容合并进 `%APPDATA%\Claude\claude_desktop_config.json` 并重启：

```json
{
  "mcpServers": {
    "mathtype-for-word": {
      "command": "pwsh.exe",
      "args": ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", "<REPO_ROOT>\\scripts\\run-mcp.ps1"]
    }
  }
}
```

### Codex

把 `skills/mathtype-for-word` 复制到 `%USERPROFILE%\.codex\skills\mathtype-for-word`，然后：

```console
codex mcp add mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
```

### ChatGPT Desktop

ChatGPT 无法直接启动本地 stdio 服务器，需要远程 MCP 端点或 [Secure MCP Tunnel](https://help.openai.com/en/articles/12584461-developer-mode-and-full-mcp-connectors-in-chatgpt-beta)，并连到运行 Office 的那台 Windows 电脑。

### 安装之后

运行一次 `configure_mathtype_word_defaults`；以后重装 Office 后也要再运行一次。

## 验证安装

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe-pptx
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File tests/run-tests.ps1 -IncludeLiveOffice
```

### AI Agent 快速测试

```text
用已安装的 MathType Office Toolkit 做一次冒烟测试：先运行两个环境检测，再用 evals/fixtures/en-paper-draft.docx 配合 en-word-manifest.json、evals/fixtures/en-presentation-draft.pptx 配合 en-powerpoint-manifest.json，在普通文件夹（不要用 %TEMP%）里生成新的 DOCX 和 PPTX。全程保持 Word、PowerPoint、MathType 隐藏静默，不要覆盖源文件。校验两个输出（Word 校验包含 CJGE 格式检查），报告路径、MathType 对象数、Word 原生编号/引用数和 PowerPoint 的 mathml_verified 数。只有两次校验都返回 ok: true 才算成功。
```

## 常见问题

| 现象 | 解决办法 |
|---|---|
| 渲染卡在“Insert Equation Number”（多见于重装 Office 之后） | 运行 `configure_mathtype_word_defaults`。MathType 的 Word 插件按字符串（REG_SZ）读取 `HKCU\Software\Design Science\DSMT7\WordCommands` 下的值 |
| 重装 Office 后在 Word 里用 MathType“设置公式格式”报错 53 `MathPage.WLL` | 把 `<MathType>\MathPage\64\MathPage.wll`（64 位 Office）复制到 `%APPDATA%\Microsoft\Word\STARTUP` 后重启 Word。本工具直接调用 MathType API，不受影响 |
| 校验报 `equation_format` 错误 | 公式是后来手动添加或修改的，运行 `apply_mathtype_equation_preferences` 后再校验 |
| 刚打开 DOCX 就报“找不到属性 Content” | 文件放在 `%TEMP%` 下，Word 以受保护视图打开；换到普通文件夹 |
| 出现“错误！未找到引用源” | `ZEqnNum…` 书签被删了；通过 MathType 重新插入引用 |
| 公式很多时 Word 报 RPC 错误或渲染超时 | 1.2.0 起自动分批；长任务用 `background: true` 并查询 `get_mathtype_render_status`；失败后用相同参数再调用一次即可续跑 |
| 正文里分式、上标的上半截被裁掉 | 固定行距小于公式高度；渲染会自动修正，也可单独运行 `fix_mathtype_line_spacing` |

完整列表见 [troubleshooting.md](skills/mathtype-for-word/references/troubleshooting.md)。

## 本 fork 的改进

相对上游 [felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word) 1.3.0（完整记录见 [CHANGELOG.md](CHANGELOG.md)）：

- **1.2.0 — 大文档、纯文本公式与正文排版**：
  - 新增 `scan_plain_text_math` / `prepare_mathtype_markers`：自动找出手打的公式和符号并插入标记、生成清单，支持 CJGE 策略（简单符号转斜体文字加下标）；
  - 渲染改为分批执行，每批独立 Word 进程并存档，支持断点续跑、失败批次拆半重试、后台运行，新增 `get_mathtype_render_status`；
  - 部分渲染模式 `allow_unresolved_markers`；
  - 新增 `fix_mathtype_line_spacing`，渲染时自动修正固定行距裁切行内公式；
  - 新增 `apply_cjge_body_format`（CJGE 正文排版）和 `report_docx_formatting`（精简格式报告）；
  - 修复：表格单元格内的标记导致渲染死循环；VML 尺寸为 `1in` 时格式检查崩溃；中文路径导致 MCP 报 `charmap` 编码错误；
  - 改进：桥接日志带时间戳和进度；超时只结束自己启动的 Office 进程；同类校验错误合并为一条。
- **1.1.0 — CJGE 公式格式**：新增 `apply_mathtype_equation_preferences`（调用 MathType API 按偏好文件重排全部公式，效果同 MathType“设置公式格式”，全程静默）；CJGE 偏好与版式配置；公式后标点；全角引用“式（n）”；校验器新增公式格式检查；PPT 公式同样按 CJGE 样式排版；联立方程组 TeX 写法；清单缺少 `references` 时不再报错。
- **1.0.0**：
  - 表格式行间公式版式 `apply_mathtype_repo_layout`（原地转换，编号引用保持 MathType 原生）；
  - PPT 公式字号统一；
  - 不再关闭用户已打开的 PowerPoint（渲染、校验和检测都一样）；
  - MathType 警告偏好改为 REG_SZ 写入（修复重装 Office 后渲染卡死）；
  - PPT 清单省略 `height_points` 不再报错；
  - 桥接脚本改为 UTF-8 输出，中文报错不再破坏 MCP 的 JSON。

## 仓库结构

| 路径 | 用途 |
|---|---|
| `skills/mathtype-for-word/` | 跨 Agent 的 skill（`SKILL.md`）、参考文档（含 `references/cjge-format.md`）和启动脚本 |
| `scripts/mathtype-word.ps1` | Office 自动化桥接脚本（Word、PowerPoint、MathType） |
| `scripts/mcp_server.py`、`scripts/run-mcp.ps1` | stdio MCP 服务器及启动脚本 |
| `scripts/mathtype_prefs.py` | 按偏好文件重排公式、公式格式检查（`apply_mathtype_equation_preferences`） |
| `scripts/repo_layout.py` | 表格式公式版式（`apply_mathtype_repo_layout`） |
| `scripts/mathtype_scan.py` | 扫描纯文本公式、插入标记（`scan_plain_text_math`、`prepare_mathtype_markers`） |
| `scripts/mathtype_batch.py` | 分批、可续跑、可后台的渲染任务（`get_mathtype_render_status`） |
| `scripts/docx_postprocess.py` | 行内公式行距检查与修正（`fix_mathtype_line_spacing`） |
| `scripts/cjge_body_format.py` | CJGE 正文排版与格式报告（`apply_cjge_body_format`、`report_docx_formatting`） |
| `config/cjge_body_profile.json` | CJGE 正文排版配置 |
| `config/cjge_equation_preferences.eqp` | CJGE MathType 偏好（尺寸与样式） |
| `config/cjge_layout_profile.json` | CJGE 行间公式版式 |
| `config/repo_format_profile.json` | word-mathtype-mcp 版式（可选） |
| `config/defaults.json` | Word 默认编号设置 |
| `examples/`、`evals/fixtures/` | 清单示例、中英文测试文件 |
| `tests/` | 静态检查、MCP 协议测试和 Office 实机测试 |
| `dist/` | `mathtype-for-word-plugin.zip`（插件包）、`mathtype-for-word.skill`（skill 包）及 SHA-256 |

重新打包插件：`python scripts/package_plugin.py`。

## 问题反馈

请在 [GitHub Issues](https://github.com/xyj0727/mathtype-office-toolkit/issues) 提交。

## 致谢与许可

[MIT](LICENSE)。

- 原项目：[felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word)，作者 Jia-Ming Zhou (Felimet)，MIT。
- 表格版式最初参考 [word-mathtype-mcp](https://github.com/songsongshuo785-art/word-mathtype-mcp)（Songchongyang，MIT），其配置保留在 `config/repo_format_profile.json`，许可见 `config/LICENSE-word-mathtype-mcp.txt`。
- CJGE 格式依据《岩土工程学报》刊出论文与官网《征稿简则》整理；`cjge_equation_preferences.eqp` 以 MathType 自带的 Times+Symbol 偏好为基础修改。
- MathType 是其所有者的商标，本项目与其无关联。
