# MathType Office Toolkit

[English](README.md) | [繁體中文](README-zhTW.md)

一套 **MCP 服务器 + AI Agent 技能（skill）**，让 Claude Code、Claude Desktop、Codex、ChatGPT 等 AI Agent 在 Microsoft Word 和 PowerPoint 里生成**真正可编辑的 MathType 7 公式**，并支持 MathType 原生的公式编号 `(1)`、可自动更新的交叉引用，以及结构化校验。

版本 **1.0.0**。基于 [felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word)（MIT）fork 并扩展，改进内容见[本 fork 的改进](#本-fork-的改进)。

---

## 目录

- [适用场景](#适用场景)
- [环境要求](#环境要求)
- [核心能力](#核心能力)
- [MCP 工具与 Skill](#mcp-工具与-skill)
- [输出格式](#输出格式)
- [静默运行](#静默运行)
- [安装](#安装)
- [验证安装](#验证安装)
- [常见问题](#常见问题)
- [本 fork 的改进](#本-fork-的改进)
- [仓库结构](#仓库结构)
- [致谢与许可](#致谢与许可)

## 适用场景

当你需要让 AI 把**数学公式写进 Word / PowerPoint，并且之后还要在 MathType 里继续修改**时使用：

- 撰写或修改含公式的毕业论文、期刊论文、报告、课程设计（Word）。
- 把 TeX（笔记、LaTeX 草稿或大模型输出）转成真正的 MathType 公式，而不是 Word 自带公式（OMath）、图片或纯文本。
- 给行间公式加 `(1)、(2)…` 编号，并在正文里插入"式 (2)"这样的引用；增删、移动公式后编号和引用会自动更新。
- 在答辩、组会 PPT 里放统一、可编辑的公式。
- 检查文档里的公式、编号和引用是否完好（没有"错误！未找到引用源"、没有残留占位符、没有混入 OMath）。

**不适用于**：Word 自带公式编辑器（OMath）、LaTeX/PDF 输出、macOS 或网页版 Office。

## 环境要求

| 项目 | 要求 |
|---|---|
| 操作系统 | Windows 10 或 11，需在有桌面的登录会话中运行（Office COM 自动化） |
| Office | 桌面版 Microsoft **Word** 和 **PowerPoint**（Microsoft 365 / Office 2016 及以上，已在 16.0 测试） |
| MathType | 桌面版 **MathType for Windows** 7，可从 [MathType 下载页](https://mathtype.tw/download/)获取。开发与测试版本为 **MathType-win-zh-7.11.1.462**（`ProductVersion 7.11.1.462`）。只装 **MathType Add-In for Microsoft 365**（任务窗格插件）**不够**，它没有本工具依赖的桌面 OLE 组件、Word 模板和 PowerPoint 插件 |
| PowerShell | **PowerShell 7 及以上**（`pwsh.exe`）；Office 桥接脚本不能在 Windows PowerShell 5.1 中运行 |
| Python | `PATH` 中有 Python 3。MCP 服务器只用标准库；可选的表格式公式版式还需要 `pywin32`（`pip install -r requirements.txt`） |
| AI 客户端 | 支持 MCP 的 Agent：Claude Code、Claude Desktop、Codex；ChatGPT 需通过远程端点或 Secure MCP Tunnel |

### 终端兼容性

桥接脚本是 PowerShell 7 脚本。你可以在 PowerShell 7、Windows 上的 Bash（含 Git Bash）或 CMD 中调用，但脚本本身必须由 `pwsh.exe` 执行。

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

1. **真正的 MathType 公式**：每个公式都是可编辑的 `Equation.DSMT4` OLE 对象，由 MathType 自身的 TeX 转换（`MTCommand_TeXToggle`）生成，双击即可在 MathType 中编辑。
2. **行内公式与行间公式**：行内公式留在句子里，行间公式单独成段。
3. **MathType 原生编号**：带编号的行间公式使用 MathType 的 `MACROBUTTON MTPlaceRef` + `SEQ MTEqn` 域，格式为简单的 `(1)、(2)、(3)`，不含章节号。
4. **自动更新的交叉引用**："式 (2)"是 MathType 引用（`GOTOBUTTON` + 指向 `ZEqnNum…` 书签的 `REF` 域），公式重新编号后引用跟着变。
5. **域更新**：增删、移动公式后一键刷新全部编号和引用。
6. **可选的表格式版式**：把行间公式排成 1×3 无边框表格（公式居中、编号靠右），编号和引用仍是 MathType 原生的。
7. **结构化校验**：统计 MathType 对象、编号域、引用域、书签和编号连续性；发现 OMath、残留占位符、断掉的引用即报错。
8. **全文自动分类**：Skill 会通读全文，判断每个表达式是行内公式、无编号行间公式、带编号行间公式还是引用。

**PowerPoint（`.pptx`）**

9. **可编辑的浮动 MathType 公式**：水平居中，对象命名为 `MathType_<id>`，并核对内嵌 MathML 与请求一致。
10. **公式字号统一**：整份演示文稿的公式使用同一个数学字号（与 Word 一致），简单公式 `σ = Eε` 和分式看起来大小协调。

**安全性**

11. 从不修改源文件，结果写入新路径并原子替换。
12. Word、PowerPoint、MathType 全程隐藏静默运行，**不会关闭你已经打开的 Word / PowerPoint**。

## MCP 工具与 Skill

**Skill 名称：`mathtype-for-word`**（目录 `skills/mathtype-for-word/`，打包文件 `dist/mathtype-for-word.skill`）。它告诉 Agent 怎么做：通读文档、给公式分类、编写清单（manifest）、套用学术排版规范、调用工具、校验结果。

**MCP 服务器名称：`mathtype-for-word`**（入口 `scripts/run-mcp.ps1` → `scripts/mcp_server.py`，stdio 协议）。工具列表：

| 工具 | 适用 | 只读 | 作用 |
|---|---|:---:|---|
| `probe_mathtype_word` | Word | ✓ | 检查 Windows、PowerShell、Word COM、MathType 7、其 Word 模板和 `Equation.DSMT4` 注册 |
| `probe_mathtype_powerpoint` | PowerPoint | ✓ | 同上，另检查 PowerPoint COM 和 MathType 的 PowerPoint 插件 |
| `configure_mathtype_word_defaults` | Word | | 保存默认编号格式 `(1)` 和 MathType 警告偏好。安装后、重装 Office 后各运行一次 |
| `render_mathtype_word_document` | Word | | 按清单把 `{{MATH:id}}` / `{{EQREF:id}}` 占位符替换成 MathType 公式、原生编号和引用 |
| `apply_mathtype_repo_layout` | Word | | 可选：把行间公式原地改成 1×3 无边框表格版式 |
| `validate_mathtype_word_document` | Word | ✓ | 校验公式对象、编号、引用、书签和占位符 |
| `update_mathtype_word_fields` | Word | | 编辑后刷新全部编号域和引用域 |
| `render_mathtype_powerpoint_presentation` | PowerPoint | | 把占位文本框替换成居中、字号统一的 MathType 公式 |
| `validate_mathtype_powerpoint_presentation` | PowerPoint | ✓ | 校验对象命名、居中、内嵌 MathML、公式字号和残留占位符 |

最简单的 Word 清单示例（更多见 `examples/example-manifest.json`）：

```json
{
  "schema_version": 1,
  "equations": [
    { "id": "stress", "marker": "{{MATH:stress}}", "tex": "\\sigma = \\frac{My}{I}", "layout": "display", "numbered": true }
  ],
  "references": [ { "marker": "{{EQREF:r1}}", "target": "stress" } ]
}
```

## 输出格式

### Word

| 元素 | 格式 |
|---|---|
| 公式对象 | `Equation.DSMT4` OLE，可在 MathType 中编辑；绝不使用 OMath、图片或纯文本 |
| 行内公式 | 嵌在句子里的行内 OLE 对象 |
| 行间公式（默认） | MathType 版式 `<制表符> 公式 <制表符> (n)`，居中制表位 + 右对齐制表位（`MTDisplayEquation` 样式） |
| 行间公式（可选表格版式） | 1×3 无边框表格：两侧列宽 72 磅（不超过版心宽度的 1/4），单元格内边距 0，行高与行距"最小值 20 磅"，垂直居中，无缩进，段前段后 0；公式在中间格居中，编号在右格右对齐 |
| 公式编号 | `(1)、(2)、(3)…`，阿拉伯数字加圆括号，不含章节号，全文统一编号，自动更新。表格版式中编号字体为 Times New Roman / 宋体 12 磅 |
| 引用 | 正文中显示 `(n)`，为 MathType `GOTOBUTTON`/`REF` 域，随编号变化 |
| 排版规范 | 标量和变量希腊字母用斜体；向量用粗体小写；矩阵和张量用粗体大写；函数名、运算符、常数、微分符号和国际单位用正体；数字上下标用正体（IEEE 规范，详见 [academic-equation-style.md](skills/mathtype-for-word/references/academic-equation-style.md)） |
| 行文规范 | 行间公式前要有引出句，公式后用"其中，…"逐一说明新出现的符号和单位 |

### PowerPoint

| 元素 | 格式 |
|---|---|
| 公式对象 | 浮动的 `Equation.DSMT4` OLE，水平居中，放在原占位文本框的位置 |
| 字号 | 整份演示文稿统一一个数学字号，依次取公式的 `font_pt`、清单的 `equation_font_pt`、占位文字的字号（默认 24 磅）；MathType 原始对象（12 磅）按 `font_pt / 12` 缩放 |
| 编号与引用 | 不支持：PowerPoint 没有 MathType 的编号/引用域，也不会用手打数字冒充 |

## 静默运行

AI Agent 操作 Word、PowerPoint 或 MathType 时必须在后台静默进行：不显示或激活程序窗口、不抢键盘焦点、不弹对话框、不模拟鼠标键盘操作界面。某一步无法静默完成时，应停止并说明原因，而不是接管用户桌面。向 PowerPoint 放入公式时会短暂使用 Windows 剪贴板。

## 安装

### 让 AI Agent 帮你安装

把下面这段话粘贴给 Claude Code、Claude Desktop、Codex 或 ChatGPT Desktop：

```text
从 https://github.com/xyj0727/mathtype-office-toolkit 安装或升级 MathType Office Toolkit。检测我可用的终端，使用 PowerShell 7、Windows 上的 Bash（含 Git Bash）或 CMD；不要在 Windows PowerShell 5.1 中运行 Office 桥接脚本，如果当前是 5.1，改用 Git Bash 或 CMD 调用 Windows 的 pwsh.exe；在 WSL Bash 中调用 Windows 的 pwsh.exe 而不是 Linux 的 pwsh。如果没有可用终端或没有 pwsh.exe，停下来并提示我按 https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6 安装 PowerShell 7。确认已安装桌面版 MathType for Windows（ProductVersion 7.11.1.462）以及桌面版 Word 和 PowerPoint，安装 mathtype-for-word skill，注册名为 mathtype-for-word 的本地 stdio MCP 服务器，运行 configure_mathtype_word_defaults、两个 MathType 检测和仓库自带测试，保留我现有的 Agent 配置，并列出所有改动过的文件。只有当输出包含可编辑的 Equation.DSMT4 对象且校验返回 ok: true 时才算成功。
```

各平台的具体路径见[安装对照表](skills/mathtype-for-word/references/installation-matrix.md)。

### 手动安装

1. 克隆仓库，记下它的绝对路径 `<REPO_ROOT>`：

   ```console
   git clone https://github.com/xyj0727/mathtype-office-toolkit.git
   ```

2. 可选（表格式版式需要）：`pip install -r requirements.txt`。
3. 按你使用的 Agent 安装 skill 并注册 MCP 服务器（保留已有的 MCP 配置）。

### Claude Code

```console
xcopy /E /I "<REPO_ROOT>\skills\mathtype-for-word" "%USERPROFILE%\.claude\skills\mathtype-for-word"
claude mcp add --scope user mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
```

仓库也提供 `.claude-plugin/plugin.json`、`.mcp.json` 和 `dist/mathtype-for-word-plugin.zip`，可按插件方式安装。

### Claude Desktop

在 **Customize > Skills** 上传 `dist/mathtype-for-word.skill`，再把下面内容合并进 `%APPDATA%\Claude\claude_desktop_config.json`，然后重启：

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
用已安装的 MathType Office Toolkit 做一次冒烟测试：先运行两个环境检测，再用 evals/fixtures/en-paper-draft.docx 配合 en-word-manifest.json、evals/fixtures/en-presentation-draft.pptx 配合 en-powerpoint-manifest.json，在普通文件夹（不要用 %TEMP%）里生成新的 DOCX 和 PPTX。全程保持 Word、PowerPoint、MathType 隐藏静默，不要覆盖源文件。校验两个输出，报告路径、MathType 对象数、Word 原生编号/引用数和 PowerPoint 的 mathml_verified 数。只有两次校验都返回 ok: true 才算成功。
```

## 常见问题

| 现象 | 解决办法 |
|---|---|
| 渲染卡在"Insert Equation Number"（多见于重装 Office 之后） | 运行 `configure_mathtype_word_defaults`。MathType 的 Word 插件按字符串（REG_SZ）读取 `HKCU\Software\Design Science\DSMT7\WordCommands` 下的值 |
| 刚打开 DOCX 就报"找不到属性 Content" | 文件放在 `%TEMP%` 下，Word 以受保护视图打开；换到普通文件夹 |
| 出现"错误！未找到引用源" | `ZEqnNum…` 书签被删了；通过 MathType 重新插入引用 |

完整列表见 [troubleshooting.md](skills/mathtype-for-word/references/troubleshooting.md)。

## 本 fork 的改进

相对上游 [felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word) 1.3.0：

- **表格式行间公式版式（可选）**：新增工具 `apply_mathtype_repo_layout`，采用 [word-mathtype-mcp](https://github.com/songsongshuo785-art/word-mathtype-mcp) 的公式版式。原段落就地转为表格（`Range.ConvertToTable`，不经剪贴板），编号和引用仍是 MathType 原生的；`validate_mathtype_word_document` 两种版式都认。
- **PowerPoint 公式字号统一**：按统一的数学字号缩放，不再固定 32 磅高；校验器能发现被拉伸的公式和混用字号。
- **不再关闭你的 PowerPoint**：PowerPoint 只能开一个实例；渲染、校验和 `probe_mathtype_powerpoint` 现在只关闭自己打开的文件。
- **重装与中文环境修复**：MathType 警告偏好改为 REG_SZ 写入（原来的 DWORD 会让重装 Office 后渲染卡死）；PPT 清单省略 `height_points` 不再报错；桥接脚本改为 UTF-8 输出，中文报错不再破坏 MCP 的 JSON。

## 仓库结构

| 路径 | 用途 |
|---|---|
| `skills/mathtype-for-word/` | 跨 Agent 的 skill（`SKILL.md`）、参考文档和启动脚本 |
| `scripts/mathtype-word.ps1` | Office 自动化桥接脚本（Word、PowerPoint、MathType） |
| `scripts/mcp_server.py`、`scripts/run-mcp.ps1` | 无第三方依赖的 stdio MCP 服务器及启动脚本 |
| `scripts/repo_layout.py` | 表格式公式版式（`apply_mathtype_repo_layout`） |
| `config/defaults.json` | Word 默认编号设置 |
| `config/repo_format_profile.json` | 表格版式使用的 word-mathtype-mcp 格式设置 |
| `examples/` | 清单示例 |
| `evals/fixtures/` | 中英文 DOCX/PPTX 测试文件 |
| `tests/` | 静态检查、MCP 协议测试和 Office 实机测试 |
| `dist/` | `mathtype-for-word-plugin.zip`（插件包）和 `mathtype-for-word.skill`（skill 包）及 SHA-256 校验文件 |

重新打包插件：`python scripts/package_plugin.py`。

## 问题反馈

请在 [GitHub Issues](https://github.com/xyj0727/mathtype-office-toolkit/issues) 提交。

## 致谢与许可

[MIT](LICENSE)。

- 原项目：[felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word)，作者 Jia-Ming Zhou (Felimet)，MIT 许可。
- 表格版式的格式设置（`config/repo_format_profile.json`）来自 [word-mathtype-mcp](https://github.com/songsongshuo785-art/word-mathtype-mcp)，作者 Songchongyang，MIT 许可，全文见 `config/LICENSE-word-mathtype-mcp.txt`。
- MathType 是其所有者的商标，本项目与其无关联。
