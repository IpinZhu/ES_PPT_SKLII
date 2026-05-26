# HEU哈工程风自动化 AI HTML-PPT 生成

让 AI 不仅能帮你“写 PPT 文案”，还能直接帮你**“排版并输出能直接全屏播放的 HTML 幻灯片”**，全程无需人类敲击任何代码命令。本项目专为打造**哈尔滨工程大学 (HEU)** 专属风格而生，通过结合大语言模型的能力与 [Marp](https://marp.app/) 渲染引擎，为 AI Agent 赋予自动设计、排版与渲染**严谨科技风的哈工程专属**幻灯片的技能。

---

## 🎯 核心功能

- **哈工程专属视觉体验 (HEU Style)**：深度融合哈尔滨工程大学的主题色彩（如经典的“哈工程蓝”）与严肃、专业的学术汇报排版基因。内置 `office_extracted.css` 主题，无论是开题报告、项目总结还是课堂展示，都能瞬间拿捏“哈工程风”。
- **自动化排版**：AI 根据文本内容自动提炼要点，并应用定制的高级组件（`.text-box`, `.flex-row`, `.grid-3` 等），完美适配学术汇报与工程展示。
- **一键渲染**：支持一键在后台调用 `Marp CLI`，将 Markdown 直接渲染成可直接全屏演出的 HTML 文件。

---

## 📦 前置依赖 (Prerequisites)

为了确保核心渲染脚本 `build_html_ppt.py` 能成功运行，您的机器（或 Agent 的运行环境）需要具备：
1. **Node.js** (包含 `npx`)：用于免安装动态调用 `@marp-team/marp-cli`。
2. **Python 3**：用于运行构建脚本。
3. **AI 终端执行权限**：如果由 AI 自动运行，确保 AI Agent 有调用本地终端执行命令的能力。

---

## 🛠️ 安装教程 (Installation)

1. **克隆本仓库到本地工作区**：
   ```bash
   git clone https://github.com/your-username/marp_engine.git
   cd marp_engine
   ```

2. **核心文件说明**：
   - `build_html_ppt.py`：核心构建脚本，用于调用 Marp。
   - `office_extracted.css`：经过精心微调的高级排版样式表。
   - `SKILL.md`：存放给 AI Agent 读取的工作流与系统提示词（System Prompt）。
   - `logo.png` / `bg1.png`：封面及内容的视觉素材。

---

## 🤖 给 AI Agent 的配置说明

要让你的 AI 具备自动做 PPT 的能力，请遵循以下配置步骤：

1. **配置上下文**：将该仓库的文件设置为 Agent 的工作区目录。
2. **注入系统指令**：将本仓库中 `SKILL.md` 文件的内容**完整复制**并粘贴给 Agent 作为 **System Prompt（系统提示词）** 或 **Skill（技能配置）**。
3. **开启工具**：确保给 Agent 开启 `run_command`（运行命令）和 `write_to_file`（写文件）的工具权限。

---

## 🚀 用户使用教程 (Usage)

当 Agent 配置完毕后，你只需要用自然语言向它下达指令即可。

**示例对话：**
> **用户**："帮我把这篇项目总结写成 PPT，要带封面和目录，排版好看一点，要求文字表述清晰，图片插入合理。"
> 
> **AI (执行流程)**：
> 1. 读取并构思 PPT 结构。
> 2. 生成包含 Marp 语法的 Markdown 源码（如 `presentation.md`）。
> 3. 自动在后台运行命令：`python build_html_ppt.py presentation.md -o presentation.html`。
> 4. 回复用户：生成完毕，请打开 HTML。

等待 AI 提示完成后，去目录下双击打开它刚刚生成的 `presentation.html`，即可使用方向键全屏播放！

---

## 💻 常用命令 (Commands)

如果你想手动运行渲染脚本，可以使用以下命令：

**编译 Markdown 为 HTML 幻灯片**：
```bash
python build_html_ppt.py <输入文件.md> -o <输出文件.html>
```
*例如：*
```bash
python build_html_ppt.py presentation.md -o presentation.html
```

**指定自定义主题**（可选）：
```bash
python build_html_ppt.py presentation.md -o presentation.html -t custom_theme.css
```
