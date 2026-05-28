# HEU哈工程风自动化 AI PPT 生成引擎

本项目专为打造**哈尔滨工程大学 (HEU)** 专属风格而生，通过结合大语言模型的能力与底层原生 PPTX 模板引擎，为 AI Agent 赋予自动设计、排版与渲染**严谨科技风的哈工程专属**幻灯片的技能。

最新版采用了**“双阶段解耦（Two-Phase Hybrid）”**工作流：既利用了 Markdown/HTML 强大的排版打样能力，又结合了 Native PPTX 模板 100% 完美的像素级可编辑性。

---

## 🎯 核心功能与工作流

- **Phase 1: 灵活打样预览 (Markdown -> HTML)**：AI 首回合会根据您的需求，发挥强大的长文本排版能力，生成内容极其丰富的 HTML 页面供您在浏览器中直观预览。
- **Phase 2: 精准提炼注入 (JSON -> PPTX)**：一旦您确认 HTML 打样无误，AI 会自动对内容进行“总结与精简缩减”，将其转化为严格的结构化 JSON，并调用 Python 底层脚本 `build_from_template.py`。
- **原生 PPTX 模板动态克隆**：脚本支持 `"normal"` (常规) 和 `"highlight"` (虚线强调框) 两种版式，根据 JSON 需求自动调用底层的 `win32com` 无损克隆相应的模板页面，并用 `python-pptx` 注入内容，生成最终的原生可编辑 PPTX。

---

## 📦 前置依赖 (Prerequisites)

为了确保核心脚本能成功运行，您的机器需要具备：
1. **Node.js** (包含 `npx`)：用于在第一阶段将 Markdown 渲染为 HTML。
2. **Python 3**：用于运行核心数据填充脚本。
3. **Windows 系统与 Microsoft PowerPoint**：底层的 PPTX 页面克隆模块依赖 `win32com.client` 调用本地的 PowerPoint 程序进行无损复制。
4. **Python 依赖包**：
   ```bash
   pip install python-pptx pywin32
   ```

---

## 🛠️ 仓库文件说明

- `templates/template.pptx`：核心母版文件。包含固定占位页面（封面、目录、正文样本 1、正文样本 2、结尾）。AI 是基于样本内容进行智能克隆与替换的。
- `scripts/build_html_ppt.py`：第一阶段打样脚本。调用 Marp 引擎。
- `scripts/build_from_template.py`：第二阶段原生构建脚本。负责读取 JSON，克隆页面，并替换文本框。
- `styles/office_extracted.css`：第一阶段打样使用的样式表。
- `SKILL.md`：AI 的“大脑芯片”。里面用英文明确划定了 Phase 1 和 Phase 2 两步走的工作流规范。
- `output/`：专用于存放生成物（Markdown源码、HTML预览、JSON数据及最终的PPTX成品）的输出文件夹，保证工作区整洁。

---

## 🤖 给 AI Agent 的配置说明

要让你的 AI 具备这种双阶段工作能力，请遵循以下配置步骤：

1. **设置工作区**：将该仓库的文件设置为 Agent 的工作目录。
2. **注入技能指令**：将本仓库中 `SKILL.md` 文件的内容**完整复制**并粘贴给 Agent 作为 **System Prompt（系统提示词）**。
3. 记得开启 AI 执行本地命令行工具的权限。

---

## 🚀 用户使用教程 (Usage)

当 Agent 配置完毕后，你只需要用自然语言向它下达指令。

**示例对话：**
> **用户**："帮我把这篇项目总结写成 PPT，要带封面和目录。"
> 
> **AI (Phase 1 执行)**：
> 1. 读取并生成 Markdown 源码存放在 `output/presentation.md`。
> 2. 自动运行命令：`python scripts/build_html_ppt.py output/presentation.md -o output/presentation.html`。
> 3. 回复用户："HTML 打样已生成，请打开 output/presentation.html 预览。确认无误后请通知我转为 PPTX。"
> 
> **用户**："没问题，提炼精简一下，转 PPTX 吧！"
> 
> **AI (Phase 2 执行)**：
> 1. 总结提取短句，生成 `output/presentation_data.json`。
> 2. 自动运行命令：`python scripts/build_from_template.py output/presentation_data.json templates/template.pptx output/output.pptx`。
> 3. 回复用户：生成完毕，请直接打开 `output/output.pptx`。

---

## 💻 本地手动命令参考 (Commands)

**第一阶段：HTML 生成**
```bash
python scripts/build_html_ppt.py output/presentation.md -o output/presentation.html
```

**第二阶段：PPTX 生成**
```bash
python scripts/build_from_template.py output/presentation_data.json templates/template.pptx output/output.pptx
```
