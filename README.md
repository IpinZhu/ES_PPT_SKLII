# HEU哈工程风自动化 AI PPT 生成引擎

本项目专为打造**哈尔滨工程大学 (HEU)** 专属风格而生，通过结合大语言模型的能力与底层原生 PPTX 模板引擎，为 AI Agent 赋予自动设计、排版与渲染**严谨科技风的哈工程专属**幻灯片的技能。

本引擎已完全脱离易变形的 HTML 中间层，采用**原生 PPTX 模板注入技术**。它能保证在 AI 自动填充大量文字时，依然 100% 还原您预设的字体、颜色、字号和精美背景图形，并保持幻灯片的完全可编辑性。

---

## 🎯 核心功能

- **哈工程专属视觉体验 (HEU Style)**：使用精美的 `template.pptx` 原生模板，预设好了哈工程蓝、红条幅边框、水印图腾和标准化的占位符。
- **动态无限扩展**：内置幻灯片克隆引擎，无论 AI 总结出多少页大纲，都能自动从模板中克隆相应排版的页面并自动对齐。
- **所见即所得的可编辑性**：AI 生成的最终成品是纯正的 `.pptx` 格式，所有的标题、目录、正文都是标准的文本框，随时可以二次排版和修改。

---

## 📦 前置依赖 (Prerequisites)

为了确保核心渲染脚本 `scripts/build_from_template.py` 能成功运行，您的机器需要具备：
1. **Python 3**：用于运行核心数据填充脚本。
2. **Windows 系统与 Microsoft PowerPoint**：底层的页面克隆模块依赖 `win32com.client` 调用本地的 PowerPoint 程序进行无损复制。
3. **Python 依赖包**：
   ```bash
   pip install python-pptx pywin32
   ```

---

## 🛠️ 仓库文件说明

- `template.pptx`：核心母版文件。包含 5 张固定占位页面（封面、目录、正文样本、结尾）。**请勿随意删除里面的测试文本**，AI 是基于文本内容进行智能替换的。
- `scripts/build_from_template.py`：核心构建脚本。负责读取 JSON，克隆页面，并替换文本框。
- `SKILL.md`：AI 的“大脑芯片”。里面用英文写满了让 AI 执行这个工作流的结构化指令（Prompt）。
- `agent.md`：一份专门写给 AI Agent 阅读的工作流介绍文档，帮助 Agent 更好地理解这个工程的架构。
- `assets/`：存放所有背景图和 Logo 素材。

---

## 🤖 给 AI Agent 的配置说明

要让你的 AI 具备自动做 PPT 的能力，请遵循以下配置步骤：

1. **设置工作区**：将该仓库的文件设置为 Agent 的工作目录。
2. **注入技能指令**：将本仓库中 `SKILL.md` 文件的内容**完整复制**并粘贴给 Agent 作为 **System Prompt（系统提示词）**。
3. **阅读架构图**：如有需要，可以提示 AI 优先阅读 `agent.md` 以了解项目运作机理。

---

## 🚀 用户使用教程 (Usage)

当 Agent 配置完毕后，你只需要用自然语言向它下达指令。

**示例对话：**
> **用户**："帮我把这篇项目总结写成 PPT，要带封面和目录，排版好看一点，要求文字表述清晰。"
> 
> **AI (执行流程)**：
> 1. 读取并提炼您的文本结构。
> 2. 在本地生成一个结构化的 `presentation.json` 数据文件。
> 3. 自动在后台运行命令：`python scripts/build_from_template.py presentation.json template.pptx output.pptx`。
> 4. 回复用户：生成完毕，请直接打开 `output.pptx`。

去目录下双击打开它刚刚生成的 PPT 即可！

---

## 💻 本地调试命令 (Commands)

如果你想自己手写 JSON 然后手动运行脚本，可以使用以下命令：

```bash
python scripts/build_from_template.py <您的数据文件.json> template.pptx <输出文件.pptx>
```

---

## 🤖 协作模型 (Collaborating Model)

本项目由 **Antigravity** (Google DeepMind) 参与协作开发，提供自动化工作流设计与核心注入脚本的编写支持。
