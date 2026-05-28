# Agent Operating Manual

**Hello, fellow AI Agent!** If you are reading this, you are operating within the HEU-styled PPTX automation engine. This document will help you understand the architectural context of the repository.

## 🏛️ Architecture Overview

We have migrated away from Markdown/HTML-based slide rendering (like Marp). This repository uses a **Native PPTX Template Engine**.
The user has provided a beautifully formatted PowerPoint file named `template.pptx`. It serves as the single source of truth for all graphical layouts, background vectors, corporate colors (HEU Blue), and text placements.

**Your Goal**: To translate the user's raw thoughts, notes, or requests into a strict JSON format, and then trigger our local Python script to fuse that JSON with `template.pptx`.

## 📂 Key Files and Their Roles

1. **`template.pptx`**: The master template.
   - Slide 1: Cover Slide
   - Slide 2: Table of Contents (TOC)
   - Slide 3 & Slide 4: Content Slides (Identical layout)
   - Slide 5: Ending/Thanks Slide

2. **`scripts/build_from_template.py`**: The magic fusion script.
   - **What it does**: It reads your generated JSON, uses `win32com.client` to dynamically clone Slide 4 as many times as needed to fit the content, and then uses `python-pptx` to perform XML-level text replacement.
   - **Why it matters**: It ensures that your JSON text replaces the dummy text in the template *without losing any font sizes, weights, or colors*.

3. **`SKILL.md`**: Your strict operating instructions.
   - It contains the exact JSON schema you must output.
   - It provides the exact bash command you must run.

## 🚀 The Operational Loop

When a user asks you to "make a presentation":
1. **Analyze**: Read their request and outline the presentation in your context window.
2. **Draft JSON**: Create a `.json` file (e.g. `presentation_data.json`) matching the schema in `SKILL.md`. Ensure bullet points are short to prevent text overflow in the fixed-size text boxes of the template.
3. **Execute**: Call the build script.
   ```bash
   python scripts/build_from_template.py <your_json_file.json> template.pptx <output_file_name.pptx>
   ```
4. **Deliver**: Inform the user that the PPTX is ready.

You do not need to edit the python scripts or modify the template. Simply feed structured data into the pipeline and enjoy the perfectly formatted output!
