import os
import subprocess
import argparse
import sys
import tempfile

KATEX_INJECTION = (
    '<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css" crossorigin="anonymous">\n'
    '<script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js" crossorigin="anonymous"></script>\n'
    '<script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/contrib/auto-render.min.js" crossorigin="anonymous"\n'
    " onload=\"renderMathInElement(document.body,{delimiters:["
    "{left:'$$',right:'$$',display:true},"
    "{left:'$',right:'$',display:false},"
    "{left:'\\\\(',right:'\\\\)',display:false},"
    "{left:'\\\\[',right:'\\\\]',display:true}"
    "],throwOnError:false,strict:'ignore'});\"></script>\n"
)


def inject_katex(html):
    if "katex" in html.lower():
        return html
    if "</head>" in html:
        return html.replace("</head>", KATEX_INJECTION + "</head>", 1)
    return KATEX_INJECTION + html


def build_ppt(input_md, output_html, theme_css="styles/office_extracted.css"):
    """Generate HTML via Marp then inject client-side KaTeX so that math
    inside raw HTML blocks (which markdown-it does not parse) still renders."""
    if not os.path.exists(input_md):
        print("Error: %s does not exist." % input_md)
        sys.exit(1)
    print("Generating HTML PPT from %s..." % input_md)

    fd, tmp_path = tempfile.mkstemp(suffix=".html")
    os.close(fd)

    npx_bin = "npx.cmd" if os.name == "nt" else "npx"
    cmd = [npx_bin, "--yes", "@marp-team/marp-cli@latest",
           input_md, "--theme", theme_css,
           "-o", tmp_path, "--html", "--no-stdin"]
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print("Failed to generate PPT: %s" % e)
        try: os.remove(tmp_path)
        except OSError: pass
        return

    with open(tmp_path, "r", encoding="utf-8") as f:
        html = f.read()
    html = inject_katex(html)
    # truncate+write in place (works even when file cannot be unlinked, e.g. Cowork mount)
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html)
    try: os.remove(tmp_path)
    except OSError: pass
    print("Success! PPT generated at: %s" % output_html)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate HTML PPT using extracted template")
    parser.add_argument("input")
    parser.add_argument("-o", "--output", default="output.html")
    parser.add_argument("-t", "--theme", default="styles/office_extracted.css")
    args = parser.parse_args()
    build_ppt(args.input, args.output, args.theme)
