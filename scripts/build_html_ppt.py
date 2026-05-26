import os
import subprocess
import argparse
import sys

def build_ppt(input_md, output_html, theme_css="styles/office_extracted.css"):
    """
    Template skill for generating HTML PPT from Markdown using Marp.
    """
    if not os.path.exists(input_md):
        print(f"Error: {input_md} does not exist.")
        sys.exit(1)
        
    print(f"Generating HTML PPT from {input_md}...")
    
    # We use npx to run marp-cli dynamically without global installation
    # --theme applies our extracted style
    # --html allows raw html if needed
    cmd = [
        "npx", "--yes", "@marp-team/marp-cli@latest",
        input_md,
        "--theme", theme_css,
        "-o", output_html,
        "--html",
        "--no-stdin"
    ]
    
    try:
        subprocess.run(cmd, check=True)
        print(f"Success! PPT generated at: {output_html}")
    except subprocess.CalledProcessError as e:
        print(f"Failed to generate PPT: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate HTML PPT using extracted template")
    parser.add_argument("input", help="Input Markdown file")
    parser.add_argument("-o", "--output", help="Output HTML file", default="output.html")
    parser.add_argument("-t", "--theme", help="CSS Theme file", default="styles/office_extracted.css")
    
    args = parser.parse_args()
    build_ppt(args.input, args.output, args.theme)
