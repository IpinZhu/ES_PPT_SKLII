import sys
import json
import os
import shutil
import win32com.client
from pptx import Presentation
from pptx.util import Inches

def prepare_slides_win32(pptx_path, slides_data):
    abs_path = os.path.abspath(pptx_path)
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(abs_path, WithWindow=False)
    
    try:
        # Original Slide 3 = Normal Template
        # Original Slide 4 = Highlight Template
        # Original Slide 5 = Ending Template
        
        # We will duplicate from 3 or 4, and move the duplicated slide to be right before the ending slide.
        for s_data in slides_data:
            layout = s_data.get('layout', 'normal')
            if layout == 'highlight':
                new_slide_range = presentation.Slides(4).Duplicate()
            else:
                new_slide_range = presentation.Slides(3).Duplicate()
            
            # Move the new slide to the position right before the ending slide
            # Currently, Ending slide is the last slide. 
            new_slide = new_slide_range(1)
            # Find the index of the ending slide (which is currently the last one because we haven't moved anything yet? No, Duplicate inserts immediately after the source)
            # A safer way: just move it to Count - 1. The ending slide is always pushed to Count.
            new_slide.MoveTo(presentation.Slides.Count - 1)
            
        # After generating all needed slides, delete the original templates (Slide 3 and Slide 4)
        # We delete Slide 3 twice because after deleting 3, the old 4 becomes the new 3.
        presentation.Slides(3).Delete()
        presentation.Slides(3).Delete()
        
        presentation.Save()
    except Exception as e:
        print(f"Error duplicating slides: {e}")
    finally:
        presentation.Close()
        if powerpoint.Presentations.Count == 0:
            powerpoint.Quit()

def replace_text_in_shape(shape, old_text, new_text):
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if old_text in r.text:
                    r.text = r.text.replace(old_text, str(new_text))

def replace_all_text(shape, new_text):
    if not shape.has_text_frame: return
    first = True
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if first:
                r.text = str(new_text)
                first = False
            else:
                r.text = ""

def get_all_shapes(shapes):
    all_shapes = []
    for shape in shapes:
        if shape.shape_type == 6: # GROUP
            all_shapes.extend(get_all_shapes(shape.shapes))
        else:
            all_shapes.append(shape)
    return all_shapes

def process_presentation(json_path, template_path, output_path):
    shutil.copyfile(template_path, output_path)
    
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    slides_data = data.get('slides', [])
    if len(slides_data) == 0:
        print("No content slides to generate.")
        return
        
    # Phase 1: Duplicate slides using win32com
    prepare_slides_win32(output_path, slides_data)
    
    # Phase 2: Populate text using python-pptx
    prs = Presentation(output_path)
    
    # Cover Slide
    cover_slide = prs.slides[0]
    cover_data = data.get('cover', {})
    for shape in cover_slide.shapes:
        replace_text_in_shape(shape, "大大大大标题", cover_data.get('title', ''))
        replace_text_in_shape(shape, "子标题", cover_data.get('subtitle', ''))
        replace_text_in_shape(shape, "2025年5月", cover_data.get('date', ''))
        if shape.has_text_frame and "申 请 人 ：" in shape.text:
            info = f"申 请 人 ：{cover_data.get('reporter', '')}\n" \
                   f"硕士导师：{cover_data.get('instructor', '')}\n" \
                   f"硕士专业：{cover_data.get('major', '')}"
            replace_all_text(shape, info)
            
    # TOC Slide
    toc_slide = prs.slides[1]
    toc_data = data.get('toc', [])
    all_toc_shapes = get_all_shapes(toc_slide.shapes)
    num_idx = 0
    title_idx = 0
    for shape in all_toc_shapes:
        if shape.has_text_frame:
            text = shape.text.strip()
            if text == "01" and num_idx < len(toc_data):
                replace_all_text(shape, f"{num_idx+1:02d}")
                num_idx += 1
            elif text == "项目背景介绍" and title_idx < len(toc_data):
                replace_all_text(shape, toc_data[title_idx])
                title_idx += 1
                
    # Content Slides
    for i, s_data in enumerate(slides_data):
        slide = prs.slides[i + 2]
        for shape in slide.shapes:
            replace_text_in_shape(shape, "➢ 核心痛点分析", s_data.get('subtitle', ''))
            if shape.has_text_frame:
                if "项目背景介绍" in shape.text:
                    replace_all_text(shape, s_data.get('title', ''))
                elif "现有方案的局限性较高" in shape.text:
                    bullets = s_data.get('bullets', [])
                    replace_all_text(shape, "\n".join(bullets))
                    
        # Handle images
        if s_data.get('image') and os.path.exists(s_data['image']):
            slide.shapes.add_picture(s_data['image'], Inches(4), Inches(2.5), width=Inches(5))
            
    prs.save(output_path)
    print(f"Successfully generated {output_path}")

if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("Usage: python build_from_template.py <data.json> <template.pptx> <output.pptx>")
        sys.exit(1)
    process_presentation(sys.argv[1], sys.argv[2], sys.argv[3])
