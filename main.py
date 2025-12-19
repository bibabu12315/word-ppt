# -*- coding: utf-8 -*-
import os
import sys
from parser.markdown_parser import MarkdownParser
from ppt.generator import PPTGenerator
from utils.create_template import create_demo_template

def main():
    # 1. 配置路径
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_md = os.path.join(base_dir, "input", "template_text", "demo.md")
    template_pptx = os.path.join(base_dir, "input", "template.pptx")
    output_pptx = os.path.join(base_dir, "output", "result.pptx")

    print("=== Word转PPT 系统启动 ===")
    print(f"Input Markdown: {input_md}")
    print(f"Template PPT:   {template_pptx}")
    print(f"Output PPT:     {output_pptx}")

    # 2. 检查并生成模板 (如果不存在)
    if not os.path.exists(template_pptx):
        print("Template not found. Generating demo template...")
        create_demo_template(template_pptx)

    # 3. 解析 Markdown
    print("\n[Step 1] Parsing Markdown...")
    parser = MarkdownParser()
    try:
        presentation_data = parser.parse_file(input_md)
        print(f"Parsed successfully.")
        print(f"Cover Title: {presentation_data.cover_title}")
        print(f"Total Slides: {len(presentation_data.slides)}")
    except Exception as e:
        print(f"Error parsing markdown: {e}")
        sys.exit(1)

    # 4. 生成 PPT
    print("\n[Step 2] Generating PPT...")
    generator = PPTGenerator(template_pptx, output_pptx)
    try:
        generator.generate(presentation_data)
        print("Done!")
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error generating PPT: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
