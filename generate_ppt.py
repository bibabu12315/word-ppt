# -*- coding: utf-8 -*-
"""
文件名称：generate_ppt.py
主要作用：Markdown 转 PPT 脚本
实现功能：
1. 读取指定的 Markdown 文件 (.md)
2. 使用 MarkdownParser 解析 Markdown 内容为结构化数据
3. 使用 PPTGenerator 基于模板生成 PowerPoint 演示文稿
4. 输出最终的 .pptx 文件
"""
import os
import sys
from parser.markdown_parser import MarkdownParser
from ppt.generator import PPTGenerator
from utils.create_template import create_demo_template
from dotenv import load_dotenv

# 加载 .env 文件中的环境变量
load_dotenv()

def main():
    # 1. 配置路径
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Markdown -> PPT 路径
    input_md = os.path.join(base_dir, "input", "generated.md")
    template_pptx = os.path.join(base_dir, "input", "template.pptx")
    output_pptx = os.path.join(base_dir, "output", "result_fixed.pptx")

    print("=== MD转PPT 脚本启动 ===")
    print(f"Input Markdown: {input_md}")
    print(f"Template PPT:   {template_pptx}")
    print(f"Output PPT:     {output_pptx}")

    # 检查输入文件是否存在
    if not os.path.exists(input_md):
        print(f"Error: Input markdown file not found: {input_md}")
        print("Please run generate_md.py first.")
        sys.exit(1)

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
        print("Done! PPT generated successfully.")
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error generating PPT: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
