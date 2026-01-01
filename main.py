# -*- coding: utf-8 -*-
"""
文件名称：main.py
主要作用：主程序入口 (CLI)
实现功能：
1. 提供命令行接口，串联 Word -> Markdown -> PPT 的完整流程
2. 也可以单独执行各个步骤
3. 包含环境变量加载和基础配置
"""
import os
import sys
import json
from parser.markdown_parser import MarkdownParser
from parser.word_parser import WordParser
from parser.json_to_md import JsonToMdConverter
from ppt.generator import PPTGenerator
from utils.create_template import create_demo_template
from dotenv import load_dotenv

# 加载 .env 文件中的环境变量
load_dotenv()

# --- Configuration ---
USE_LLM = 1  # 0: Default, 1: AI Enhanced

def main():
    # 1. 配置路径
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Word -> JSON 路径
    input_docx = os.path.join(base_dir, "input", "article.docx")
    output_json = os.path.join(base_dir, "build", "article.json")
    
    # JSON -> Markdown 路径
    generated_md = os.path.join(base_dir, "input", "generated.md")
    
    # Markdown -> PPT 路径
    input_md = generated_md
    
    template_pptx = os.path.join(base_dir, "input", "template.pptx")
    output_pptx = os.path.join(base_dir, "output", "result.pptx")

    print("=== Word转PPT 系统启动 ===")
    print(f"Mode: {'AI Enhanced' if USE_LLM == 1 else 'Default'}")
    
    # --- Step 0: Word -> JSON ---
    print("\n[Step 0] Parsing Word to JSON...")
    if os.path.exists(input_docx):
        word_parser = WordParser()
        try:
            data = word_parser.parse(input_docx)
            word_parser.save_json(data, output_json)
            print("Word parsing completed.")
        except Exception as e:
            print(f"Error parsing Word: {e}")
            sys.exit(1)
    else:
        print(f"Warning: {input_docx} not found. Skipping Word parsing.")

    # --- Step 0.5: JSON -> Markdown ---
    print("\n[Step 0.5] Converting JSON to Markdown...")
    if os.path.exists(output_json):
        converter = JsonToMdConverter()
        try:
            with open(output_json, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            md_content = converter.convert(json_data, mode=USE_LLM)
            
            with open(generated_md, 'w', encoding='utf-8') as f:
                f.write(md_content)
            print(f"Markdown generated at: {generated_md}")
        except Exception as e:
            print(f"Error converting JSON to MD: {e}")
            sys.exit(1)
    else:
        print(f"Error: {output_json} not found. Cannot generate Markdown.")
        sys.exit(1)

    print(f"\nInput Markdown: {input_md}")
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
