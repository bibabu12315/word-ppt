# -*- coding: utf-8 -*-
"""
文件名称：generate_md.py
主要作用：Word 转 Markdown 脚本
实现功能：
1. 读取指定的 Word 文档 (.docx)
2. 使用 WordParser 解析文档结构
3. 使用 JsonToMdConverter 将解析结果转换为 Markdown 格式
4. 将生成的 Markdown 保存到 output 目录，供后续编辑或生成 PPT 使用
"""
import os
import sys
import json
from parser.word_parser import WordParser
from parser.json_to_md import JsonToMdConverter
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
    
    print("=== Word转MD 脚本启动 ===")
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
            print("\nSuccess! You can now edit the markdown file before generating PPT.")
        except Exception as e:
            print(f"Error converting JSON to MD: {e}")
            sys.exit(1)
    else:
        print(f"Error: {output_json} not found. Cannot generate Markdown.")
        sys.exit(1)

if __name__ == "__main__":
    main()
