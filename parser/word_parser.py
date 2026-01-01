# -*- coding: utf-8 -*-
"""
文件名称：parser/word_parser.py
主要作用：Word 文档解析器
实现功能：
1. 使用 python-docx 读取 Word 文档
2. 提取文档中的标题、段落、列表等内容
3. 识别文档层级结构
4. 将文档内容转换为中间格式 (JSON)
"""
import json
import os
from docx import Document

class WordParser:
    """
    Word 文档解析器
    职责：读取 docx 文件，解析为结构化的 JSON 数据。
    不使用 AI，仅基于文档结构（样式、层级）进行规则解析。
    """

    def parse(self, docx_source) -> dict:
        """
        解析 Word 文档
        :param docx_source: Word 文档路径 (str) 或 文件对象 (bytes stream)
        :return: 解析后的字典数据
        """
        if isinstance(docx_source, str):
            if not os.path.exists(docx_source):
                raise FileNotFoundError(f"Word file not found: {docx_source}")

        document = Document(docx_source)
        
        # 初始化结果结构
        source_name = os.path.basename(docx_source) if isinstance(docx_source, str) else "uploaded_file"
        result = {
            "meta": {
                "source": source_name
            },
            "sections": []
        }

        # 当前正在处理的 section
        current_section = None
        
        # 默认创建一个 section，防止文档开头没有标题的情况
        # 如果文档第一行就是标题，这个默认 section 可能会是空的，最后可以清理掉
        current_section = {
            "level": 0,
            "title": "Preamble", # 前言/导语
            "blocks": []
        }
        result["sections"].append(current_section)

        for para in document.paragraphs:
            text = para.text.strip()
            if not text:
                continue

            style_name = para.style.name
            
            # 1. 处理标题 (Heading 1 - 9)
            if style_name.startswith('Heading'):
                try:
                    # 获取标题级别，例如 "Heading 1" -> 1
                    level = int(style_name.split()[-1])
                except ValueError:
                    level = 1 # 默认处理
                
                # 创建新 section
                new_section = {
                    "level": level,
                    "title": text,
                    "blocks": []
                }
                result["sections"].append(new_section)
                current_section = new_section
            
            # 2. 处理列表 (List Bullet)
            elif 'List Bullet' in style_name or style_name.startswith('List'):
                # 检查上一个 block 是否是 list，如果是则合并
                if current_section["blocks"] and current_section["blocks"][-1]["type"] == "list":
                    current_section["blocks"][-1]["items"].append(text)
                else:
                    # 创建新的 list block
                    new_block = {
                        "type": "list",
                        "items": [text]
                    }
                    current_section["blocks"].append(new_block)
            
            # 3. 处理普通段落 (Normal 或其他)
            else:
                # 默认为普通段落
                new_block = {
                    "type": "paragraph",
                    "text": text
                }
                current_section["blocks"].append(new_block)

        # 清理空的默认 section (如果它没有内容且后面有其他 section)
        if len(result["sections"]) > 1 and not result["sections"][0]["blocks"]:
            result["sections"].pop(0)

        return result

    def save_json(self, data: dict, output_path: str):
        """
        保存解析结果为 JSON 文件
        """
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"JSON saved to: {output_path}")
