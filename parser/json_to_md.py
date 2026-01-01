# -*- coding: utf-8 -*-
"""
文件名称：parser/json_to_md.py
主要作用：JSON 转 Markdown 转换器
实现功能：
1. 将 WordParser 解析出的 JSON 数据转换为 Markdown 格式
2. 按照 PPT 生成所需的结构组织 Markdown 内容
3. 处理标题、列表等文本格式
"""
import json

class JsonToMdConverter:
    """
    JSON 转 Markdown 转换器
    职责：将结构化的 JSON 数据转换为 Markdown 格式。
    支持模式切换：
    - mode=0: 默认转换 (直接映射)
    - mode=1: AI 增强转换 (预留接口)
    """

    def convert(self, json_data: dict, mode: int = 0) -> str:
        """
        执行转换
        :param json_data: 字典格式的文档数据
        :param mode: 0=默认, 1=AI增强
        :return: Markdown 字符串
        """
        if mode == 1:
            return self._convert_with_ai(json_data)
        else:
            return self._convert_default(json_data)

    def _convert_default(self, json_data: dict) -> str:
        lines = []
        
        sections = json_data.get("sections", [])
        
        # 1. 标题
        # 使用第一个 section 的标题作为主标题 (如果存在)
        main_title = "项目成果汇报"
        if sections and sections[0].get("level") == 1:
             main_title = sections[0].get("title")
        
        lines.append(f"# {main_title}\n")
        
        for i, section in enumerate(sections):
            level = section.get("level", 0)
            title = section.get("title", "")
            blocks = section.get("blocks", [])
            
            # Level 0: Metadata / Preamble
            if level == 0:
                for block in blocks:
                    if block["type"] == "paragraph":
                        lines.append(f"{block['text']}")
                lines.append("") # Empty line
                continue

            # Level 1: Major Section (##)
            if level == 1:
                # 特殊处理：第一个 Section 通常是封面信息
                if i == 0:
                    # 不生成 ## 标题，也不生成分隔线
                    # 直接输出内容块 (即 Key: Value 对)
                    pass
                else:
                    # 在每个一级标题前添加分隔线
                    lines.append("\n---\n")
                    lines.append(f"## {title}\n")
            
            # Level 2: Sub Section (###)
            elif level == 2:
                lines.append(f"\n### {title}\n")
            
            # Level 3+: (####)
            elif level >= 3:
                prefix = "#" * (level + 1)
                lines.append(f"\n{prefix} {title}\n")

            # Process Blocks
            for block in blocks:
                if block["type"] == "paragraph":
                    lines.append(f"{block['text']}\n")
                elif block["type"] == "list":
                    for item in block["items"]:
                        lines.append(f"- {item}")
                    lines.append("") # Empty line after list

        # Add a final separator if needed, or just leave it. 
        # demo.md ends with ---
        lines.append("\n---\n")
        
        return "\n".join(lines)

    def _convert_with_ai(self, json_data: dict) -> str:
        """
        AI 增强转换：全文重构模式
        """
        print("AI mode selected. Restructuring full document with Qwen-Max...")
        from ai.llm_client import LLMClient
        llm = LLMClient()

        # 1. 数据扁平化：将 JSON 转换为带标记的纯文本
        full_text_buffer = []
        
        # 提取元数据（如果有）
        meta = json_data.get("meta", {})
        if meta:
            full_text_buffer.append(f"[文档元数据]\n{json.dumps(meta, ensure_ascii=False)}\n")

        sections = json_data.get("sections", [])
        for section in sections:
            title = section.get("title", "无标题章节")
            level = section.get("level", 0)
            
            # 添加章节标记
            full_text_buffer.append(f"\n[原始章节 (Level {level})：{title}]")
            
            # 提取文本内容
            blocks = section.get("blocks", [])
            for block in blocks:
                if block["type"] == "paragraph":
                    full_text_buffer.append(block["text"])
                elif block["type"] == "list":
                    for item in block["items"]:
                        full_text_buffer.append(f"- {item}")

        full_text_input = "\n".join(full_text_buffer)
        
        # 2. 调用 AI 进行全文重构
        # 考虑到输入可能很长，这里依赖 LLM 的长文本能力
        markdown_content = llm.restructure_content(full_text_input)
        
        print("AI restructuring completed.")
        return markdown_content
