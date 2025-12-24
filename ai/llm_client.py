# -*- coding: utf-8 -*-
import os
import json
import requests
from typing import List, Dict, Optional

class LLMClient:
    """
    通义千问 (Qwen) API 客户端
    兼容 OpenAI 接口格式
    """
    
    def __init__(self):
        self.api_key = os.getenv("DASHSCOPE_API_KEY")
        if not self.api_key:
            print("Warning: DASHSCOPE_API_KEY environment variable not set.")
            
        self.api_url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
        self.model = "qwen-plus" # 切换为 qwen-plus 模型

    def chat_completion(self, messages: List[Dict[str, str]]) -> Optional[str]:
        """
        发送聊天请求
        """
        if not self.api_key:
            return None

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": self.model,
            "messages": messages,
            "temperature": 0.7  # 控制随机性，0.7 比较平衡
        }

        try:
            response = requests.post(self.api_url, headers=headers, json=payload)
            response.raise_for_status() # 检查 HTTP 错误
            
            result = response.json()
            # 解析 OpenAI 兼容格式的响应
            content = result['choices'][0]['message']['content']
            return content
            
        except Exception as e:
            print(f"LLM API Error: {e}")
            return None

    def restructure_content(self, full_text: str) -> str:
        """
        全文重构：让 AI 分析全文逻辑，重新规划章节，并输出 Markdown
        """
        system_prompt = (
            "你是一个专业的 PPT 架构师。你的任务是将用户提供的原始文档重构为一份结构清晰、内容精炼的 PPT 大纲（Markdown 格式）。"
            "\n\n"
            "### 核心原则\n"
            "1. **结构优先**：\n"
            "   - 如果用户原文已有清晰合理的章节（如背景、方案、成果），请**保留**该结构。\n"
            "   - 如果用户原文结构混乱或未分段，请**大胆重构**，将其拆分为逻辑通顺的章节（建议结构：背景目标 -> 整体方案 -> 关键技术 -> 创新点 -> 成果展示 -> 价值）。\n"
            "   - **数量限制（绝对严格）**：生成的二级标题（##）数量**必须严格控制在 4 到 8 个之间**。如果内容过多，**必须**合并相关章节。绝对不允许超过 8 个章节！\n"
            "2. **内容适度**：\n"
            "   - PPT 页面上的文字不要过于简略，但也要避免大段文字。\n"
            "   - 每个页面（二级标题 ##）下的描述（普通文本）控制在 **20-40字**，作为该页的核心总结。\n"
            "   - **内容块标题（三级标题 ###）**：文字需简练，控制在 **15字以内**。\n"
            "   - **详细内容（列表 -）**：每个三级标题（###）下，**必须包含 2 到 3 个要点（Bullet points）**。每个要点的内容控制在 **30-70字**，确保信息充实且易读。\n"
            "3. **格式严格**：\n"
            "   - 输出必须是标准的 Markdown。\n"
            "   - **一级标题 #**：仅用于封面（如 # 项目汇报）。\n"
            "   - **封面元数据**：如果原文中明确包含【组织/机构名称】（如公司、学校）、【汇报人/主讲人】、【部门/班级】、【日期】等信息，请在 Markdown 开头（# 标题下方）以 `键：值` 的形式列出（如 `汇报人：张三`）。**注意：如果原文未提供某项信息，请绝对不要生成该字段，也不要编造，直接忽略即可，以便保留 PPT 模板中的原始文字。**\n"
            "   - **二级标题 ##**：对应 PPT 的一页幻灯片（如 ## 一、项目背景）。\n"
            "   - **页面描述（关键！）**：在每个二级标题（##）下方，必须紧跟一段**普通文本**（不带任何符号），作为该页面的核心总结（对应 page_desc）。字数严格控制在 **20-40字**。\n"
            "   - **三级标题 ###**：对应页面内的内容块（如 ### 现状分析）。\n"
            "   - **列表 -**：对应内容块下的要点。\n"
            "   - **关键词（新功能）**：在每个三级标题（###）的内容块末尾，可以添加一行 `**关键词：XXX**`，用于提炼该块的核心概念（对应 pageX_keywordY）。\n"
            "   - **分隔符**：在每个二级标题（##）之前必须添加 `---` 分隔符。\n"
            "\n\n"
            "### 示例结构\n"
            "---\n"
            "## 一、项目背景\n"
            "本项目旨在解决光伏发电效率低下的痛点，通过智能优化算法提升能源收益。\n"
            "\n"
            "### 行业现状\n"
            "- 光伏电池板受局部阴影影响大\n"
            "- 现有方案成本高昂\n"
            "**关键词：阴影遮挡**\n"
            "\n\n"
            "### 输入格式说明\n"
            "输入文本中包含 `[原始章节：Title]` 标记，代表用户原始的划分，供你参考。"
        )

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": full_text}
        ]

        print("Sending full content to LLM for restructuring (this may take a while)...")
        result = self.chat_completion(messages)
        
        if not result:
            return "# Error\nAI 生成失败，请检查 API Key 或网络。"
            
        # 清理可能存在的 Markdown 代码块标记
        cleaned_result = result.replace("```markdown", "").replace("```", "").strip()
        return cleaned_result
