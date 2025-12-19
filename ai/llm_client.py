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
        self.model = "qwen-plus"

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

    def refine_text(self, text: str) -> str:
        """
        专门用于提炼和简化文本的便捷方法
        """
        system_prompt = (
            "你是一个专业的 PPT 内容策划专家。你的任务是将用户提供的长文本重写为 PPT 幻灯片上的核心要点。"
            "原则：极简、有力、结构化。"
            "要求："
            "1. **极度精简**：删除所有废话、连接词和修饰语，只保留核心信息。将文本长度缩减至原来的 30%-50%。"
            "2. **要点化**：不要使用完整的句子，使用短语或关键词。如果内容较多，可以拆分为多个短句。"
            "3. **演讲配合**：记住 PPT 只是提词器，详细解释留给演讲者口述，不要把所有细节都写在 PPT 上。"
            "4. **格式**：直接输出修改后的文本，不要包含任何解释或开场白。"
            "示例："
            "输入：'当前项目进度在等待smt公司贴片上，smt公司贴片完成后需将开发板给到硬件调试，调试完成后才可以开始软件调试。'"
            "输出：'等待 SMT 贴片 -> 硬件调试 -> 软件调试'"
        )
        
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text}
        ]
        
        result = self.chat_completion(messages)
        return result if result else text # 如果失败，返回原文
