# -*- coding: utf-8 -*-
from dataclasses import dataclass, field
from typing import List, Dict

@dataclass
class ContentBlock:
    """
    内容块，对应 Markdown 中的三级标题 (###) 及其下方的列表内容。
    对应 PPT 中的 pageX_bulletY
    """
    subtitle: str = ""  # 小标题 (### 后的文本)
    bullets: List[str] = field(default_factory=list)  # 列表项内容 (- 后的文本)

@dataclass
class SlideData:
    """
    页面数据，对应 Markdown 中的二级标题 (##) 及其包含的所有内容。
    对应 PPT 中的 pageX
    """
    title: str = ""  # 页面主标题 (## 后的文本)
    description: str = "" # 页面描述文本 (## 下方的普通段落)
    blocks: List[ContentBlock] = field(default_factory=list)  # 页面内的内容块列表
    page_index: int = 0  # 逻辑页码，从 1 开始 (page1, page2...)

@dataclass
class PresentationData:
    """
    整个演示文稿的数据结构。
    """
    # 封面元数据 (如：项目名称、汇报人、日期等)
    # 解析自 Markdown 顶部的键值对
    meta_info: Dict[str, str] = field(default_factory=dict)
    
    # 封面主标题 (# 后的文本)
    cover_title: str = ""
    
    # 正文页面列表
    slides: List[SlideData] = field(default_factory=list)
