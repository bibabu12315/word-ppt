# -*- coding: utf-8 -*-
import re
import os
from typing import List, Optional
from .data_structs import PresentationData, SlideData, ContentBlock

class MarkdownParser:
    """
    Markdown 解析器
    职责：将符合约定的 Markdown 文本解析为 PresentationData 结构化数据。
    """

    def __init__(self):
        # 编译正则表达式以提高性能
        self.re_h1 = re.compile(r'^#\s+(.+)$')          # 一级标题：# 标题
        self.re_h2 = re.compile(r'^##\s+(.+)$')         # 二级标题：## 标题
        self.re_h3 = re.compile(r'^###\s+(.+)$')        # 三级标题：### 标题
        self.re_bullet = re.compile(r'^-\s+(.+)$')      # 列表项：- 内容
        self.re_key_value = re.compile(r'^([^：:]+)[：:]\s*(.+)$') # 键值对：Key: Value

    def parse_file(self, file_path: str) -> PresentationData:
        """
        读取并解析 Markdown 文件
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Input file not found: {file_path}")

        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        return self.parse_lines(lines)

    def parse_lines(self, lines: List[str]) -> PresentationData:
        """
        解析文本行列表
        """
        data = PresentationData()
        current_slide: Optional[SlideData] = None
        current_block: Optional[ContentBlock] = None
        
        # 状态标记：是否还在处理封面区域（在遇到第一个 ## 之前）
        is_cover_section = True
        slide_counter = 0

        for line in lines:
            line = line.strip()
            
            # 1. 跳过空行和分隔符
            if not line or line == '---':
                continue
            
            # 2. 跳过注释 <!-- ... -->
            if line.startswith('<!--'):
                continue

            # 3. 匹配一级标题 (封面标题)
            match_h1 = self.re_h1.match(line)
            if match_h1:
                data.cover_title = match_h1.group(1).strip()
                continue

            # 4. 匹配二级标题 (新页面)
            match_h2 = self.re_h2.match(line)
            if match_h2:
                is_cover_section = False # 结束封面区域
                slide_counter += 1
                
                # 创建新页面
                # 规范化：封面是 Page 0，目录是 Page 1，正文第一页是 Page 2
                # slide_counter 从 1 开始，所以正文第一页应该是 slide_counter + 1
                current_slide = SlideData(
                    title=match_h2.group(1).strip(),
                    page_index=slide_counter + 1 
                )
                data.slides.append(current_slide)
                
                # 重置当前内容块，因为换页了
                current_block = None 
                continue

            # 5. 匹配三级标题 (页面内的小节/内容块)
            match_h3 = self.re_h3.match(line)
            if match_h3:
                if current_slide is None:
                    # 如果还没有页面就出现了三级标题，这属于异常结构，暂时忽略或归入封面(不建议)
                    print(f"Warning: H3 found before H2: {line}")
                    continue
                
                # 创建新内容块
                current_block = ContentBlock(
                    subtitle=match_h3.group(1).strip()
                )
                current_slide.blocks.append(current_block)
                continue

            # 6. 匹配列表项 (Bullet points)
            match_bullet = self.re_bullet.match(line)
            if match_bullet:
                content = match_bullet.group(1).strip()
                
                if current_block:
                    # 如果在内容块内，添加到内容块
                    current_block.bullets.append(content)
                elif current_slide:
                    # 如果在页面内但不在内容块内（直接在 H2 下面的列表），创建一个默认的空标题块
                    # 或者直接忽略，取决于严格程度。这里我们创建一个匿名块。
                    # 为了简单，我们假设必须先有 H3。如果直接有 bullet，创建一个默认块。
                    if not current_slide.blocks:
                        current_block = ContentBlock(subtitle="") # 匿名块
                        current_slide.blocks.append(current_block)
                    else:
                        # 延续上一个块？或者视为新块？
                        # 假设延续上一个块（如果上一个块存在）
                        current_slide.blocks[-1].bullets.append(content)
                else:
                    # 封面区域的列表？通常封面没有列表，忽略
                    pass
                continue

            # 7. 匹配键值对 (仅在封面区域有效)
            if is_cover_section:
                match_kv = self.re_key_value.match(line)
                if match_kv:
                    key = match_kv.group(1).strip()
                    value = match_kv.group(2).strip()
                    data.meta_info[key] = value
                    continue

            # 8. 关键词匹配 (**关键词：XXX**)
            if line.startswith('**关键词：') and line.endswith('**'):
                keyword_content = line.replace('**关键词：', '').replace('**', '').strip()
                if current_block:
                    current_block.keyword = keyword_content
                continue

            # 9. 普通文本段落
            if not is_cover_section and current_slide:
                # Case A: 还没有进入具体的内容块(H3)，视为页面描述
                if current_block is None:
                    # 如果有多行描述，用换行符连接
                    if current_slide.description:
                        current_slide.description += "\n" + line
                    else:
                        current_slide.description = line
                # Case B: 已经在内容块内，视为普通段落内容，添加到 bullets (作为无项目符号的文本)
                else:
                    current_block.bullets.append(line)
                continue

            # 其他情况忽略
            pass

        return data
