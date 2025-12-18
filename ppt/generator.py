# -*- coding: utf-8 -*-
import os
from pptx import Presentation
from pptx.util import Pt
from parser.data_structs import PresentationData

class PPTGenerator:
    """
    PPT 生成器
    职责：读取 PPT 模板，根据 PresentationData 填充内容，保存为新文件。
    """

    def __init__(self, template_path: str, output_path: str):
        self.template_path = template_path
        self.output_path = output_path
        
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")

    def generate(self, data: PresentationData):
        """
        执行生成过程
        """
        prs = Presentation(self.template_path)
        
        # 1. 构建全局 Shape 索引
        # 为了快速查找 pageX_title 等占位符，我们需要遍历所有页面
        # 结构: { "page1_title": shape_obj, "page1_bullet1": shape_obj, ... }
        shape_map = self._build_shape_map(prs)
        
        # 2. 填充封面
        self._fill_cover(shape_map, data)
        
        # 3. 填充正文页面
        self._fill_slides(shape_map, data)
        
        # 4. 保存
        # 确保输出目录存在
        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
        prs.save(self.output_path)
        print(f"PPT generated successfully: {self.output_path}")

    def _build_shape_map(self, prs) -> dict:
        """
        遍历整个 PPT，建立 shape_name -> shape 对象的映射。
        注意：如果不同页面有相同名字的 shape，后面的会覆盖前面的吗？
        根据用户的命名规则 page1_..., page2_...，名字应该是全局唯一的。
        """
        mapping = {}
        for slide_idx, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                
                # 获取 shape 名称 (在 Selection Pane 中看到的名字)
                name = shape.name
                mapping[name] = shape
                # print(f"Found shape: {name} on slide {slide_idx}") # Debug
        return mapping

    def _fill_cover(self, shape_map: dict, data: PresentationData):
        """
        填充封面信息
        尝试匹配常见封面占位符
        """
        # 映射规则：代码中的字段 -> PPT 模板中的 Shape Name
        # 这里我们定义一些可能的命名约定
        cover_mapping = {
            "cover_title": data.cover_title,
            "cover_project": data.meta_info.get("项目名称", ""),
            "cover_presenter": data.meta_info.get("汇报人", ""),
            "cover_dept": data.meta_info.get("部门 / 团队", ""),
            "cover_date": data.meta_info.get("日期", ""),
            "cover_company": data.meta_info.get("公司名称", "")
        }

        for shape_name, text_content in cover_mapping.items():
            if shape_name in shape_map and text_content:
                self._set_text(shape_map[shape_name], text_content)

    def _fill_slides(self, shape_map: dict, data: PresentationData):
        """
        填充正文页面
        """
        for slide_data in data.slides:
            page_idx = slide_data.page_index # e.g., 1, 2, 3
            
            # 1. 填充页面标题: page{i}_title
            title_key = f"page{page_idx}_title"
            if title_key in shape_map:
                self._set_text(shape_map[title_key], slide_data.title)
            else:
                print(f"Warning: Shape '{title_key}' not found in template.")

            # 1.5 填充页面描述: page{i}_desc
            desc_key = f"page{page_idx}_desc"
            if desc_key in shape_map and slide_data.description:
                self._set_text(shape_map[desc_key], slide_data.description)

            # 2. 填充内容块: page{i}_bullet{j}
            for block_idx, block in enumerate(slide_data.blocks):
                bullet_key = f"page{page_idx}_bullet{block_idx + 1}" # bullet1, bullet2...
                
                if bullet_key in shape_map:
                    # 组合文本：小标题 + 列表
                    # 格式：
                    # 小标题
                    # - 要点1
                    # - 要点2
                    
                    full_text_lines = []
                    if block.subtitle:
                        full_text_lines.append(block.subtitle)
                    
                    # 这里我们不手动加 "- "，因为 PPT 的文本框通常自带 bullet 样式
                    # 或者我们可以根据需要添加
                    # 简单起见，直接填入文本，让 PPT 样式控制
                    for b in block.bullets:
                        full_text_lines.append(b)
                    
                    self._set_text_with_formatting(shape_map[bullet_key], block.subtitle, block.bullets)
                else:
                    print(f"Warning: Shape '{bullet_key}' not found in template (for content: {block.subtitle}).")

    def _set_text(self, shape, text):
        """
        设置文本，尽量保留原有格式（字体、颜色等）
        策略：复用第一个段落的第一个 Run，而不是清空整个 TextFrame
        """
        text_frame = shape.text_frame
        
        # 确保至少有一个段落
        if not text_frame.paragraphs:
            text_frame.add_paragraph()
            
        p = text_frame.paragraphs[0]
        
        # 如果段落里有 run，复用第一个
        if p.runs:
            # 修改第一个 run 的文字
            p.runs[0].text = text
            # 清空后续 run 的文本，避免重叠（视觉上删除）
            for i in range(1, len(p.runs)):
                p.runs[i].text = ""
        else:
            # 没有 run，创建一个，它会继承段落/样式默认值
            run = p.add_run()
            run.text = text

    def _set_text_with_formatting(self, shape, title, bullets):
        """
        设置文本，尝试保留模板字体格式
        """
        text_frame = shape.text_frame
        
        # 1. 捕获模板样式 (从第一个段落的第一个 run)
        # 我们假设模板里的占位符（如 "page1_bullet1"）已经设置好了期望的字体
        template_run = None
        if text_frame.paragraphs and text_frame.paragraphs[0].runs:
            template_run = text_frame.paragraphs[0].runs[0]
            
        # 2. 清除内容
        # 注意：clear() 会移除所有段落并重置格式，所以我们需要手动恢复字体
        text_frame.clear()
        
        # 3. 添加小标题
        if title:
            p = text_frame.paragraphs[0] # clear 后会剩下一个空段落
            run = p.add_run()
            run.text = title
            
            # 恢复字体并加粗
            if template_run:
                self._copy_font_style(template_run, run)
            run.font.bold = True 
        
        # 4. 添加列表项
        for bullet in bullets:
            p = text_frame.add_paragraph()
            p.level = 1 # 缩进
            
            run = p.add_run()
            run.text = bullet
            
            # 恢复字体
            if template_run:
                self._copy_font_style(template_run, run)
                # 列表项通常不强制加粗，除非模板本身就是粗体
                # 这里我们不强制设为 False，而是跟随模板

    def _copy_font_style(self, src_run, dest_run):
        """
        将 src_run 的字体样式复制到 dest_run
        """
        if not src_run or not dest_run:
            return
        
        try:
            # 1. 复制字体名称 (包含中文字体处理)
            if src_run.font.name:
                dest_run.font.name = src_run.font.name
                # 处理中文字体 (East Asian typeface)
                self._set_ea_font(dest_run, src_run.font.name)
            
            # 2. 复制大小
            if src_run.font.size:
                dest_run.font.size = src_run.font.size
            
            # 3. 复制加粗/斜体
            if src_run.font.bold is not None:
                dest_run.font.bold = src_run.font.bold
            if src_run.font.italic is not None:
                dest_run.font.italic = src_run.font.italic
                
            # 4. 复制颜色
            if src_run.font.color:
                if src_run.font.color.type == 1: # RGB
                    dest_run.font.color.rgb = src_run.font.color.rgb
                elif src_run.font.color.type == 2: # Theme Color
                    dest_run.font.color.theme_color = src_run.font.color.theme_color
                
        except Exception as e:
            print(f"Warning: Failed to copy font style: {e}")

    def _set_ea_font(self, run, font_name):
        """
        设置东亚字体 (解决中文字体不生效的问题)
        """
        from pptx.oxml.ns import qn
        rPr = run._element.get_or_add_rPr()
        ea = rPr.find(qn('a:ea'))
        if ea is None:
            ea = rPr.makeelement(qn('a:ea'))
            rPr.append(ea)
        ea.set('typeface', font_name)
