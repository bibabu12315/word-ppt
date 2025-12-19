# -*- coding: utf-8 -*-
import os
import copy
import re
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from parser.data_structs import PresentationData
from utils.ppt_utils import duplicate_slide, move_slide, duplicate_shape

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
        
        # 验证模板结构：必须至少有 4 页 (封面, 目录, 内容, 结尾)
        if len(prs.slides) < 4:
            print("Warning: Template should have at least 4 slides (Cover, TOC, Content, End).")
        
        # 1. 填充封面 (Slide 0)
        print("Generating Cover...")
        self._fill_cover(prs.slides[0], data)
        
        # 2. 填充目录 (Slide 1)
        print("Generating TOC...")
        self._fill_toc(prs.slides[1], data)
        
        # 3. 生成并填充内容页
        print("Generating Content Slides...")
        content_template_index = 2
        
        num_chapters = len(data.slides)
        if num_chapters > 0:
            # 3.1 复制 Slide 2 (N-1 次)
            for i in range(1, num_chapters):
                duplicate_slide(prs, content_template_index)
            
            # 3.2 移动 End 页到最后
            move_slide(prs, 3, len(prs.slides) - 1)
            
            all_titles = [s.title for s in data.slides]
            
            for i, chapter_data in enumerate(data.slides):
                slide_index = 2 + i
                slide = prs.slides[slide_index]
                self._fill_content_page(slide, chapter_data, all_titles, i)
                
        # 4. 保存
        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
        prs.save(self.output_path)
        print(f"PPT generated successfully: {self.output_path}")

    def _build_shape_map(self, slide) -> dict:
        """
        建立单个幻灯片的 shape_name -> shape 映射
        """
        mapping = {}
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            mapping[shape.name] = shape
        return mapping

    def _fill_cover(self, slide, data: PresentationData):
        """
        填充封面信息
        """
        shape_map = self._build_shape_map(slide)
        cover_mapping = {
            "cover_title": data.cover_title,
            "cover_company": data.meta_info.get("公司名称", ""),
            "cover_project": data.meta_info.get("项目名称", ""),
            "cover_presenter": data.meta_info.get("汇报人", ""),
            "cover_dept": data.meta_info.get("部门 / 团队", ""),
            "cover_date": data.meta_info.get("日期", "")
        }

        for shape_name, text_content in cover_mapping.items():
            if shape_name in shape_map and text_content:
                self._set_text(shape_map[shape_name], text_content)

    def _fill_toc(self, slide, data: PresentationData):
        """
        填充目录页
        动态生成: page1_title_num (01), page1_title (标题)
        """
        shape_map = self._build_shape_map(slide)
        
        # 查找原型 Shape
        proto_num = shape_map.get("page1_title_num")
        proto_title = shape_map.get("page1_title")
        
        if not proto_num or not proto_title:
            print("Warning: TOC prototypes (page1_title_num, page1_title) not found.")
            return

        # --- Dynamic Layout Calculation ---
        # Access presentation object via slide -> part -> package -> presentation_part -> presentation
        # Or simpler: since we created 'prs' in generate(), we could pass it.
        # But here we only have 'slide'.
        # In python-pptx, slide.part.package.presentation_part.presentation gives the Presentation object.
        # However, let's try a safer way if possible.
        # Actually, we can just use a fixed height if we can't get it, but getting it is better.
        try:
            prs = slide.part.package.presentation_part.presentation
            page_height = prs.slide_height
        except AttributeError:
            # Fallback: Assume standard 16:9 (10 inches height? No, 7.5 inches usually)
            # 7.5 inches = 6858000 EMUs
            page_height = 6858000 

        
        # Determine layout constraints
        start_top_num = proto_num.top
        start_top_title = proto_title.top
        
        # Use the lower starting point as the reference for "start_y"
        start_y = max(start_top_num, start_top_title)
        item_height = max(proto_num.height, proto_title.height)
        
        # Define bottom margin (use same as top margin for symmetry, or at least 1 inch)
        # Assuming 96 dpi, 1 inch = 914400 EMUs. 
        # Let's use a safe bottom margin.
        bottom_margin = start_y 
        max_y = page_height - bottom_margin
        
        num_items = len(data.slides)
        
        # Default spacing
        step_y = item_height * 1.5
        
        if num_items > 1:
            # Calculate available vertical span for the *starts* of the items
            # The last item starts at `max_y - item_height`
            available_span = max_y - start_y - item_height
            
            # Calculate required spacing to fit exactly
            calculated_step = available_span / (num_items - 1)
            
            # Use the smaller of (calculated_step, default_spacing) to avoid spreading too much
            # But if calculated_step is very small (negative even), we must compress.
            # So we actually want:
            # If calculated_step < default_spacing: use calculated_step (compress to fit)
            # If calculated_step > default_spacing: use default_spacing (don't spread too much)
            
            # However, we shouldn't overlap too much. 
            # Minimum spacing = item_height (touching)
            
            step_y = min(calculated_step, item_height * 1.5)
            
            # Ensure we don't overlap if possible (unless page is too small)
            if step_y < item_height:
                # Warning: Items will overlap. 
                # We could enforce step_y = item_height, but then it overflows page.
                # User asked to "limit them in the page", so we respect page bounds even if it overlaps.
                pass

        # --- Fix for Issue 2: Clone FIRST, then fill ---
        # We collect all shapes (original + clones) first, WITHOUT modifying them yet.
        # This ensures all clones are based on the clean prototype.
        
        toc_items = [] # List of (num_shape, title_shape)
        
        for i in range(len(data.slides)):
            if i == 0:
                # Use the prototype itself for the first item
                toc_items.append((proto_num, proto_title))
            else:
                # Clone the prototype (which is still clean because we haven't modified it yet)
                new_num = duplicate_shape(proto_num, slide)
                new_title = duplicate_shape(proto_title, slide)
                
                # Position
                offset = i * step_y
                new_num.top = start_top_num + int(offset)
                new_title.top = start_top_title + int(offset)
                
                toc_items.append((new_num, new_title))
        
        # Now fill text for all items
        for i, (num_shape, title_shape) in enumerate(toc_items):
            num_text = f"{i+1:02d}"
            # Clean title for TOC (remove "一、", "1." etc)
            title_text = self._clean_title(data.slides[i].title)
            
            self._set_text(num_shape, num_text)
            self._set_text(title_shape, title_text)

    def _fill_content_page(self, slide, chapter_data, all_titles, current_index):
        """
        填充内容页
        包含: 导航栏, 描述, 正文
        """
        shape_map = self._build_shape_map(slide)
        
        # 1. 生成导航栏 (page1_title)
        proto_nav = shape_map.get("page1_title")
        if proto_nav:
            margin_x = proto_nav.width * 1.1 
            start_left = proto_nav.left
            
            # --- Fix for Issue 4: Clone FIRST, then fill ---
            nav_items = []
            for i in range(len(all_titles)):
                if i == 0:
                    nav_items.append(proto_nav)
                else:
                    new_shape = duplicate_shape(proto_nav, slide)
                    new_shape.left = start_left + int(i * margin_x)
                    nav_items.append(new_shape)
            
            # Fill text and color
            for i, shape in enumerate(nav_items):
                # Clean title for Nav Bar
                clean_title = self._clean_title(all_titles[i])
                self._set_text(shape, clean_title)
                
                if i != current_index:
                    self._set_font_color(shape, RGBColor(192, 192, 192)) # Gray
                else:
                    self._set_font_color(shape, RGBColor(0, 0, 0)) # Black
        else:
            print("Warning: Nav prototype (page1_title) not found on content slide.")

        # 2. 填充描述 (page1_desc)
        if "page1_desc" in shape_map:
            self._set_text(shape_map["page1_desc"], chapter_data.description)

        # 3. 填充正文 (page1_bullet1)
        if "page1_bullet1" in shape_map:
            full_blocks = []
            for block in chapter_data.blocks:
                full_blocks.append(block)
            self._set_blocks_text(shape_map["page1_bullet1"], full_blocks)

    def _set_blocks_text(self, shape, blocks):
        """
        将多个 ContentBlock 填充到一个文本框
        """
        text_frame = shape.text_frame
        
        # Capture style from first run of first paragraph
        template_run = None
        if text_frame.paragraphs and text_frame.paragraphs[0].runs:
            template_run = text_frame.paragraphs[0].runs[0]
            
        text_frame.clear() # Clear all content
        
        for i, block in enumerate(blocks):
            # Subtitle
            if block.subtitle:
                p = text_frame.add_paragraph() if i > 0 or text_frame.paragraphs else text_frame.paragraphs[0]
                run = p.add_run()
                run.text = block.subtitle
                if template_run:
                    self._copy_font_style(template_run, run)
                run.font.bold = True
            
            # Bullets
            for bullet in block.bullets:
                p = text_frame.add_paragraph()
                p.level = 1
                run = p.add_run()
                run.text = bullet
                if template_run:
                    self._copy_font_style(template_run, run)

    def _set_text(self, shape, text):
        """
        设置文本，保留原有格式。
        Fix: 确保清除多余的段落，防止旧文本残留。
        """
        text_frame = shape.text_frame
        
        if not text_frame.paragraphs:
            text_frame.add_paragraph()
            
        p = text_frame.paragraphs[0]
        
        # Set text on first run
        if p.runs:
            p.runs[0].text = text
            # Clear subsequent runs in the first paragraph
            for i in range(1, len(p.runs)):
                p.runs[i].text = ""
        else:
            run = p.add_run()
            run.text = text

        # --- Fix for Issue 2: Remove extra paragraphs ---
        # If the original shape had multiple paragraphs, remove them.
        # We iterate backwards to avoid index issues.
        for i in range(len(text_frame.paragraphs) - 1, 0, -1):
            p_element = text_frame.paragraphs[i]._p
            p_element.getparent().remove(p_element)

    def _set_font_color(self, shape, rgb_color):
        """
        强制设置文本框内所有文字的颜色
        """
        if not shape.has_text_frame:
            return
        for p in shape.text_frame.paragraphs:
            for run in p.runs:
                run.font.color.rgb = rgb_color

    def _copy_font_style(self, src_run, dest_run):
        """
        将 src_run 的字体样式复制到 dest_run
        """
        if not src_run or not dest_run:
            return
        
        try:
            if src_run.font.name:
                dest_run.font.name = src_run.font.name
                self._set_ea_font(dest_run, src_run.font.name)
            
            if src_run.font.size:
                dest_run.font.size = src_run.font.size
            
            if src_run.font.bold is not None:
                dest_run.font.bold = src_run.font.bold
            if src_run.font.italic is not None:
                dest_run.font.italic = src_run.font.italic
                
            if src_run.font.color:
                if src_run.font.color.type == 1: # RGB
                    dest_run.font.color.rgb = src_run.font.color.rgb
                elif src_run.font.color.type == 2: # Theme Color
                    dest_run.font.color.theme_color = src_run.font.color.theme_color
                
        except Exception as e:
            print(f"Warning: Failed to copy font style: {e}")

    def _set_ea_font(self, run, font_name):
        """
        设置东亚字体
        """
        from pptx.oxml.ns import qn
        try:
            # Try standard access
            rPr = run._element.get_or_add_rPr()
        except AttributeError:
            # Fallback for _Run objects where _element might be named _r
            if hasattr(run, '_r'):
                rPr = run._r.get_or_add_rPr()
            else:
                return

        ea = rPr.find(qn('a:ea'))
        if ea is None:
            ea = rPr.makeelement(qn('a:ea'))
            rPr.append(ea)
        ea.set('typeface', font_name)

    def _clean_title(self, title):
        """
        Remove leading numbering like "一、", "1.", "1、"
        """
        if not title:
            return ""
        # Pattern: Start of string, followed by (digits or chinese numerals), followed by dot or comma or chinese comma, optional whitespace
        pattern = r"^([0-9]+|[一二三四五六七八九十百]+)[、.．]\s*"
        return re.sub(pattern, "", title)
