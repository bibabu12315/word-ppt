# -*- coding: utf-8 -*-
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

def create_demo_template(output_path: str):
    """
    创建一个符合项目命名约定的测试用 PPT 模板。
    包含：
    - 封面页 (cover_title, presenter, date...)
    - 10 页正文页 (page1...page10)
      - 每页包含 pageX_title
      - 每页包含 pageX_bullet1, pageX_bullet2 (左右布局)
    """
    prs = Presentation()
    
    # 1. 创建封面页 (Slide 0)
    # 使用空白版式 (6 = Blank)
    slide_layout = prs.slide_layouts[6] 
    slide = prs.slides.add_slide(slide_layout)
    
    # 添加封面标题框
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
    title_box.name = "cover_title"
    title_box.text = "封面标题占位符"
    p = title_box.text_frame.paragraphs[0]
    p.font.size = Pt(44)
    p.alignment = PP_ALIGN.CENTER

    # 添加元数据框
    meta_box = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(8), Inches(1))
    meta_box.name = "presenter"
    meta_box.text = "汇报人占位符"
    p = meta_box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    
    # 2. 创建正文页 (Page 1 - 10)
    for i in range(1, 11):
        slide = prs.slides.add_slide(slide_layout)
        
        # 页面标题
        # 命名规则: page{i}_title
        title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
        title_shape.name = f"page{i}_title"
        title_shape.text = f"Page {i} Title"
        title_shape.text_frame.paragraphs[0].font.size = Pt(32)
        title_shape.text_frame.paragraphs[0].font.bold = True
        
        # 内容框 1 (左侧)
        # 命名规则: page{i}_bullet1
        body1 = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(4.2), Inches(5))
        body1.name = f"page{i}_bullet1"
        body1.text = f"Page {i} Bullet 1 Area"
        
        # 内容框 2 (右侧)
        # 命名规则: page{i}_bullet2
        body2 = slide.shapes.add_textbox(Inches(5.0), Inches(1.8), Inches(4.2), Inches(5))
        body2.name = f"page{i}_bullet2"
        body2.text = f"Page {i} Bullet 2 Area"

        # 可以在这里添加更多框，如 page{i}_bullet3...

    # 确保目录存在
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    print(f"Template created at: {output_path}")

if __name__ == "__main__":
    create_demo_template("d:\\project_code\\Word转PPT\\input\\template.pptx")
