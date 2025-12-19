import copy
import uuid
from pptx.shapes.autoshape import Shape
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.picture import Picture
from pptx.shapes.group import GroupShape
from pptx.shapes.connector import Connector

def duplicate_slide(pres, index):
    """
    Duplicate the slide at the given index in the presentation.
    Adds the new slide to the end of the presentation.
    Returns the new slide.
    """
    source_slide = pres.slides[index]
    blank_slide_layout = pres.slide_layouts[6] # Use blank layout as base
    dest_slide = pres.slides.add_slide(blank_slide_layout)

    # Copy all shapes from source to dest
    for shape in source_slide.shapes:
        new_shape = duplicate_shape(shape, dest_slide, keep_name=True)
    
    return dest_slide

def duplicate_shape(shape, slide, keep_name=False):
    """
    Duplicate a shape onto a specific slide.
    Preserves style (fill, line, effects) by copying spPr.
    """
    new_shape = None
    
    # 1. Text Box / AutoShape
    if isinstance(shape, Shape):
        try:
            autoshape_type_id = shape.auto_shape_type
            new_shape = slide.shapes.add_shape(
                autoshape_type_id,
                shape.left, shape.top, shape.width, shape.height
            )
        except ValueError:
            # Likely a TextBox
            new_shape = slide.shapes.add_textbox(
                shape.left, shape.top, shape.width, shape.height
            )

        # --- Critical Fix for Issue 3: Copy Shape Properties (Fill, Line, Gradient, etc.) ---
        # We overwrite the new shape's spPr with a deep copy of the original
        # This preserves gradient fills, borders, shadows, etc.
        # Fix: Use element manipulation instead of property assignment
        new_element = new_shape.element
        new_spPr = copy.deepcopy(shape.element.spPr)
        
        if new_element.spPr is not None:
            # Replace existing spPr
            # Find index to maintain order
            idx = new_element.index(new_element.spPr)
            new_element.remove(new_element.spPr)
            new_element.insert(idx, new_spPr)
        else:
            # Insert after nvSpPr (which is always first)
            new_element.insert(1, new_spPr)
        
        # Re-apply position (since spPr might contain old position)
        # Actually, spPr contains <a:xfrm>, so copying it copies the position too.
        # But we passed left/top to add_shape, which set the initial xfrm.
        # Overwriting spPr overwrites that.
        # So if we want to move it later, we can. 
        # But duplicate_shape is supposed to create an exact clone at the same position initially.
        # So copying spPr is perfect.

        # Copy text content
        if shape.has_text_frame:
            # We need to be careful. Copying spPr does NOT copy text.
            # Text is in txBody.
            # We can copy txBody too!
            # Fix: Use element manipulation instead of property assignment
            new_txBody = copy.deepcopy(shape.element.txBody)
            
            if new_element.txBody is not None:
                idx = new_element.index(new_element.txBody)
                new_element.remove(new_element.txBody)
                new_element.insert(idx, new_txBody)
            else:
                new_element.append(new_txBody)
            
            # If we copy txBody, we don't need to manually copy paragraphs/runs.
            # This is much better and preserves all text formatting perfectly.
            pass
        
    # 2. Group (Recursive)
    elif isinstance(shape, GroupShape):
        pass
        
    # 3. Picture
    elif isinstance(shape, Picture):
        with open("temp_img.png", "wb") as f:
            f.write(shape.image.blob)
        new_shape = slide.shapes.add_picture(
            "temp_img.png", shape.left, shape.top, shape.width, shape.height
        )
        
    # 4. Copy Name (Fix for Issue 1)
    if new_shape:
        if keep_name:
            new_shape.name = shape.name
        else:
            # Append a unique suffix to avoid name collision confusion
            new_shape.name = f"{shape.name}_copy_{uuid.uuid4().hex[:8]}"
        
    return new_shape

def _copy_font(src_font, dest_font):
    """Helper to copy font attributes (Legacy, mostly unused if we copy txBody)"""
    if src_font.name:
        dest_font.name = src_font.name
    if src_font.size:
        dest_font.size = src_font.size
    if src_font.bold is not None:
        dest_font.bold = src_font.bold
    if src_font.italic is not None:
        dest_font.italic = src_font.italic
    if src_font.color and src_font.color.type == 1: # RGB
        dest_font.color.rgb = src_font.color.rgb

def move_slide(pres, old_index, new_index):
    """
    Move a slide from old_index to new_index.
    """
    xml_slides = pres.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[old_index])
    xml_slides.insert(new_index, slides[old_index])
