def add_header_footer_number(slide, slide_num_str):
    # PROSERVE header
    header = slide.shapes.add_textbox(Inches(8), Inches(0), Inches(2), Inches(0.5))
    header_tf = header.text_frame
    header_tf.text = "PROSERVE"
    header_tf.paragraphs[0].font.size = Pt(32)
    header_tf.paragraphs[0].font.bold = True
    header_tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    header_tf.paragraphs[0].alignment = PP_ALIGN.RIGHT

    # Footer
    footer = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(0.5))
    footer_tf = footer.text_frame
    footer_tf.text = "Zscaler, Inc. All rights reserved. © 2025"
    footer_tf.paragraphs[0].font.size = Pt(8)
    footer_tf.paragraphs[0].font.color.rgb = RGBColor(128, 128, 128)
    footer_tf.paragraphs[0].alignment = PP_ALIGN.LEFT

    # Slide number
    slide_num = slide.shapes.add_textbox(Inches(9.5), Inches(6.5), Inches(0.5), Inches(0.5))
    slide_num_tf = slide_num.text_frame
    slide_num_tf.text = slide_num_str
    slide_num_tf.paragraphs[0].font.size = Pt(12)
    slide_num_tf.paragraphs[0].font.color.rgb = RGBColor(128, 128, 128)
    slide_num_tf.paragraphs[0].alignment = PP_ALIGN.RIGHT
