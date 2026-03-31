"""Text formatting: rich text with multiple paragraphs, fonts, colors, and styles."""

from slidecraft import Inches, PP_ALIGN, Presentation, Pt, RGBColor

prs = Presentation()
slide = prs.slides.add()

tb = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(6))
tf = tb.text_frame
tf.word_wrap = True

# Title paragraph
p1 = tf.paragraphs[0]
p1.alignment = PP_ALIGN.CENTER
r = p1.add_run()
r.text = "Rich Text Demo"
r.font.size = Pt(32)
r.font.bold = True
r.font.color = RGBColor(0, 51, 102)

# Mixed formatting in one paragraph
p2 = tf.add_paragraph()
p2.space_before = Pt(18)
r1 = p2.add_run()
r1.text = "Bold, "
r1.font.bold = True
r1.font.size = Pt(18)
r2 = p2.add_run()
r2.text = "italic, "
r2.font.italic = True
r2.font.size = Pt(18)
r3 = p2.add_run()
r3.text = "underlined, "
r3.font.underline = True
r3.font.size = Pt(18)
r4 = p2.add_run()
r4.text = "and colored."
r4.font.size = Pt(18)
r4.font.color = RGBColor(200, 50, 50)

# Left-aligned paragraph
p3 = tf.add_paragraph()
p3.alignment = PP_ALIGN.LEFT
p3.space_before = Pt(12)
r = p3.add_run()
r.text = "Left-aligned text in Arial"
r.font.name = "Arial"
r.font.size = Pt(16)

# Right-aligned paragraph
p4 = tf.add_paragraph()
p4.alignment = PP_ALIGN.RIGHT
r = p4.add_run()
r.text = "Right-aligned text in Courier New"
r.font.name = "Courier New"
r.font.size = Pt(16)
r.font.color = RGBColor(0, 128, 0)

# Justified paragraph
p5 = tf.add_paragraph()
p5.alignment = PP_ALIGN.JUSTIFY
p5.space_before = Pt(12)
r = p5.add_run()
r.text = "This paragraph is justified. It demonstrates how longer text wraps and aligns evenly on both sides."
r.font.size = Pt(14)

prs.save("text_formatting.pptx")
print("Saved text_formatting.pptx")
