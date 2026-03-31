"""Basic presentation: create a slide with a styled text box and save."""

from slidecraft import Inches, PP_ALIGN, Presentation, Pt, RGBColor

prs = Presentation()
slide = prs.slides.add()

# Add a text box
tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
tf = tb.text_frame
tf.word_wrap = True

# Style the first paragraph
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
run = p.runs[0] if p.runs else p.add_run()
run.text = "Hello, SlideCraft!"
run.font.bold = True
run.font.size = Pt(36)
run.font.color = RGBColor(0, 51, 102)

# Add a subtitle paragraph
p2 = tf.add_paragraph()
p2.alignment = PP_ALIGN.CENTER
run2 = p2.add_run()
run2.text = "A modern Python PowerPoint library"
run2.font.size = Pt(18)
run2.font.italic = True
run2.font.color = RGBColor(100, 100, 100)

prs.save("basic_presentation.pptx")
print("Saved basic_presentation.pptx")
