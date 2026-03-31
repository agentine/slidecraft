"""Multi-slide presentation: combines text, shapes, tables, and charts."""

from slidecraft import ChartData, ChartType, Inches, PP_ALIGN, Presentation, Pt, RGBColor

prs = Presentation()

# --- Slide 1: Title ---
s1 = prs.slides.add()
tb = s1.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(2))
tf = tb.text_frame
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "Quarterly Report"
r.font.size = Pt(40)
r.font.bold = True
r.font.color = RGBColor(0, 51, 102)
p2 = tf.add_paragraph()
p2.alignment = PP_ALIGN.CENTER
r2 = p2.add_run()
r2.text = "FY 2026 - Q1"
r2.font.size = Pt(20)
r2.font.color = RGBColor(100, 100, 100)

# --- Slide 2: Key Metrics (shapes + text) ---
s2 = prs.slides.add()
for i, (label, value) in enumerate([("Revenue", "$1.2M"), ("Users", "4,200"), ("NPS", "72")]):
    x = Inches(0.5 + i * 3.2)
    box = s2.shapes.add_shape(2, x, Inches(2), Inches(2.8), Inches(3))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = value
    r.font.size = Pt(36)
    r.font.bold = True
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = label
    r2.font.size = Pt(16)

# --- Slide 3: Chart ---
s3 = prs.slides.add()
cd = ChartData()
cd.categories = ["Jan", "Feb", "Mar"]
cd.add_series("Revenue", [380, 420, 400])
cd.add_series("Expenses", [300, 310, 330])
s3.shapes.add_chart(ChartType.COLUMN, cd, Inches(1), Inches(1), Inches(8), Inches(5.5))

# --- Slide 4: Summary table ---
s4 = prs.slides.add()
table = s4.shapes.add_table(4, 2, Inches(2), Inches(1.5), Inches(6), Inches(3))
for col, hdr in enumerate(["Metric", "Value"]):
    c = table.cell(0, col)
    c.text = hdr
    table.cell(0, col).text_frame.paragraphs[0].runs[0].font.bold = True
for row, (k, v) in enumerate([("Revenue", "$1.2M"), ("Growth", "+15%"), ("Target", "$1.5M")], 1):
    table.cell(row, 0).text = k
    table.cell(row, 1).text = v

prs.save("multi_slide.pptx")
print("Saved multi_slide.pptx")
