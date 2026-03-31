"""Tables: create a table with headers, fill cells, and style rows."""

from slidecraft import Inches, PP_ALIGN, Presentation, Pt, RGBColor

prs = Presentation()
slide = prs.slides.add()

rows, cols = 5, 3
table = slide.shapes.add_table(rows, cols, Inches(1.5), Inches(1.5), Inches(7), Inches(4))

# Set column widths
table.columns[0].width = Inches(2.5)
table.columns[1].width = Inches(2.5)
table.columns[2].width = Inches(2)

# Header row
headers = ["Product", "Category", "Price"]
for col_idx, text in enumerate(headers):
    cell = table.cell(0, col_idx)
    cell.text = text
    run = cell.text_frame.paragraphs[0].runs[0]
    run.font.bold = True
    run.font.size = Pt(14)
    run.font.color = RGBColor(255, 255, 255)

# Data rows
data = [
    ("Widget A", "Hardware", "$29.99"),
    ("Widget B", "Software", "$49.99"),
    ("Widget C", "Hardware", "$19.99"),
    ("Widget D", "Services", "$99.99"),
]
for row_idx, (product, category, price) in enumerate(data, start=1):
    table.cell(row_idx, 0).text = product
    table.cell(row_idx, 1).text = category
    cell = table.cell(row_idx, 2)
    cell.text = price
    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

prs.save("tables.pptx")
print("Saved tables.pptx")
