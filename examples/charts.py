"""Charts: bar, column, line, and pie charts with multi-series data."""

from slidecraft import ChartData, ChartType, Inches, Presentation

prs = Presentation()

# --- Column chart (multi-series) ---
slide1 = prs.slides.add()
cd = ChartData()
cd.categories = ["Q1", "Q2", "Q3", "Q4"]
cd.add_series("Revenue", [120, 200, 180, 300])
cd.add_series("Costs", [80, 150, 110, 200])
slide1.shapes.add_chart(ChartType.COLUMN, cd, Inches(1), Inches(1), Inches(8), Inches(5.5))

# --- Bar chart ---
slide2 = prs.slides.add()
cd2 = ChartData()
cd2.categories = ["Engineering", "Sales", "Marketing"]
cd2.add_series("Headcount", [45, 30, 20])
cd2.add_series("Open Roles", [5, 8, 3])
slide2.shapes.add_chart(ChartType.BAR, cd2, Inches(1), Inches(1), Inches(8), Inches(5.5))

# --- Line chart ---
slide3 = prs.slides.add()
cd3 = ChartData()
cd3.categories = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
cd3.add_series("Users", [1000, 1500, 1800, 2400, 3100, 4000])
cd3.add_series("Active", [800, 1200, 1400, 2000, 2600, 3400])
slide3.shapes.add_chart(ChartType.LINE, cd3, Inches(1), Inches(1), Inches(8), Inches(5.5))

# --- Pie chart (single series) ---
slide4 = prs.slides.add()
cd4 = ChartData()
cd4.categories = ["Desktop", "Mobile", "Tablet"]
cd4.add_series("Traffic", [55, 35, 10])
slide4.shapes.add_chart(ChartType.PIE, cd4, Inches(2), Inches(1), Inches(6), Inches(5.5))

prs.save("charts.pptx")
print("Saved charts.pptx")
