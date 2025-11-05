# generate_readme_pdf.py
# Generates README_Data_Dashboard.pdf using ReportLab

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

readme_text = '''
# Data Dashboard and Report

Project Overview:
This project analyzes two months of simulated sales data and creates an Excel-based dashboard (pivot tables and charts).

Files generated:
- Sales_Dashboard.xlsx         (raw simulated sales data)
- Sales_Dashboard_with_Charts.xlsx  (includes a Pivot_Charts sheet with visuals)

Technologies:
- Python (pandas, numpy, openpyxl, reportlab)
- Microsoft Excel

How to run:
1. Run generate_sales_data.py to create Sales_Dashboard.xlsx
2. Run add_pivot_charts.py to create Sales_Dashboard_with_Charts.xlsx with charts
3. Optionally run this script to recreate the README PDF

Insights (sample):
- Top products: Smartphones & Laptops
- Top regions: South & West
- Profit margins are stable; consider marketing focus on high-converting regions.
'''

doc = SimpleDocTemplate('README_Data_Dashboard.pdf', pagesize=A4)
styles = getSampleStyleSheet()
story = []
for para in readme_text.strip().split('\n\n'):
    story.append(Paragraph(para.replace('\n', '<br/>'), styles['Normal']))
    story.append(Spacer(1,8))
doc.build(story)
print('Saved README_Data_Dashboard.pdf')
