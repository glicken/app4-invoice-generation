import pandas as pd
from fpdf import FPDF

pdf = FPDF(orientation="L", unit="mm", format="A4")
pdf.add_page()

pdf.set_font("Arial", size=16, style='B')
pdf.cell(w=50, h=8, txt="2025 Centuries", ln=1)

df = pd.read_excel("2025_Centuries.xlsx", sheet_name="Sheet1")


# Add header

columns = df.columns
pdf.set_font(family="Times", size=10, style='B')
pdf.set_text_color(80, 80, 80)
pdf.cell(w=25, h=8, txt=columns[0], border=1)
pdf.cell(w=50, h=8, txt=columns[1], border=1)
pdf.cell(w=40, h=8, txt=columns[2], border=1)
pdf.cell(w=30, h=8, txt=columns[3], border=1)
pdf.cell(w=20, h=8, txt=columns[4], border=1)
pdf.cell(w=30, h=8, txt=columns[5], border=1)
pdf.cell(w=30, h=8, txt=columns[6], border=1, ln=1)

for index, row in df.iterrows():
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(50, 50, 50)
    pdf.cell(w=25, h=8, txt=str(row[0],), border=1)
    pdf.cell(w=50, h=8, txt=str(row[1],), border=1)
    pdf.cell(w=40, h=8, txt=str(row[2],), border=1)
    pdf.cell(w=30, h=8, txt=str(row[3],), border=1)
    pdf.cell(w=20, h=8, txt=str(row[4],), border=1)
    pdf.cell(w=30, h=8, txt=str(row[5],), border=1)
    pdf.cell(w=30, h=8, txt=str(row[6],), border=1, ln=1)

output = "2025_Centuries.pdf"
pdf.output(output)
print(f"PDF file saved as {output}")
