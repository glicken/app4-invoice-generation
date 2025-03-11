
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# ensure PDFs folder exists
output_dir = Path("PDFs")
output_dir.mkdir(parents=True, exist_ok=True)

filepaths = glob.glob("invoices/*.xlsx")   

for filepath in filepaths:
    # Create PDF file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # import filenames:
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
   

    # Create PDFs folder and enclosed PDFs
    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=6)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # add rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)
        
   
    # calculate total price
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    pdf.set_font(family="Times", size=10, style='B')
    pdf.cell(w=30, h=8, txt=f"The total price is ${total_sum}", border=0, ln=1)
    pdf.cell(w=20, h=8, txt="Pythonhow", border=0)
    pdf.image("pythonhow.png", h=10, w=10)

    # Save the PDF file inside the PDFs folder
    pdf.output(output_dir / f"{filename}.pdf")
    
