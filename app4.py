import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    # import filenames:
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    # Create PDFs folder abnd enclosed PDFs
    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}")
    pdf.output(f"PDFs/{filename}.pdf")
    
