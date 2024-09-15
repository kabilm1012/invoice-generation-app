import pandas as pd
import glob
from fpdf import FPDF
import pathlib

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    filepath = pathlib.Path(filepath)
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = filepath.stem
    lst = filename.split("-")
    invoice_nr = lst[0]
    date = lst[1]
    pdf.set_font(family="Times", style="B", size=18)
    pdf.cell(w=0, h=8, txt=f"Invoice.nr.{invoice_nr}", align="L", ln=1)
    pdf.cell(w=0, h=8, txt=f"Date: {date}", align="L", ln=1)
    pdf.output(f"PDFs/{filename}.pdf")

