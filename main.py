import pandas as pd
import glob
from fpdf import FPDF
import pathlib

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    filepath = pathlib.Path(filepath)

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = filepath.stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", style="B", size=18)
    pdf.cell(w=0, h=8, txt=f"Invoice.nr.{invoice_nr}", align="L", ln=1)
    pdf.cell(w=0, h=8, txt=f"Date: {date}", align="L", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add table headers
    columns = df.columns
    columns = [item.replace('_', ' ').title() for item in columns]
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=f"{row['product_id']}", border=1)
        pdf.cell(w=70, h=8, txt=f"{row['product_name']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['amount_purchased']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['price_per_unit']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['total_price']}", border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")


