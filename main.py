import pandas as pd
import glob
import openpyxl as pyxl
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_num = filename.split("-")
    pdf.set_font(family="Times", style="B", size=24)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=0, h=12, txt=f"Invoice {invoice_num[0]}", align="L", ln=1)
    pdf.output(f"PDF_INVOICES/Invoice{invoice_num[0]}.pdf")
