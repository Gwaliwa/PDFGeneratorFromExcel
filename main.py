import os
import glob
import pandas as pd
from fpdf import FPDF
from pathlib import Path

filepath = glob.glob("invoices/*.xlsx")
for file in filepath:
    df = pd.read_excel(file, sheet_name="Sheet 1")
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    fname = Path(file).stem
    number = fname.split("-")[0]
    pdf.add_page()
    pdf.set_font("Arial", size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"invoice no:{number}")
    pdf.output(f"PDF/{fname}.pdf")
    print(df)
