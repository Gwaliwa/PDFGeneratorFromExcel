import os
import glob
import pandas as pd
from fpdf import FPDF
from pathlib import Path

filepath = glob.glob("invoices/*.xlsx")

# create a pdf files
for file in filepath:
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    fname = Path(file).stem
    number = fname.split("-")[0]
    date=fname.split("-")[1]
    pdf.set_font("Arial", size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice no:{number}", ln=1, align="L")
    pdf.set_font("Arial", size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"Date:{date}", ln=1, align="L")

    #create a table
    df = pd.read_excel(file, sheet_name="Sheet 1")

    #create header
    columns = df.columns.tolist()
    columns = [item.replace("_", " ").capitalize() for item in columns]
    pdf.set_font("Arial", size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"{columns[0]}", align="L", border=1)
    pdf.cell(w=70, h=8, txt=f"{columns[1]}", align="L", border=1)
    pdf.cell(w=50, h=8, txt=f"{columns[2]}", align="L", border=1)
    pdf.cell(w=50, h=8, txt=f"{columns[3]}", align="L", border=1)
    pdf.cell(w=50, h=8, txt=f"{columns[4]}", align="L", border=1, ln=1)

    #contents in the table
    for index, row in df.iterrows():
        pdf.set_font("Arial", size=12, style="B")
        pdf.cell(w=50, h=8, txt=f"{row["product_id"]}", align="L", border=1)
        pdf.cell(w=70, h=8, txt=f"{row["product_name"]}", align="L", border=1)
        pdf.cell(w=50, h=8, txt=f"{row["amount_purchased"]}", align="L", border=1)
        pdf.cell(w=50, h=8, txt=f"{row["price_per_unit"]}", align="L", border=1)
        pdf.cell(w=50, h=8, txt=f"{row["total_price"]}", ln=1, align="L", border=1)


    pdf.output(f"PDF/{fname}.pdf")
    print(df)
