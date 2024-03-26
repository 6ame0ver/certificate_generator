import pandas as pd
import os
import streamlit as st
from fpdf import FPDF

xls = pd.ExcelFile('product_data.xlsx')
df = pd.read_excel(xls)
product = df.to_dict()

# Set webpage layout to wide
st.set_page_config(layout="wide")

# Add a header
st.title("PDF Certificate Generator")

# Order number
col, buff = st.columns([1, 1])
order_number = col.text_input("Order Number: ")

# Line number
line_number = col.number_input("Line number: ", min_value=1, step=1)

# Product ID
product_id = col.selectbox("Product Number:", product.keys())
new_product_id = str(product_id)

# Item description
description = str(product[product_id])
new_description = description[5:-2]
col.text_input("Product Description:", new_description)

# Revision no
revision_number = col.number_input("Revision no: ", min_value=1, step=1)

# Quantity
quantity = col.number_input("Quantity: ", min_value=1, step=1)

# Country of origin
country_of_origin = col.text_input("Country of origin: ")

# Date of manufacturing
date_of_manufacturing = col.text_input("Date of manufacturing: ")

# Serial numbers
serials = st.text_area("Serial Numbers: ")

# Name
name = col.text_input("Your name: ")

# Date
date = str(col.date_input("Date: "))

# Generate button. This button should pass all data to the main file to generate PDF file and ask where to save it.
create_pdf = st.button("Generate PDF")
if create_pdf:
    class FPDF(FPDF):
        def header(self):
            pass

        def footer(self):
            self.set_y(-15)
            self.set_font(family="Arial", size=11)
            self.set_text_color(180, 180, 180)
            self.cell(w=0, h=10, txt="QMF 15 03 16.08.2022", align="L", ln=True)


    # df = pd.read_excel("Cert_data.xlsx")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Document title.
    title = "Certificate of Conformity"
    pdf.set_font(family="Arial", size=14, style="BU")
    pdf.cell(w=200, h=10, txt=title, align="C", ln=True)

    # Globe address.
    address = """ 
    Globe Microsystems Limited
    Unit E
    Argent Court
    Hook Rise South
    Surbiton
    KT6 7NL"""
    pdf.set_font(family="Arial", size=11)
    pdf.multi_cell(w=100, h=4, txt=address)
    pdf.ln(6)

    # Order and line number.
    # order_number
    # line_number
    pdf.set_font(family="Arial", size=11)
    pdf.cell(w=188, h=8, txt=f"Your Order Number: {order_number}, Line Number: {line_number}", ln=True)

    # Tabel for product and description. Possibility to type product or choose from the list.
    # Add a header
    pdf.set_font(family="Arial", size=11, style="B")
    pdf.cell(w=30, h=8, txt="Item Number", border=1)
    pdf.cell(w=18, h=8, txt="Rev. No", border=1)
    pdf.cell(w=120, h=8, txt="Item Name", border=1)
    pdf.cell(w=20, h=8, txt="Qty", border=1, ln=True)

    # Add row to table. Revision, Description should come out automatically. Manually enter Qty.
    pdf.set_font(family="Arial", size=11)
    pdf.cell(w=30, h=8, txt=f"{new_product_id}", border=1)
    pdf.cell(w=18, h=8, txt=f"{revision_number}", border=1)
    pdf.cell(w=120, h=8, txt=f"{new_description}", border=1)
    pdf.cell(w=20, h=8, txt=f"{quantity}", border=1, ln=True)

    # Country of origin.
    pdf.set_font(family="Arial", size=11, style="B")
    pdf.cell(w=30, h=8, txt=f"Country of origin: {country_of_origin}", ln=True)

    # Date of manufacturing.
    pdf.set_font(family="Arial", size=11, style="B")
    pdf.cell(w=30, h=8, txt=f"Date of manufacture: {date_of_manufacturing}", ln=True)

    # Serial numbers.
    pdf.set_font(family="Arial", size=11, style="B")
    pdf.multi_cell(w=188, h=6, txt=f"Serial Numbers: {serials}")
    pdf.ln(5)

    # Create Deviations table.
    # Add a header
    pdf.set_font(family="Arial", size=11, style="B")
    pdf.cell(w=188, h=8, txt="Deviations Permits/ Concessions (if any)", border=1, ln=True)
    # Add row
    pdf.set_font(family="Arial", size=11)
    pdf.cell(w=188, h=20, txt="", border=1, ln=True)

    # Add text under the table:
    text_1 = "If is certified that apart from the approved deviation permit/concessions noted above if any,"
    text_2 = "the products listed above conform in all respects to the contract requirements."
    pdf.set_font(family="Arial", size=11)
    pdf.cell(w=188, h=8, txt=text_1)
    pdf.ln(5)
    pdf.set_font(family="Arial", size=11)
    pdf.cell(w=188, h=8, txt=text_2)
    pdf.ln(12)

    # Authorised Name.
    pdf.set_font(family="Arial", size=11)
    pdf.cell(w=30, h=8, txt=f"Authorised Name: {name}", ln=True)

    # Authorised Signature,
    pdf.set_font(family="Arial", size=11)
    pdf.cell(w=30, h=8, txt="Authorised Signature: ___________________", ln=True)

    # Date.
    year = date[:4]
    month = date[5:7]
    day = date[8:]
    doc_date = f"Date: {day} / {month} / {year}"
    pdf.set_font(family="Arial", size=11)
    pdf.cell(w=30, h=8, txt=f"{doc_date}", ln=True)

    # Produce the PDF document.
    pdf.output("output.pdf")

    # After pdf file is created need to be open for signature and save.
    os.startfile("output.pdf")