import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
import os

# Function to fill the PDF form
def fill_pdf(data, template_path):
    reader = PdfReader(template_path)
    writer = PdfWriter()

    # Create a canvas overlay to add text to the PDF
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    # Fill the form fields with data
    can.drawString(125, 680, data['Name'])
    can.drawString(400, 680, data['Surname'])
    can.drawString(136, 655, data['DOB'])
    can.drawString(125, 675, data['Address'])
    #can.drawString(125, 650, data['Postcode'])
    can.drawString(136, 577, data['Telephone'])
    can.drawString(136, 552, data['Email'])
    can.drawString(155, 515, data['Job 1 Title'])
    can.drawString(155, 500, data['Job 1 Company'])
    can.drawString(155, 485, data['Job 1 Start'])
    can.drawString(300, 485, data['Job 1 End'])
    can.drawString(155, 465, data['Job 2 Title'])
    can.drawString(155, 447, data['Job 2 Company'])
    can.drawString(155, 430, data['Job 2 Start'])
    can.drawString(300, 430, data['Job 2 End'])
    can.drawString(225, 333, data['Course 1 Title'])
    can.drawString(225, 320, data['College 1 Name'])
    can.drawString(225, 305, data['Year 1 Completion'])
    can.drawString(225, 292, data['Level 1 Qualification'])
    can.drawString(225, 270, data['Course 2 Title'])
    can.drawString(225, 255, data['College 2 Name'])
    can.drawString(225, 240, data['Year 2 Completion'])
    can.drawString(225, 230, data['Level 2 Qualification'])
    can.drawString(125, 225, data['Hobbies'])
    can.drawString(125, 200, data['Languages'])
    #can.drawString(125, 175, data['Driving Licence'])
    
    can.save()

    # Move the overlay content to the first page of the PDF
    packet.seek(0)
    overlay = PdfReader(packet)
    page = reader.pages[0]
    page.merge_page(overlay.pages[0])
    writer.add_page(page)

    # Write the final PDF
    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output

# Function to extract individual data from dataframe row
def extract_data(row):
    return {
        'Name': row['First Name'] if pd.notna(row['First Name']) else '',
        'Surname': row['Last Name'] if pd.notna(row['Last Name']) else '',
        'DOB': row['DOB'].strftime('%Y-%m-%d') if pd.notna(row['DOB']) else '',
        'Address': row['Address'] if pd.notna(row['Address']) else '',
        #'Postcode': row['Postcode'] if pd.notna(row['Postcode']) else '',
        'Telephone': str(row['Mobile']) if pd.notna(str(row['Mobile'])) else '',
        'Email': row['Email'] if pd.notna(row['Email']) else '',
        'Job 1 Title': row['Job 1 Title'] if pd.notna(row['Job 1 Title']) else '',
        'Job 1 Company': row['Job 1 Company'] if pd.notna(row['Job 1 Company']) else '',
        'Job 1 Start': row['Job 1 Start'].strftime('%Y %m') if pd.notna(row['Job 1 Start']) else '',
        'Job 1 End': row['Job 1 End'].strftime('%Y %m') if pd.notna(row['Job 1 End']) else '',
        'Job 2 Title': row['Job 2 Title'] if pd.notna(row['Job 2 Title']) else '',
        'Job 2 Company': row['Job 2 Company'] if pd.notna(row['Job 2 Company']) else '',
        'Job 2 Start': row['Job 2 Start'].strftime('%Y %m') if pd.notna(row['Job 2 Start']) else '',
        'Job 2 End': row['Job 2 End'].strftime('%Y %m') if pd.notna(row['Job 2 End']) else '',
        'Course 1 Title': row['1 Course title:'] if pd.notna(row['1 Course title:']) else '',
        'College 1 Name': row['1 College or training name'] if pd.notna(row['1 College or training name']) else '',
        'Year 1 Completion': str(int(row['1 Year of completion'])) if pd.notna(row['1 Year of completion']) else '',
        'Level 1 Qualification': row['1 Level of qualification'] if pd.notna(row['1 Level of qualification']) else '',
        'Course 2 Title': row['2 Course title:'] if pd.notna(row['2 Course title:']) else '',
        'College 2 Name': row['2 College or training name'] if pd.notna(row['2 College or training name']) else '',
        'Year 2 Completion': str(int(row['2 Year of completion'])) if pd.notna(row['2 Year of completion']) else '',
        'Level 2 Qualification': row['2  Level of qualification'] if pd.notna(row['2  Level of qualification']) else '',
        'Hobbies': '',  # Assuming Hobbies field is empty in the given data
        'Languages': '',  # Assuming Languages field is empty in the given data
        #'Driving Licence': 'Yes' if row['Driving Licence'] == 'Yes' else 'No'
    }

# Streamlit app
st.title("Excel to PDF filler")

# Upload the Excel file
excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
pdf_template = st.file_uploader("Upload PDF Template", type=["pdf"])

if excel_file and pdf_template:
    df = pd.read_excel(excel_file)
    
    # Directory to save the output files
    output_dir = "output_pdfs"
    os.makedirs(output_dir, exist_ok=True)

    for index, row in df.iterrows():
        data = extract_data(row)
        output_pdf_path = os.path.join(output_dir, f"{data['Name']}_{data['Surname']}.pdf")
        filled_pdf = fill_pdf(data, pdf_template)
        
        # Save the filled PDF to the output directory
        with open(output_pdf_path, "wb") as f:
            f.write(filled_pdf.getvalue())

    st.success(f"PDFs created successfully in {output_dir}")

    # Display download links for the generated PDFs
    for file in os.listdir(output_dir):
        with open(os.path.join(output_dir, file), "rb") as f:
            btn = st.download_button(
                label=f"Download {file}",
                data=f,
                file_name=file,
                mime="application/pdf"
            )
