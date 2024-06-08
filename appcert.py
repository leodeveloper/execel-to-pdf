
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import HexColor
#from reportlab.pdfbase.ttfonts import TTFont
#from reportlab.pdfbase import pdfmetrics
from io import BytesIO
import os

# Load the Excel file
excel_file_path = 'Data/LEARNER DATA 2023.xlsx'
df = pd.read_excel(excel_file_path)

# Print the column names to verify them
#print(df.columns)

# Function to fill the PDF form
def fill_pdf(data, template_path, output_path):
    # Read the existing PDF
    reader = PdfReader(template_path)
    writer = PdfWriter()

    # Create a canvas overlay to add text to the PDF
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    page_width, page_height = letter

    def draw_centered_text(text, y, font, font_size, color):
        can.setFont(font, font_size)
        can.setFillColor(HexColor(color))
        text_width = can.stringWidth(text, font, font_size)
        can.drawString(((page_width+180) - text_width) / 2, y, text)

    # First set of fields
    FullName = data['Name'] + data['Surname']
    draw_centered_text(FullName.upper(), 415, "Helvetica", 28, "#bf9237")  # Example color: orange


    #pdfmetrics.registerFont(TTFont('CustomFont', 'custom_font.ttf'))
    #can.setFont("Helvetica", 28)
    #can.setFillColor(HexColor("#bf9237"))  # Example color: orange

    # Fill the form fields with data
    #can.drawString(136, 415, data['Name'] + data['Surname'])
    #can.drawString(300, 552, data['Surname'])

    # Set different font and color for the second set of fields
    can.setFont("Courier", 16)
    can.setFillColor(HexColor("#595959"))  # Example color: green
    draw_centered_text(data['Level 1 Qualification'], 292, "Courier", 16, "#595959")
    
    can.save()

    # Move the overlay content to the first page of the PDF
    packet.seek(0)
    overlay = PdfReader(packet)
    page = reader.pages[0]
    page.merge_page(overlay.pages[0])
    writer.add_page(page)

    # Write the final PDF
    with open(output_path, "wb") as output_stream:
        writer.write(output_stream)

# Function to extract individual data from dataframe row
def extract_data(row):
    return {
        'Name': row['First Name'] if pd.notna(row['First Name']) else '',
        'Surname': row['Last Name'] if pd.notna(row['Last Name']) else '',
        'Level 1 Qualification': row['Last Name'] if pd.notna(row['Last Name']) else ''
    }

# Path to the PDF template and output directory
pdf_template_path = 'Data/Emergency First Aid at Work.pdf'
output_dir = 'DataOutPut/'
os.makedirs(output_dir, exist_ok=True)

# Loop through each row in the dataframe and create a filled PDF
for index, row in df.iterrows():
    data = extract_data(row)
    output_pdf_path = os.path.join(output_dir, f"{data['Name']}_{data['Surname']}.pdf")
    fill_pdf(data, pdf_template_path, output_pdf_path)



#Name font Book Antiqua and size 28
#Date fonr Calibri Bold and size 16
