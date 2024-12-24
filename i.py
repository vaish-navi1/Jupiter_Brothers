from PyPDF2 import PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import fitz  # PyMuPDF for image extraction
import os

# Function to extract content from input PDF
def extract_input_data(input_pdf):
    reader = PdfReader(input_pdf)
    content = ""
    for page in reader.pages:
        content += page.extract_text() + "\n"
    return content

# Function to extract the main image from PDF
def extract_main_image(input_pdf, output_folder):
    doc = fitz.open(input_pdf)
    for i in range(len(doc)):
        for img_index, img in enumerate(doc[i].get_images(full=True)):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            image_path = os.path.join(output_folder, f"main_image.{image_ext}")
            with open(image_path, "wb") as img_file:
                img_file.write(image_bytes)
            return image_path  # Return the first image found
    return None

# Function to process and format the data
def process_data(input_content):
    lines = input_content.splitlines()
    data = {
        "style": "",
        "sizes": "",
        "email": "",
        "care_address": "",
        "specs": [],
    }
    for line in lines:
        if "Style:" in line:
            data["style"] = line.split(":")[-1].strip()
        elif "Sizes:" in line:
            data["sizes"] = line.split(":")[-1].strip()
        elif "E-mail:" in line:
            data["email"] = line.split(":")[-1].strip()
        elif "Care Address:" in line:
            data["care_address"] = line.split(":")[-1].strip()
        elif "Main Fabric" in line:
            data["specs"].append(["Main Fabric", "100% Organic in-conversion cotton", 3, 0.25])
        elif "Rib" in line:
            data["specs"].append(["Rib", "", 2, 0.75])
        elif "Finish/Treatment" in line:
            data["specs"].append(["Finish/Treatment", "", 1, 2])
        elif "Thread" in line:
            data["specs"].append(["Thread", "", 2, 0.25])
        elif "Main Label" in line:
            data["specs"].append(["Main Label", "100% Recycled Polyester", 1, 0.25])
        elif "Size Label" in line:
            data["specs"].append(["Size Label", "100% Recycled Polyester", 2, 0.25])
        elif "Care Label" in line:
            data["specs"].append(["Care Label", "100% Recycled Polyester", 1, 0.75])
        elif "Hang Tag" in line:
            data["specs"].append(["Hang Tag", "100% Nylon", 1, 0.5])

    return data

# Function to generate the output PDF with a table and the main image
def generate_output_pdf(output_pdf, data, main_image_path):
    doc = SimpleDocTemplate(output_pdf, pagesize=letter)
    elements = []

    styles = getSampleStyleSheet()
    title = Paragraph("Costing Sheet", styles['Title'])
    elements.append(title)

    elements.append(Paragraph(f"Style: {data['style']}", styles['Normal']))
    elements.append(Paragraph(f"Sizes: {data['sizes']}", styles['Normal']))
    elements.append(Paragraph(f"E-mail: {data['email']}", styles['Normal']))
    elements.append(Paragraph(f"Care Address: {data['care_address']}", styles['Normal']))

    # Table for spec sheet
    table_data = [["Placement", "Composition", "Qty", "Rate", "Total"]]
    total_cost = 0
    for spec in data["specs"]:
        total = spec[2] * spec[3]
        total_cost += total
        table_data.append([spec[0], spec[1], spec[2], spec[3], total])

    table_data.append(["", "", "", "Total Cost", total_cost])

    table = Table(table_data, hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(table)

    # Adding the main image
    if main_image_path:
        elements.append(Paragraph("Main Image:", styles['Heading2']))
        elements.append(Image(main_image_path, width=300, height=300))

    doc.build(elements)

# Main script
input_pdf_path = r"C:\Users\vaish\Downloads\input_doc.pdf"
output_pdf_path = r"C:\Users\vaish\Downloads\generated_output.pdf"
image_output_folder =r"C:\Users\vaish\Downloads\extract_images"

if not os.path.exists(image_output_folder):
    os.makedirs(image_output_folder)

input_content = extract_input_data(input_pdf_path)
data = process_data(input_content)
main_image_path = extract_main_image(input_pdf_path, image_output_folder)
generate_output_pdf(output_pdf_path, data, main_image_path)

print(f"Output PDF generated at: {output_pdf_path}")
print(f"Main image extracted to: {main_image_path}")
