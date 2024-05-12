from docx import Document
from flask import Flask, render_template, request, send_file
import pdfkit

app = Flask(__name__)

# Configuration for wkhtmltopdf path
config = pdfkit.configuration(wkhtmltopdf='C:\Users\Dell\Downloads\wkhtmltopdf')

# Route for the form
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Get form data
        name = request.form['name']
        date = request.form['date']
        
        # Placeholder replacements
        replacements = {
            "[[NAME]]": name,
            "[[DATE]]": date,
            # Add more placeholder-replacement pairs as needed
        }

        # Replace placeholders in Word document
        replace_placeholders("original_document.docx", replacements)

        # Convert modified .docx to PDF
        convert_to_pdf("modified_document.docx", "output_document.pdf")

        # Provide the PDF for download
        return send_file("output_document.pdf", as_attachment=True)

    # Render the form template for GET requests
    return render_template('form.html')

# Function to replace placeholders in Word document
def replace_placeholders(doc_path, replacements):
    doc = Document(doc_path)
    for paragraph in doc.paragraphs:
        for placeholder, replacement in replacements.items():
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, replacement)
    doc.save("modified_document.docx")

# Function to convert modified .docx to PDF using pdfkit
def convert_to_pdf(docx_path, pdf_path):
    try:
        pdfkit.from_file(docx_path, pdf_path, configuration=config)
        print("PDF conversion successful!")
    except Exception as e:
        print("PDF conversion failed:", e)

if __name__ == '__main__':
    app.run(debug=True)


