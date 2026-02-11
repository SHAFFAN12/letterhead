import sys
import os
from weasyprint import HTML, CSS
from pdf2docx import Converter

def convert_html_to_pdf(html_path, pdf_path):
    """Converts HTML to PDF using WeasyPrint."""
    print(f"Converting {html_path} to {pdf_path}...")
    try:
        # Load HTML
        # base_url is set to the directory of the HTML file so relative links (css, images) work
        html = HTML(filename=html_path, base_url=os.path.dirname(html_path))
        
        # Determine CSS path (assuming style.css is in the same directory)
        css_path = os.path.join(os.path.dirname(html_path), 'style.css')
        stylesheets = []
        if os.path.exists(css_path):
            stylesheets.append(CSS(filename=css_path))
        
        # Render PDF
        html.write_pdf(pdf_path, stylesheets=stylesheets)
        print(f"Successfully created {pdf_path}")
        return True
    except Exception as e:
        print(f"Error converting to PDF: {e}")
        return False

def convert_pdf_to_docx(pdf_path, docx_path):
    """Converts PDF to DOCX using pdf2docx."""
    print(f"Converting {pdf_path} to {docx_path}...")
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        print(f"Successfully created {docx_path}")
        return True
    except Exception as e:
        print(f"Error converting to DOCX: {e}")
        return False

if __name__ == "__main__":
    # Define paths
    base_dir = os.path.dirname(os.path.abspath(__file__))
    html_file = os.path.join(base_dir, 'index.html')
    pdf_file = os.path.join(base_dir, 'letterhead.pdf')
    docx_file = os.path.join(base_dir, 'letterhead.docx')

    # Check if HTML file exists
    if not os.path.exists(html_file):
        print(f"Error: {html_file} not found.")
        sys.exit(1)

    # Convert HTML to PDF
    if convert_html_to_pdf(html_file, pdf_file):
        # Convert PDF to DOCX
        convert_pdf_to_docx(pdf_file, docx_file)
    else:
        sys.exit(1)
