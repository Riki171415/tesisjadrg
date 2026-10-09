import fitz
import sys

def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

if __name__ == "__main__":
    text = extract_text(r"C:\Users\PUSBIKES-KEMKES\Downloads\mdc\MDC Logic 18-20251017.drawio.pdf")
    print(text)
