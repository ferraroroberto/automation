import os
import sys
import logging
from pypdf import PdfReader
from tkinter import Tk, filedialog

# Setup logger
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

def select_pdf_file():
    """Open a file dialog to select a PDF file."""
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF Files", "*.pdf")]
    )
    return file_path

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF and return as string."""
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
        return text
    except Exception as e:
        logger.error(f"Failed to extract text from PDF: {e}")
        return None

def save_text_to_file(text, output_path):
    """Save text to a .txt file."""
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(text)
        logger.info(f"Text successfully saved to: {output_path}")
    except Exception as e:
        logger.error(f"Failed to save text file: {e}")

def count_words(text):
    """Count the number of words in a string."""
    return len(text.split())

def main():
    # Get PDF path from command-line or dialog
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = select_pdf_file()

    if not pdf_path or not os.path.exists(pdf_path):
        logger.error("No file selected or file does not exist.")
        return

    if not pdf_path.lower().endswith(".pdf"):
        logger.error("Selected file is not a PDF.")
        return

    logger.info(f"Processing file: {pdf_path}")

    # Extract text
    text = extract_text_from_pdf(pdf_path)
    if text is None:
        return

    # Save text to file
    txt_path = os.path.splitext(pdf_path)[0] + ".txt"
    save_text_to_file(text, txt_path)

    # Word count
    word_count = count_words(text)
    logger.info(f"Conversion complete. Total words: {word_count}")

if __name__ == "__main__":
    main()
