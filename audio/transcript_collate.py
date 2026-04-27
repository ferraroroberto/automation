# source > https://claude.ai/chat/2b85a103-e8ca-448c-ad33-d8888e43654f

import logging
import os
import re
from tkinter import filedialog, Tk
from pathlib import Path

log = logging.getLogger(__name__)


def select_directory():
    """Opens a dialog to select a directory and returns the path."""
    root = Tk()
    root.withdraw()
    directory = filedialog.askdirectory(title="Select Directory containing MD files")
    return directory

def get_sequence_number(filename):
    """Extracts the sequence number from filename."""
    match = re.search(r'(\d+)', filename)
    return int(match.group(1)) if match else float('inf')

def process_files(directory):
    """Process all MD files in the directory."""
    md_files = [f for f in os.listdir(directory) if f.endswith('.md')]

    if not md_files:
        log.info("No .md files found in the selected directory.")
        return None

    md_files.sort(key=get_sequence_number)

    combined_content = []

    for filename in md_files:
        filepath = os.path.join(directory, filename)
        with open(filepath, 'r', encoding='utf-8') as file:
            content = file.readlines()

            if content and '[' in content[-1] and ']' in content[-1]:
                content = content[:-1]

            combined_content.extend(content)
            combined_content.append('\n')

    return combined_content

def save_combined_file(directory, content):
    """Prompts for a filename and saves the combined content."""
    if not content:
        return

    root = Tk()
    root.withdraw()
    output_filename = filedialog.asksaveasfilename(
        initialdir=directory,
        title="Save combined file as",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")],
        confirmoverwrite=True
    )

    if output_filename:
        with open(output_filename, 'w', encoding='utf-8') as file:
            file.writelines(content)
        log.info("Combined file saved as: %s", output_filename)
    else:
        log.info("Save operation cancelled.")

def main():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    directory = select_directory()
    if not directory:
        log.info("No directory selected. Exiting.")
        return

    log.info("Processing files...")
    combined_content = process_files(directory)

    if combined_content:
        save_combined_file(directory, combined_content)

    log.info("Process completed.")

if __name__ == "__main__":
    main()
