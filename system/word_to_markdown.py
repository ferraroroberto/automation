#!/usr/bin/env python3
"""
Word to Markdown Converter

A dual-interface application that converts Word documents to Markdown format.
Supports both command-line and Tkinter file picker interfaces.
Automatically tests conversion by processing all Word documents found in the project.
"""

import argparse
import glob
import re
import logging
import os
import shutil
import sys
from pathlib import Path
from typing import List, Optional

try:
    from docx import Document
    from docx.document import Document as DocumentType
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph
except ImportError:
    print("❌ Error: python-docx package not found.")
    print("Please install it with: pip install python-docx")
    sys.exit(1)

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False
    print("⚠️  Warning: tkinter not available. GUI interface will be disabled.")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class WordToMarkdownConverter:
    """Converts Word documents to Markdown format."""
    
    def __init__(self):
        self.project_root = self._find_project_root()
        # Project-root-relative (not cwd-relative), so running from anywhere
        # other than the repo root doesn't scatter test folders under the
        # wrong CWD (audit issue #67).
        self.test_folder = self.project_root / "system/test-word-to-markdown"
        
    def _find_project_root(self) -> Path:
        """Find the project root directory by looking for .git folder."""
        current = Path.cwd()
        while current != current.parent:
            if (current / ".git").exists():
                return current
            current = current.parent
        return Path.cwd()
    
    def _create_test_folder(self):
        """Create the test folder if it doesn't exist."""
        self.test_folder.mkdir(parents=True, exist_ok=True)
        logger.info(f"📁 Test folder created/verified: {self.test_folder}")
    
    def _find_word_documents(self) -> List[Path]:
        """Find all Word documents in the project."""
        pattern = "**/*.docx"
        documents = []
        
        for doc_path in self.project_root.glob(pattern):
            # Skip the test folder itself
            if self.test_folder in doc_path.parents:
                continue
            documents.append(doc_path)
        
        logger.info(f"🔍 Found {len(documents)} Word documents in project")
        return documents
    
    def _copy_to_test_folder(self, doc_path: Path) -> Path:
        """Copy document to test folder."""
        dest_path = self.test_folder / doc_path.name
        shutil.copy2(doc_path, dest_path)
        logger.debug(f"📋 Copied {doc_path.name} to test folder")
        return dest_path
    
    def _extract_paragraphs_and_tables(self, doc: DocumentType) -> List:
        """Extract paragraphs and tables in document order."""
        elements = []
        for element in doc.element.body:
            if isinstance(element, CT_P):
                elements.append(Paragraph(element, doc))
            elif isinstance(element, CT_Tbl):
                elements.append(Table(element, doc))
        return elements
    
    def _paragraph_to_markdown(self, paragraph: Paragraph) -> str:
        """Convert a paragraph to Markdown format."""
        text = paragraph.text.strip()
        if not text:
            return ""
        
        # Check for heading styles (standard "Heading N" and "Title")
        style_name = paragraph.style.name.lower().strip().replace('\xa0', ' ')

        logger.debug(
            f"🧪 Paragraph text preview='{text[:60]}' | p_style='{paragraph.style.name}'"
        )

        # Determine heading level from paragraph style
        heading_level: Optional[int] = None
        heading_pattern = re.compile(r"heading\s*([1-6])", re.IGNORECASE)
        match = heading_pattern.search(style_name)
        if match:
            heading_level = int(match.group(1))

        has_title = 'title' in style_name
        logger.debug(
            f"🔎 style_name='{style_name}' | heading_level={heading_level} | has_title={has_title}"
        )

        if heading_level is not None or has_title:
            level = heading_level if heading_level is not None else 1
            logger.debug(f"🏷️ Rendering as H{level}")
            if level == 1:
                return f"# {text}\n\n"
            if level == 2:
                return f"## {text}\n\n"
            if level == 3:
                return f"### {text}\n\n"
            if level == 4:
                return f"#### {text}\n\n"
            if level == 5:
                return f"##### {text}\n\n"
            if level == 6:
                return f"###### {text}\n\n"
        
        # Check for list styles
        if paragraph.style.name.lower().startswith('list'):
            logger.debug("• Rendering as list item")
            return f"- {text}\n"
        
        # Regular paragraph
        logger.debug("📝 Rendering as paragraph")
        return f"{text}\n\n"
    
    def _table_to_markdown(self, table: Table) -> str:
        """Convert a table to Markdown format."""
        if not table.rows:
            return ""
        
        markdown = []
        
        # Add headers
        headers = []
        for cell in table.rows[0].cells:
            headers.append(cell.text.strip())
        
        markdown.append("| " + " | ".join(headers) + " |")
        markdown.append("| " + " | ".join(["---"] * len(headers)) + " |")
        
        # Add data rows
        for row in table.rows[1:]:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text.strip())
            markdown.append("| " + " | ".join(row_data) + " |")
        
        markdown.append("")  # Add empty line after table
        return "\n".join(markdown)
    
    def convert_document(self, doc_path: Path) -> str:
        """Convert a Word document to Markdown."""
        try:
            logger.info(f"🔄 Converting {doc_path.name} to Markdown...")
            
            # Load the document
            doc = Document(doc_path)
            
            # Extract content
            elements = self._extract_paragraphs_and_tables(doc)
            
            # Convert to Markdown
            markdown_content = []
            
            for element in elements:
                if isinstance(element, Paragraph):
                    markdown_content.append(self._paragraph_to_markdown(element))
                elif isinstance(element, Table):
                    markdown_content.append(self._table_to_markdown(element))
            
            # Join all content
            markdown_text = "".join(markdown_content).strip()
            
            logger.info(f"✅ Successfully converted {doc_path.name}")
            return markdown_text
            
        except Exception as e:
            logger.error(f"❌ Error converting {doc_path.name}: {str(e)}")
            return f"# Error converting document\n\nAn error occurred while converting this document: {str(e)}"
    
    def save_markdown(self, content: str, output_path: Path):
        """Save Markdown content to file."""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
            logger.info(f"💾 Markdown saved to: {output_path}")
        except Exception as e:
            logger.error(f"❌ Error saving {output_path}: {str(e)}")
    
    def convert_single_document(self, doc_path: str) -> bool:
        """Convert a single document specified by path."""
        doc_path = Path(doc_path)
        
        if not doc_path.exists():
            logger.error(f"❌ File not found: {doc_path}")
            return False
        
        if not doc_path.suffix.lower() == '.docx':
            logger.error(f"❌ File is not a Word document: {doc_path}")
            return False
        
        # Convert to Markdown
        markdown_content = self.convert_document(doc_path)
        
        # Create output path with .md extension
        output_path = doc_path.with_suffix('.md')
        
        # Save Markdown
        self.save_markdown(markdown_content, output_path)
        
        return True
    
    def run_test_conversion(self):
        """Run automatic test conversion of all Word documents in project."""
        logger.info("🧪 Starting automatic test conversion...")
        
        # Create test folder
        self._create_test_folder()
        
        # Find all Word documents
        documents = self._find_word_documents()
        
        if not documents:
            logger.info("ℹ️  No Word documents found in project")
            return
        
        # Process each document
        success_count = 0
        for doc_path in documents:
            try:
                # Copy to test folder
                test_doc_path = self._copy_to_test_folder(doc_path)
                
                # Convert to Markdown
                markdown_content = self.convert_document(test_doc_path)
                
                # Save Markdown with same name + .md extension
                output_path = self.test_folder / f"{doc_path.stem}.md"
                self.save_markdown(markdown_content, output_path)
                
                success_count += 1
                
            except Exception as e:
                logger.error(f"❌ Failed to process {doc_path.name}: {str(e)}")
        
        logger.info(f"🎉 Test conversion completed: {success_count}/{len(documents)} documents processed")
        logger.info(f"📁 Results saved in: {self.test_folder}")


class TkinterInterface:
    """Simple Tkinter interface for file selection."""
    
    def __init__(self, converter: WordToMarkdownConverter):
        self.converter = converter
        self.root = tk.Tk()
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the user interface."""
        self.root.title("Word to Markdown Converter")
        self.root.geometry("400x200")
        self.root.resizable(False, False)
        
        # Center the window
        self.root.eval('tk::PlaceWindow . center')
        
        # Main frame
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(expand=True, fill='both')
        
        # Title
        title_label = tk.Label(main_frame, text="Word to Markdown Converter", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Instructions
        instruction_label = tk.Label(main_frame, 
                                   text="Click 'Select File' to choose a Word document\nfor conversion to Markdown format.",
                                   justify='center')
        instruction_label.pack(pady=(0, 20))
        
        # Select file button
        select_button = tk.Button(main_frame, text="Select File", 
                                command=self.select_file, 
                                font=("Arial", 12), 
                                width=15, height=2)
        select_button.pack(pady=(0, 10))
        
        # Test conversion button
        test_button = tk.Button(main_frame, text="Run Test Conversion", 
                               command=self.run_test_conversion, 
                               font=("Arial", 10), 
                               width=15, height=1)
        test_button.pack()
        
        # Status label
        self.status_label = tk.Label(main_frame, text="", fg="blue")
        self.status_label.pack(pady=(10, 0))
    
    def select_file(self):
        """Open file dialog and convert selected file."""
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        
        if file_path:
            self.status_label.config(text="Converting...", fg="orange")
            self.root.update()
            
            try:
                success = self.converter.convert_single_document(file_path)
                if success:
                    self.status_label.config(text="✅ Conversion completed!", fg="green")
                    messagebox.showinfo("Success", "Document converted successfully!")
                else:
                    self.status_label.config(text="❌ Conversion failed", fg="red")
                    messagebox.showerror("Error", "Failed to convert document")
            except Exception as e:
                self.status_label.config(text="❌ Error occurred", fg="red")
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def run_test_conversion(self):
        """Run the automatic test conversion."""
        self.status_label.config(text="Running test conversion...", fg="orange")
        self.root.update()
        
        try:
            self.converter.run_test_conversion()
            self.status_label.config(text="✅ Test conversion completed!", fg="green")
            messagebox.showinfo("Success", "Test conversion completed successfully!")
        except Exception as e:
            self.status_label.config(text="❌ Test conversion failed", fg="red")
            messagebox.showerror("Error", f"Test conversion failed: {str(e)}")
    
    def run(self):
        """Start the Tkinter interface."""
        self.root.mainloop()


def main():
    """Main function with command-line interface."""
    parser = argparse.ArgumentParser(
        description="Convert Word documents to Markdown format"
    )
    parser.add_argument(
        "file_path", 
        nargs='?', 
        help="Path to the Word document to convert"
    )
    parser.add_argument(
        "--test", 
        action="store_true", 
        help="Run automatic test conversion of all Word documents in project"
    )
    parser.add_argument(
        "--gui", 
        action="store_true", 
        help="Launch the Tkinter GUI interface"
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug logging"
    )
    
    args = parser.parse_args()
    
    # Initialize converter
    converter = WordToMarkdownConverter()
    if args.debug:
        logger.setLevel(logging.DEBUG)
        logger.debug("🔧 Debug logging enabled")
    
    if args.gui:
        if not TKINTER_AVAILABLE:
            logger.error("❌ Tkinter not available. Cannot launch GUI.")
            return
        
        logger.info("🖥️  Launching Tkinter interface...")
        interface = TkinterInterface(converter)
        interface.run()
        
    elif args.test:
        logger.info("🧪 Running automatic test conversion...")
        converter.run_test_conversion()
        
    elif args.file_path:
        logger.info(f"🔄 Converting single document: {args.file_path}")
        success = converter.convert_single_document(args.file_path)
        if success:
            logger.info("✅ Conversion completed successfully!")
        else:
            logger.error("❌ Conversion failed!")
            sys.exit(1)
            
    else:
        # No arguments provided, show help
        if TKINTER_AVAILABLE:
            logger.info("🖥️  No arguments provided. Launching GUI interface...")
            interface = TkinterInterface(converter)
            interface.run()
        else:
            parser.print_help()
            logger.info("💡 Use --help for usage information")


if __name__ == "__main__":
    main()