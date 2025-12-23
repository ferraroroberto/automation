"""
Excel Format Manager Module

Manages Excel file formatting:
- Saves Excel format specifications (column widths, cell formats) to JSON
- Converts text in columns with "url" in name to clickable URLs
- Applies formatting from JSON specification to Excel files
"""

import json
import logging
import os
import re
from typing import Dict, List, Optional

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.cell.cell import Cell

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def save_excel_format_to_json(excel_path: str, json_path: str) -> bool:
    """Save Excel format specifications to JSON file.
    
    Extracts and saves:
    - Column widths
    - Row heights (header row)
    - Cell formats (font, alignment, fill, borders, number format)
    - Header row formatting
    
    Args:
        excel_path: Path to the Excel file to read format from.
        json_path: Path to save the JSON format specification.
        
    Returns:
        True if successful, False otherwise.
    """
    if not os.path.exists(excel_path):
        logger.error(f"❌ Excel file not found: {excel_path}")
        return False
    
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        
        format_spec = {
            "column_widths": {},
            "row_heights": {},
            "header_format": {},
            "cell_formats": {}
        }
        
        # Extract column widths
        DEFAULT_COLUMN_WIDTH = 8.43  # Excel default column width
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            if col_letter in ws.column_dimensions:
                width = ws.column_dimensions[col_letter].width
                # Save width if it's set and different from default
                if width is not None and width != DEFAULT_COLUMN_WIDTH:
                    format_spec["column_widths"][col_letter] = width
        
        # Extract row heights (focus on header row and first few data rows)
        DEFAULT_ROW_HEIGHT = 15.0  # Excel default row height
        for row_idx in range(1, min(ws.max_row + 1, 11)):  # First 10 rows
            if row_idx in ws.row_dimensions:
                height = ws.row_dimensions[row_idx].height
                # Save height if it's set and different from default
                if height is not None and height != DEFAULT_ROW_HEIGHT:
                    format_spec["row_heights"][row_idx] = height
        
        # Extract header row format (row 1)
        if ws.max_row > 0:
            header_cell = ws[1][0]  # First cell of header row
            if isinstance(header_cell, Cell):
                format_spec["header_format"] = _extract_cell_format(header_cell)
        
        # Extract cell formats for first data row (row 2) as sample
        if ws.max_row > 1:
            for col_idx, cell in enumerate(ws[2], 1):
                col_letter = get_column_letter(col_idx)
                if isinstance(cell, Cell) and cell.value:
                    cell_format = _extract_cell_format(cell)
                    if cell_format:
                        format_spec["cell_formats"][col_letter] = cell_format
        
        # Identify date columns and set dd/mm/yyyy format
        date_columns = []
        if ws.max_row > 0:
            for col_idx, header_cell in enumerate(ws[1], 1):
                col_letter = get_column_letter(col_idx)
                if isinstance(header_cell, Cell) and header_cell.value:
                    header_text = str(header_cell.value).lower()
                    if "date" in header_text:
                        date_columns.append(col_letter)
                        logger.info(f"📅 Found date column: {header_cell.value} (column {col_letter})")
        
        # Ensure date columns have dd/mm/yyyy format
        for col_letter in date_columns:
            if col_letter not in format_spec["cell_formats"]:
                format_spec["cell_formats"][col_letter] = {}
            format_spec["cell_formats"][col_letter]["number_format"] = "dd/mm/yyyy"
        
        # Save to JSON
        os.makedirs(os.path.dirname(json_path), exist_ok=True)
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(format_spec, f, indent=2, ensure_ascii=False)
        
        logger.info(f"✅ Format specification saved to: {json_path}")
        return True
        
    except InvalidFileException as e:
        logger.error(f"❌ Invalid Excel file: {e}")
        return False
    except Exception as e:
        logger.error(f"❌ Error saving format: {e}")
        return False


def _extract_cell_format(cell: Cell) -> Dict:
    """Extract format information from a cell.
    
    Args:
        cell: openpyxl Cell object.
        
    Returns:
        Dictionary with format specifications.
    """
    format_dict = {}
    
    if cell.font:
        format_dict["font"] = {
            "name": cell.font.name,
            "size": cell.font.size,
            "bold": cell.font.bold,
            "italic": cell.font.italic,
            "underline": cell.font.underline,
            "color": str(cell.font.color.rgb) if cell.font.color and cell.font.color.rgb else None
        }
    
    if cell.alignment:
        format_dict["alignment"] = {
            "horizontal": cell.alignment.horizontal,
            "vertical": cell.alignment.vertical,
            "wrap_text": cell.alignment.wrap_text
        }
    
    if cell.fill:
        format_dict["fill"] = {
            "fill_type": cell.fill.fill_type,
            "start_color": str(cell.fill.start_color.rgb) if cell.fill.start_color and cell.fill.start_color.rgb else None,
            "end_color": str(cell.fill.end_color.rgb) if cell.fill.end_color and cell.fill.end_color.rgb else None
        }
    
    if cell.border:
        format_dict["border"] = {
            "left": _extract_side(cell.border.left),
            "right": _extract_side(cell.border.right),
            "top": _extract_side(cell.border.top),
            "bottom": _extract_side(cell.border.bottom)
        }
    
    if cell.number_format and cell.number_format != "General":
        format_dict["number_format"] = cell.number_format
    
    return format_dict


def _extract_side(side: Optional[Side]) -> Optional[Dict]:
    """Extract border side information.
    
    Args:
        side: openpyxl Side object.
        
    Returns:
        Dictionary with side specifications or None.
    """
    if not side or side.style is None:
        return None
    
    # Extract color safely
    color_value = None
    if side.color:
        try:
            if hasattr(side.color, 'rgb') and side.color.rgb:
                color_value = str(side.color.rgb)
            elif hasattr(side.color, 'value') and side.color.value:
                color_value = str(side.color.value)
        except Exception:
            # If color extraction fails, just skip it
            color_value = None
    
    return {
        "style": side.style,
        "color": color_value
    }


def convert_url_columns_to_hyperlinks(excel_path: str) -> bool:
    """Convert text in columns with 'url' in column name to clickable hyperlinks.
    
    Args:
        excel_path: Path to the Excel file to modify.
        
    Returns:
        True if successful, False otherwise.
    """
    if not os.path.exists(excel_path):
        logger.error(f"❌ Excel file not found: {excel_path}")
        return False
    
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        
        # Find columns with "url" in name (case-insensitive)
        url_columns = []
        for col_idx, cell in enumerate(ws[1], 1):
            if isinstance(cell, Cell) and cell.value:
                col_name = str(cell.value).lower()
                if "url" in col_name:
                    url_columns.append(col_idx)
                    logger.info(f"📎 Found URL column: {cell.value} (column {get_column_letter(col_idx)})")
        
        if not url_columns:
            logger.info("ℹ️  No columns with 'url' in name found")
            return True
        
        # Convert text to hyperlinks in identified columns
        converted_count = 0
        for col_idx in url_columns:
            col_letter = get_column_letter(col_idx)
            for row_idx in range(2, ws.max_row + 1):  # Skip header row
                cell = ws[f"{col_letter}{row_idx}"]
                if isinstance(cell, Cell) and cell.value:
                    url_text = str(cell.value).strip()
                    # Only convert if it looks like a URL
                    if url_text and (url_text.startswith("http://") or url_text.startswith("https://")):
                        cell.hyperlink = url_text
                        cell.style = "Hyperlink"
                        converted_count += 1
        
        wb.save(excel_path)
        logger.info(f"✅ Converted {converted_count} URLs to hyperlinks")
        return True
        
    except InvalidFileException as e:
        logger.error(f"❌ Invalid Excel file: {e}")
        return False
    except Exception as e:
        logger.error(f"❌ Error converting URLs: {e}")
        return False


def apply_format_from_json(excel_path: str, json_path: str) -> bool:
    """Apply formatting from JSON specification to Excel file.
    
    Args:
        excel_path: Path to the Excel file to format.
        json_path: Path to the JSON format specification file.
        
    Returns:
        True if successful, False otherwise.
    """
    if not os.path.exists(excel_path):
        logger.error(f"❌ Excel file not found: {excel_path}")
        return False
    
    if not os.path.exists(json_path):
        logger.error(f"❌ JSON format file not found: {json_path}")
        return False
    
    try:
        # Load format specification
        with open(json_path, 'r', encoding='utf-8') as f:
            format_spec = json.load(f)
        
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        
        # Apply column widths
        if "column_widths" in format_spec:
            for col_letter, width in format_spec["column_widths"].items():
                ws.column_dimensions[col_letter].width = width
            logger.info(f"✅ Applied {len(format_spec['column_widths'])} column widths")
        
        # Apply row heights
        if "row_heights" in format_spec:
            for row_idx, height in format_spec["row_heights"].items():
                ws.row_dimensions[int(row_idx)].height = height
            logger.info(f"✅ Applied {len(format_spec['row_heights'])} row heights")
        
        # Apply header format (row 1)
        if "header_format" in format_spec and format_spec["header_format"]:
            for col_idx in range(1, ws.max_column + 1):
                col_letter = get_column_letter(col_idx)
                cell = ws[f"{col_letter}1"]
                _apply_cell_format(cell, format_spec["header_format"])
            logger.info("✅ Applied header format")
        
        # Apply cell formats
        if "cell_formats" in format_spec:
            for col_letter, cell_format in format_spec["cell_formats"].items():
                # Apply to all data rows (skip header)
                for row_idx in range(2, ws.max_row + 1):
                    cell = ws[f"{col_letter}{row_idx}"]
                    _apply_cell_format(cell, cell_format)
            logger.info(f"✅ Applied cell formats for {len(format_spec['cell_formats'])} columns")
        
        wb.save(excel_path)
        logger.info(f"✅ Format applied successfully to: {excel_path}")
        return True
        
    except json.JSONDecodeError as e:
        logger.error(f"❌ Invalid JSON file: {e}")
        return False
    except InvalidFileException as e:
        logger.error(f"❌ Invalid Excel file: {e}")
        return False
    except Exception as e:
        logger.error(f"❌ Error applying format: {e}")
        return False


def _convert_color_string(color_str: Optional[str]) -> Optional[str]:
    """Convert color string to aRGB hex format if needed.
    
    Args:
        color_str: Color string (may be RGB hex or aRGB hex).
        
    Returns:
        aRGB hex string (8 hex digits) or None.
    """
    if not color_str:
        return None
    
    # Skip error messages or invalid strings
    if not isinstance(color_str, str):
        return None
    
    # Skip error message strings
    if "must be of type" in color_str.lower() or "error" in color_str.lower():
        return None
    
    # Remove any 'FF' prefix if present (some formats include it)
    color_str = color_str.upper().strip()
    
    # If it's already 8 hex digits (aRGB), return as is
    if len(color_str) == 8 and all(c in '0123456789ABCDEF' for c in color_str):
        return color_str
    
    # If it's 6 hex digits (RGB), add alpha channel (FF for opaque)
    if len(color_str) == 6 and all(c in '0123456789ABCDEF' for c in color_str):
        return f"FF{color_str}"
    
    # If it starts with 'FF' and is 8 chars, it's already aRGB
    if color_str.startswith('FF') and len(color_str) == 8:
        return color_str
    
    # Try to extract hex from string like "FF000000" or "000000"
    hex_match = re.search(r'([0-9A-F]{6,8})', color_str)
    if hex_match:
        hex_val = hex_match.group(1).upper()
        if len(hex_val) == 6:
            return f"FF{hex_val}"
        elif len(hex_val) == 8:
            return hex_val
    
    # If we can't parse it, return None silently (no warning for error messages)
    return None


def _apply_cell_format(cell: Cell, format_dict: Dict) -> None:
    """Apply format dictionary to a cell.
    
    Args:
        cell: openpyxl Cell object.
        format_dict: Dictionary with format specifications.
    """
    if not isinstance(cell, Cell):
        return
    
    # Apply font
    if "font" in format_dict and format_dict["font"]:
        font_dict = format_dict["font"]
        font_color = _convert_color_string(font_dict.get("color"))
        cell.font = Font(
            name=font_dict.get("name", "Calibri"),
            size=font_dict.get("size", 11),
            bold=font_dict.get("bold", False),
            italic=font_dict.get("italic", False),
            underline=font_dict.get("underline", None),
            color=font_color if font_color else None
        )
    
    # Apply alignment
    if "alignment" in format_dict and format_dict["alignment"]:
        align_dict = format_dict["alignment"]
        cell.alignment = Alignment(
            horizontal=align_dict.get("horizontal", "general"),
            vertical=align_dict.get("vertical", "bottom"),
            wrap_text=align_dict.get("wrap_text", False)
        )
    
    # Apply fill
    if "fill" in format_dict and format_dict["fill"]:
        fill_dict = format_dict["fill"]
        start_color = _convert_color_string(fill_dict.get("start_color"))
        end_color = _convert_color_string(fill_dict.get("end_color"))
        cell.fill = PatternFill(
            fill_type=fill_dict.get("fill_type", "none"),
            start_color=start_color if start_color else None,
            end_color=end_color if end_color else None
        )
    
    # Apply border
    if "border" in format_dict and format_dict["border"]:
        border_dict = format_dict["border"]
        cell.border = Border(
            left=_create_side(border_dict.get("left")),
            right=_create_side(border_dict.get("right")),
            top=_create_side(border_dict.get("top")),
            bottom=_create_side(border_dict.get("bottom"))
        )
    
    # Apply number format
    if "number_format" in format_dict:
        cell.number_format = format_dict["number_format"]


def _create_side(side_dict: Optional[Dict]) -> Optional[Side]:
    """Create a Side object from dictionary.
    
    Args:
        side_dict: Dictionary with side specifications.
        
    Returns:
        Side object or None.
    """
    if not side_dict:
        return None
    
    border_color = _convert_color_string(side_dict.get("color"))
    return Side(
        style=side_dict.get("style", None),
        color=border_color if border_color else None
    )


def main() -> None:
    """Main execution function - example usage."""
    # Load configuration from JSON
    config_path = os.path.join(
        os.path.dirname(__file__),
        "linkedin_profiles_data.json"
    )
    
    if not os.path.exists(config_path):
        logger.error(f"❌ Config file not found: {config_path}")
        return
    
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    destination_file = config.get("destination_file")
    if not destination_file:
        logger.error("❌ 'destination_file' not found in config")
        return
    
    # Create format JSON path (same directory as config)
    format_json_path = os.path.join(
        os.path.dirname(config_path),
        "excel_format_spec.json"
    )
    
    logger.info("🚀 Excel Format Manager")
    logger.info(f"📄 Excel file: {destination_file}")
    logger.info(f"📋 Format JSON: {format_json_path}")
    
    # Example: Save format
    if os.path.exists(destination_file):
        logger.info("\n1️⃣ Saving format to JSON...")
        save_excel_format_to_json(destination_file, format_json_path)
        
        logger.info("\n2️⃣ Converting URL columns to hyperlinks...")
        convert_url_columns_to_hyperlinks(destination_file)
        
        logger.info("\n3️⃣ Applying format from JSON...")
        apply_format_from_json(destination_file, format_json_path)
    else:
        logger.warning(f"⚠️  Excel file does not exist: {destination_file}")


if __name__ == "__main__":
    main()

