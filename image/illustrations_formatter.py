#!/usr/bin/env python3
"""
illustrations_formatter.py - Core module for image formatting

This module provides functionality to format images for Instagram or
to fixed dimensions (1920x1080) by adding padding.

"""

import os
import sys
import time
import logging
import argparse
import json
from pathlib import Path
from typing import Tuple, List, Optional, Dict, Any
from PIL import Image
import concurrent.futures
from dataclasses import dataclass


@dataclass
class ProcessingResult:
    """Data class to store processing results"""
    total_images: int
    successful: int
    failed: int
    skipped: int
    elapsed_time: float
    errors: List[Tuple[str, str]]


class IllustrationsFormatter:
    """Main class for formatting images"""
    
    SUPPORTED_FORMATS = {'.png', '.jpg', '.jpeg', '.webp', '.bmp'}
    DEFAULT_ASPECT_RATIOS = {
        '3:4': 0.75,     # Instagram portrait (3:4)
        '4:5': 0.8,      # Instagram portrait (4:5)
        '9:16': 0.5625,  # Instagram Stories/Reels (9:16)
        '1:1': 1.0,      # Square (no change needed)
        '16:9': 1.7778,  # Landscape
    }
    CONFIG_FILE = 'illustrations_formatter_config.json'
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        """Initialize the formatter with optional logger"""
        self.logger = logger or self._setup_default_logger()
        self.config = self.load_config()
        
    def _setup_default_logger(self) -> logging.Logger:
        """Setup default logger configuration"""
        logger = logging.getLogger('IllustrationsFormatter')
        logger.setLevel(logging.INFO)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # Format
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        console_handler.setFormatter(formatter)
        
        logger.addHandler(console_handler)
        return logger
    
    def load_config(self) -> Dict[str, Any]:
        """Load configuration from JSON file"""
        config_path = Path(__file__).parent / self.CONFIG_FILE
        
        # Default configuration
        default_config = {
            'source_folder': '',
            'destination_folder': '',
            'destination_folder_instagram': r'C:\Users\rober\iCloudDrive\6LVTQB9699~com~seriflabs~affinitydesigner\Roberto\archived_IGformat',
            'destination_folder_1920x1080': r'C:\Users\rober\iCloudDrive\6LVTQB9699~com~seriflabs~affinitydesigner\Roberto\archived_1920x1080',
            'aspect_ratio': '3:4',
            'background_color': '',
            'format_type': 'instagram'
        }
        
        try:
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                # Merge with defaults to ensure all keys exist
                default_config.update(config)
                return default_config
        except (json.JSONDecodeError, IOError) as e:
            self.logger.warning(f"Could not load config file: {e}")
        
        return default_config
    
    def parse_aspect_ratio(self, ratio_str: str) -> float:
        """
        Parse aspect ratio string and return float value
        
        Args:
            ratio_str: Aspect ratio as string (e.g., '4:5', '9:16', '0.8')
            
        Returns:
            Float representation of the aspect ratio
            
        Raises:
            ValueError: If the format is invalid
        """
        # Check if it's a predefined ratio
        if ratio_str in self.DEFAULT_ASPECT_RATIOS:
            return self.DEFAULT_ASPECT_RATIOS[ratio_str]
        
        # Try to parse as float
        try:
            return float(ratio_str)
        except ValueError:
            pass
        
        # Try to parse as ratio (e.g., '4:5')
        if ':' in ratio_str:
            try:
                width, height = map(float, ratio_str.split(':'))
                if height == 0:
                    raise ValueError("Height cannot be zero")
                return width / height
            except (ValueError, TypeError):
                pass
        
        raise ValueError(f"Invalid aspect ratio format: {ratio_str}")
    
    def get_background_color(self, image: Image.Image) -> Tuple[int, int, int]:
        """
        Extract background color from image corners
        
        Args:
            image: PIL Image object
            
        Returns:
            RGB tuple of the background color
        """
        # Sample pixels from corners
        width, height = image.size
        corners = [
            (0, 0),                    # Top-left
            (width - 1, 0),           # Top-right
            (0, height - 1),          # Bottom-left
            (width - 1, height - 1)   # Bottom-right
        ]
        
        # Get colors from corners
        colors = []
        for x, y in corners:
            pixel = image.getpixel((x, y))
            if isinstance(pixel, tuple) and len(pixel) >= 3:
                colors.append(pixel[:3])  # Get RGB, ignore alpha if present
            else:
                # Grayscale image
                colors.append((pixel, pixel, pixel))
        
        # Use the most common color or average
        # For simplicity, we'll use the first corner color
        # In production, you might want to implement more sophisticated logic
        return colors[0]
    
    def process_image(self, input_path: Path, output_path: Path, 
                     aspect_ratio: float, background_color: Optional[Tuple[int, int, int]] = None) -> None:
        """
        Process a single image to add padding for aspect ratio format
        
        Args:
            input_path: Path to input image
            output_path: Path to save processed image
            aspect_ratio: Target aspect ratio (width/height)
            background_color: Optional RGB tuple for background color
            
        Raises:
            Exception: If image processing fails
        """
        self.logger.info(f"Processing: {input_path.name}")
        
        # Open image
        with Image.open(input_path) as img:
            # Convert to RGB if necessary
            if img.mode not in ('RGB', 'RGBA'):
                img = img.convert('RGB')
            
            original_width, original_height = img.size
            current_ratio = original_width / original_height
            
            # Check if image is already in desired ratio (with small tolerance)
            if abs(current_ratio - aspect_ratio) < 0.01:
                # Just save a copy
                img.save(output_path, 'PNG')
                self.logger.info(f"Image already in target ratio: {input_path.name}")
                return
            
            # Determine background color if not provided
            if background_color is None:
                background_color = self.get_background_color(img)
            
            # Calculate new dimensions
            if current_ratio > aspect_ratio:
                # Image is wider than target - add padding top and bottom
                new_width = original_width
                new_height = int(original_width / aspect_ratio)
            else:
                # Image is taller than target - add padding left and right
                new_height = original_height
                new_width = int(original_height * aspect_ratio)
            
            # Create new image with padding
            new_img = Image.new('RGB', (new_width, new_height), background_color)
            
            # Calculate position to paste original image (centered)
            x_offset = (new_width - original_width) // 2
            y_offset = (new_height - original_height) // 2
            
            # Paste original image onto new image
            if img.mode == 'RGBA':
                new_img.paste(img, (x_offset, y_offset), img)
            else:
                new_img.paste(img, (x_offset, y_offset))
            
            # Save the result
            output_path.parent.mkdir(parents=True, exist_ok=True)
            new_img.save(output_path, 'PNG')
            
            self.logger.info(f"Saved: {output_path.name} ({original_width}x{original_height} -> {new_width}x{new_height})")
    
    def process_image_fixed_size(self, input_path: Path, output_path: Path,
                                 target_width: int, target_height: int,
                                 background_color: Optional[Tuple[int, int, int]] = None) -> None:
        """
        Process a single image to fixed dimensions (1920x1080)
        
        Args:
            input_path: Path to input image
            output_path: Path to save processed image
            target_width: Target width in pixels
            target_height: Target height in pixels
            background_color: Optional RGB tuple for background color
            
        Raises:
            Exception: If image processing fails
        """
        self.logger.info(f"Processing: {input_path.name}")
        
        # Open image
        with Image.open(input_path) as img:
            # Convert to RGB if necessary
            if img.mode not in ('RGB', 'RGBA'):
                img = img.convert('RGB')
            
            original_width, original_height = img.size
            
            # Check if image is already in target size
            if original_width == target_width and original_height == target_height:
                # Just save a copy
                img.save(output_path, 'PNG')
                self.logger.info(f"Image already in target size: {input_path.name}")
                return
            
            # Determine background color if not provided
            if background_color is None:
                background_color = self.get_background_color(img)
            
            # Calculate scaling to fit within target dimensions while maintaining aspect ratio
            scale_w = target_width / original_width
            scale_h = target_height / original_height
            scale = min(scale_w, scale_h)  # Use smaller scale to fit within bounds
            
            # Calculate new dimensions
            new_width = int(original_width * scale)
            new_height = int(original_height * scale)
            
            # Resize image
            resized_img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # Create new image with target dimensions
            new_img = Image.new('RGB', (target_width, target_height), background_color)
            
            # Calculate position to paste resized image (centered)
            x_offset = (target_width - new_width) // 2
            y_offset = (target_height - new_height) // 2
            
            # Paste resized image onto new image
            if resized_img.mode == 'RGBA':
                new_img.paste(resized_img, (x_offset, y_offset), resized_img)
            else:
                new_img.paste(resized_img, (x_offset, y_offset))
            
            # Save the result
            output_path.parent.mkdir(parents=True, exist_ok=True)
            new_img.save(output_path, 'PNG')
            
            self.logger.info(f"Saved: {output_path.name} ({original_width}x{original_height} -> {target_width}x{target_height})")
    
    def process_folder(self, source_folder: str, destination_folder: str, 
                      aspect_ratio: str, background_color: Optional[Tuple[int, int, int]] = None,
                      progress_callback=None) -> ProcessingResult:
        """
        Process all images in a folder
        
        Args:
            source_folder: Path to source folder
            destination_folder: Path to destination folder
            aspect_ratio: Target aspect ratio as string
            background_color: Optional RGB tuple for background color
            progress_callback: Optional callback function for progress updates
            
        Returns:
            ProcessingResult object with statistics
        """
        start_time = time.time()
        errors = []
        
        # Validate inputs
        source_path = Path(source_folder)
        dest_path = Path(destination_folder)
        
        if not source_path.exists():
            raise ValueError(f"Source folder does not exist: {source_folder}")
        
        if not source_path.is_dir():
            raise ValueError(f"Source path is not a directory: {source_folder}")
        
        # Parse aspect ratio
        try:
            ratio_float = self.parse_aspect_ratio(aspect_ratio)
        except ValueError as e:
            raise ValueError(f"Invalid aspect ratio: {e}")
        
        # Create destination folder
        dest_path.mkdir(parents=True, exist_ok=True)
        
        # Find all image files
        image_files = []
        for ext in self.SUPPORTED_FORMATS:
            image_files.extend(source_path.glob(f'*{ext}'))
            image_files.extend(source_path.glob(f'*{ext.upper()}'))
        
        # Remove duplicates and filter square/rectangular images
        image_files = list(set(image_files))
        total_images = len(image_files)
        
        if total_images == 0:
            self.logger.warning(f"No supported image files found in: {source_folder}")
            return ProcessingResult(0, 0, 0, 0, 0.0, [])
        
        self.logger.info(f"Found {total_images} image(s) to process")
        
        successful = 0
        failed = 0
        skipped = 0
        
        # Process images
        for idx, img_path in enumerate(image_files, 1):
            try:
                output_path = dest_path / img_path.name
                
                # Skip if the destination file already exists and has the correct ratio
                if output_path.exists():
                    try:
                        with Image.open(output_path) as existing_img:
                            existing_ratio = existing_img.width / existing_img.height
                            if abs(existing_ratio - ratio_float) < 0.01:
                                self.logger.info(f"Skipping already processed image: {output_path.name}")
                                skipped += 1
                                if progress_callback:
                                    progress_callback(idx, total_images, f"Skipping: {img_path.name}")
                                continue
                    except Exception as e:
                        self.logger.warning(f"Could not check existing file {output_path.name}, will re-process. Error: {e}")

                self.process_image(img_path, output_path, ratio_float, background_color)
                successful += 1
                
                if progress_callback:
                    progress_callback(idx, total_images, f"Processing: {img_path.name}")
                    
            except Exception as e:
                failed += 1
                error_msg = f"Failed to process {img_path.name}: {str(e)}"
                self.logger.error(error_msg)
                errors.append((img_path.name, str(e)))
                
                if progress_callback:
                    progress_callback(idx, total_images, f"Error: {img_path.name}")
        
        elapsed_time = time.time() - start_time
        
        # Log summary
        self.logger.info(f"\nProcessing complete!")
        self.logger.info(f"Total images: {total_images}")
        self.logger.info(f"Successful: {successful}")
        self.logger.info(f"Skipped: {skipped}")
        self.logger.info(f"Failed: {failed}")
        self.logger.info(f"Time elapsed: {elapsed_time:.2f} seconds")
        
        return ProcessingResult(total_images, successful, failed, skipped, elapsed_time, errors)
    
    def process_folder_fixed_size(self, source_folder: str, destination_folder: str,
                                  target_width: int, target_height: int,
                                  background_color: Optional[Tuple[int, int, int]] = None,
                                  progress_callback=None) -> ProcessingResult:
        """
        Process all images in a folder to fixed dimensions
        
        Args:
            source_folder: Path to source folder
            destination_folder: Path to destination folder
            target_width: Target width in pixels
            target_height: Target height in pixels
            background_color: Optional RGB tuple for background color
            progress_callback: Optional callback function for progress updates
            
        Returns:
            ProcessingResult object with statistics
        """
        start_time = time.time()
        errors = []
        
        # Validate inputs
        source_path = Path(source_folder)
        dest_path = Path(destination_folder)
        
        if not source_path.exists():
            raise ValueError(f"Source folder does not exist: {source_folder}")
        
        if not source_path.is_dir():
            raise ValueError(f"Source path is not a directory: {source_folder}")
        
        # Create destination folder
        dest_path.mkdir(parents=True, exist_ok=True)
        
        # Find all image files
        image_files = []
        for ext in self.SUPPORTED_FORMATS:
            image_files.extend(source_path.glob(f'*{ext}'))
            image_files.extend(source_path.glob(f'*{ext.upper()}'))
        
        # Remove duplicates
        image_files = list(set(image_files))
        total_images = len(image_files)
        
        if total_images == 0:
            self.logger.warning(f"No supported image files found in: {source_folder}")
            return ProcessingResult(0, 0, 0, 0, 0.0, [])
        
        self.logger.info(f"Found {total_images} image(s) to process")
        
        successful = 0
        failed = 0
        skipped = 0
        
        # Process images
        for idx, img_path in enumerate(image_files, 1):
            try:
                output_path = dest_path / img_path.name
                
                # Skip if the destination file already exists and has the correct size
                if output_path.exists():
                    try:
                        with Image.open(output_path) as existing_img:
                            if existing_img.width == target_width and existing_img.height == target_height:
                                self.logger.info(f"Skipping already processed image: {output_path.name}")
                                skipped += 1
                                if progress_callback:
                                    progress_callback(idx, total_images, f"Skipping: {img_path.name}")
                                continue
                    except Exception as e:
                        self.logger.warning(f"Could not check existing file {output_path.name}, will re-process. Error: {e}")

                self.process_image_fixed_size(img_path, output_path, target_width, target_height, background_color)
                successful += 1
                
                if progress_callback:
                    progress_callback(idx, total_images, f"Processing: {img_path.name}")
                    
            except Exception as e:
                failed += 1
                error_msg = f"Failed to process {img_path.name}: {str(e)}"
                self.logger.error(error_msg)
                errors.append((img_path.name, str(e)))
                
                if progress_callback:
                    progress_callback(idx, total_images, f"Error: {img_path.name}")
        
        elapsed_time = time.time() - start_time
        
        # Log summary
        self.logger.info(f"\nProcessing complete!")
        self.logger.info(f"Total images: {total_images}")
        self.logger.info(f"Successful: {successful}")
        self.logger.info(f"Skipped: {skipped}")
        self.logger.info(f"Failed: {failed}")
        self.logger.info(f"Time elapsed: {elapsed_time:.2f} seconds")
        
        return ProcessingResult(total_images, successful, failed, skipped, elapsed_time, errors)


def parse_color(color_str: str) -> Tuple[int, int, int]:
    """Parse color string to RGB tuple"""
    if not color_str:
        return None
    
    # Remove common prefixes
    color_str = color_str.strip().lstrip('#').lower()
    
    # Parse hex color
    if len(color_str) == 6:
        try:
            r = int(color_str[0:2], 16)
            g = int(color_str[2:4], 16)
            b = int(color_str[4:6], 16)
            return (r, g, b)
        except ValueError:
            pass
    
    # Parse comma-separated RGB
    if ',' in color_str:
        try:
            parts = [int(x.strip()) for x in color_str.split(',')]
            if len(parts) == 3 and all(0 <= x <= 255 for x in parts):
                return tuple(parts)
        except ValueError:
            pass
    
    raise ValueError(f"Invalid color format: {color_str}")


def main():
    """Main entry point for command line usage"""
    # Load configuration first
    formatter = IllustrationsFormatter()
    config = formatter.config
    
    parser = argparse.ArgumentParser(
        description='Format images for Instagram or fixed dimensions',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s -s ./images -d ./output -r 4:5
  %(prog)s -s ./photos -d ./formatted -r 9:16 -c 255,255,255
  %(prog)s -s ./input -d ./output -r 0.8 -c "#FF5733"
        """
    )
    
    parser.add_argument('-s', '--source', 
                       default=config.get('source_folder', ''),
                       help='Source folder containing images (default from config)')
    parser.add_argument('-d', '--destination', 
                       default=config.get('destination_folder', ''),
                       help='Destination folder for processed images (default from config)')
    parser.add_argument('-r', '--ratio', 
                       default=config.get('aspect_ratio', '3:4'),
                       help='Target aspect ratio (e.g., 4:5, 9:16, 0.8) (default from config)')
    parser.add_argument('-c', '--color', 
                       default=config.get('background_color', ''),
                       help='Background color as hex (#RRGGBB) or RGB (R,G,B) (default from config)')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='Enable verbose logging')
    
    args = parser.parse_args()
    
    # Check if required arguments are provided either via CLI or config
    if not args.source:
        print("Error: Source folder must be specified via -s/--source or in config file", file=sys.stderr)
        sys.exit(1)
    
    if not args.destination:
        print("Error: Destination folder must be specified via -d/--destination or in config file", file=sys.stderr)
        sys.exit(1)
    
    # Setup logging
    if args.verbose:
        formatter.logger.setLevel(logging.DEBUG)
    
    # Parse background color
    bg_color = None
    if args.color:
        try:
            bg_color = parse_color(args.color)
        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)
    
    # Process images
    try:
        result = formatter.process_folder(
            args.source,
            args.destination,
            args.ratio,
            bg_color
        )
        
        # Print summary
        print(f"\nProcessing complete!")
        print(f"Total images: {result.total_images}")
        print(f"Successful: {result.successful}")
        print(f"Skipped: {result.skipped}")
        print(f"Failed: {result.failed}")
        print(f"Time elapsed: {result.elapsed_time:.2f} seconds")
        
        if result.errors:
            print("\nErrors:")
            for filename, error in result.errors:
                print(f"  - {filename}: {error}")
        
        # Exit with error code if any failures
        sys.exit(1 if result.failed > 0 else 0)
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()