#!/usr/bin/env python3
"""
illustrations_formatter.py - Core module for image formatting

Provides functionality to format images for Instagram aspect ratios or to
exact pixel dimensions (1920 × 1080), adding padding as required.

Key features
------------
* Batch-process an entire folder of images.
* Pad images to a target aspect ratio (Instagram) or a fixed canvas size
  (1920 × 1080).
* Auto-detect background colour from the image corners.
* **Extend-border mode** – instead of filling the padding strips with a flat
  colour, the outermost column / row of pixels is replicated outward.
  This eliminates the visible colour-band artefact that occurs when the
  image border colour does not exactly match the detected corner colour
  (e.g. the iceberg illustration whose ocean edge differs from the plain
  background).
* Convert a *single* square image to 1920 × 1080 via
  :meth:`IllustrationsFormatter.convert_single_to_1920x1080`.
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
    """Aggregated statistics returned after a batch processing run."""
    total_images: int
    successful: int
    failed: int
    skipped: int
    elapsed_time: float
    errors: List[Tuple[str, str]]


class IllustrationsFormatter:
    """Main class for formatting images."""

    SUPPORTED_FORMATS = {'.png', '.jpg', '.jpeg', '.webp', '.bmp'}
    DEFAULT_ASPECT_RATIOS = {
        '3:4':  0.75,    # Instagram portrait (3:4)
        '4:5':  0.8,     # Instagram portrait (4:5)
        '9:16': 0.5625,  # Instagram Stories / Reels (9:16)
        '1:1':  1.0,     # Square
        '16:9': 1.7778,  # Landscape
    }
    CONFIG_FILE = 'illustrations_formatter_config.json'

    def __init__(self, logger: Optional[logging.Logger] = None):
        """Initialize the formatter with an optional custom logger."""
        self.logger = logger or self._setup_default_logger()
        self.config = self.load_config()

    # ------------------------------------------------------------------
    # Setup helpers
    # ------------------------------------------------------------------

    def _setup_default_logger(self) -> logging.Logger:
        logger = logging.getLogger('IllustrationsFormatter')
        logger.setLevel(logging.INFO)
        handler = logging.StreamHandler()
        handler.setLevel(logging.INFO)
        handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S',
        ))
        logger.addHandler(handler)
        return logger

    def load_config(self) -> Dict[str, Any]:
        """Load configuration from the JSON file, merging with built-in defaults."""
        config_path = Path(__file__).parent / self.CONFIG_FILE
        default_config = {
            'source_folder': '',
            'destination_folder': '',
            'destination_folder_instagram': (
                r'C:\Users\rober\iCloudDrive'
                r'\6LVTQB9699~com~seriflabs~affinitydesigner'
                r'\Roberto\archived_IGformat'
            ),
            'destination_folder_1920x1080': (
                r'C:\Users\rober\iCloudDrive'
                r'\6LVTQB9699~com~seriflabs~affinitydesigner'
                r'\Roberto\archived_1920x1080'
            ),
            'aspect_ratio': '3:4',
            'background_color': '',
            'format_type': 'instagram',
            'extend_border': False,
        }
        try:
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                default_config.update(config)
                return default_config
        except (json.JSONDecodeError, IOError) as e:
            self.logger.warning(f"Could not load config file: {e}")
        return default_config

    # ------------------------------------------------------------------
    # Parsing helpers
    # ------------------------------------------------------------------

    def parse_aspect_ratio(self, ratio_str: str) -> float:
        """
        Parse an aspect-ratio string and return its float value.

        Args:
            ratio_str: Ratio as a named string (``'4:5'``, ``'9:16'``),
                       a plain float string (``'0.8'``), or a
                       ``'width:height'`` expression.

        Returns:
            ``width / height`` as a float.

        Raises:
            ValueError: If the format is unrecognised.
        """
        if ratio_str in self.DEFAULT_ASPECT_RATIOS:
            return self.DEFAULT_ASPECT_RATIOS[ratio_str]
        try:
            return float(ratio_str)
        except ValueError:
            pass
        if ':' in ratio_str:
            try:
                w, h = map(float, ratio_str.split(':'))
                if h == 0:
                    raise ValueError("Height cannot be zero")
                return w / h
            except (ValueError, TypeError):
                pass
        raise ValueError(f"Invalid aspect ratio format: {ratio_str}")

    def get_background_color(self, image: Image.Image) -> Tuple[int, int, int]:
        """
        Sample the four corners of *image* and return the top-left colour.

        For images where the corner colours differ from each other (e.g.
        a coloured illustrative border), prefer ``extend_border=True`` in
        the processing methods to avoid a visible colour-band artefact.

        Args:
            image: PIL Image object.

        Returns:
            RGB tuple of the detected background colour.
        """
        width, height = image.size
        corners = [
            (0, 0),
            (width - 1, 0),
            (0, height - 1),
            (width - 1, height - 1),
        ]
        colors = []
        for x, y in corners:
            pixel = image.getpixel((x, y))
            if isinstance(pixel, tuple) and len(pixel) >= 3:
                colors.append(pixel[:3])
            else:
                colors.append((pixel, pixel, pixel))
        return colors[0]

    # ------------------------------------------------------------------
    # Internal pixel-extension helper
    # ------------------------------------------------------------------

    @staticmethod
    def _extend_edges(
        canvas: Image.Image,
        img_x: int,
        img_y: int,
        img_w: int,
        img_h: int,
    ) -> None:
        """
        Fill the padding strips of *canvas* by replicating the outermost
        pixel column / row of the already-pasted image region.

        This avoids the flat-colour band that appears when the image border
        colour does not exactly match the corner-sampled background colour.
        The canvas is modified **in-place**.

        Processing order: left → right → top → bottom.  In the rare case
        where padding exists in all four directions the corner rectangles
        will inherit the flat background colour, which is acceptable.

        Args:
            canvas: Destination RGB canvas (source image already pasted).
            img_x:  Left edge (x) of the pasted image on the canvas.
            img_y:  Top edge (y) of the pasted image on the canvas.
            img_w:  Pixel width of the pasted image region.
            img_h:  Pixel height of the pasted image region.
        """
        cw, ch = canvas.size

        # Left strip
        if img_x > 0:
            left_col = canvas.crop((img_x, 0, img_x + 1, ch))
            canvas.paste(
                left_col.resize((img_x, ch), Image.Resampling.NEAREST),
                (0, 0),
            )

        # Right strip
        right_start = img_x + img_w
        right_pad = cw - right_start
        if right_pad > 0:
            right_col = canvas.crop((right_start - 1, 0, right_start, ch))
            canvas.paste(
                right_col.resize((right_pad, ch), Image.Resampling.NEAREST),
                (right_start, 0),
            )

        # Top strip
        if img_y > 0:
            top_row = canvas.crop((0, img_y, cw, img_y + 1))
            canvas.paste(
                top_row.resize((cw, img_y), Image.Resampling.NEAREST),
                (0, 0),
            )

        # Bottom strip
        bottom_start = img_y + img_h
        bottom_pad = ch - bottom_start
        if bottom_pad > 0:
            bottom_row = canvas.crop((0, bottom_start - 1, cw, bottom_start))
            canvas.paste(
                bottom_row.resize((cw, bottom_pad), Image.Resampling.NEAREST),
                (0, bottom_start),
            )

    # ------------------------------------------------------------------
    # Public image-processing methods
    # ------------------------------------------------------------------

    def process_image(
        self,
        input_path: Path,
        output_path: Path,
        aspect_ratio: float,
        background_color: Optional[Tuple[int, int, int]] = None,
        extend_border: bool = False,
    ) -> None:
        """
        Pad a single image to the target aspect ratio.

        The original image is centred on a new canvas whose size is the
        minimum required to reach *aspect_ratio*.  The padding area is
        filled with *background_color* (auto-detected when ``None``).
        When *extend_border* is ``True`` the flat fill is replaced by
        replicating the outermost edge pixels — useful when the image has
        a coloured border that differs from the detected corner colour.

        Args:
            input_path:       Path to the source image.
            output_path:      Destination path for the processed image.
            aspect_ratio:     Target aspect ratio as ``width / height``.
            background_color: RGB padding colour.  Auto-detected from corners
                              when ``None``.
            extend_border:    Replicate edge pixels instead of a flat fill.

        Raises:
            Exception: If image processing fails.
        """
        self.logger.info(f"Processing: {input_path.name}")

        with Image.open(input_path) as img:
            if img.mode not in ('RGB', 'RGBA'):
                img = img.convert('RGB')

            orig_w, orig_h = img.size
            current_ratio = orig_w / orig_h

            if abs(current_ratio - aspect_ratio) < 0.01:
                img.save(output_path, 'PNG')
                self.logger.info(f"Image already in target ratio: {input_path.name}")
                return

            if background_color is None:
                background_color = self.get_background_color(img)

            if current_ratio > aspect_ratio:
                new_w = orig_w
                new_h = int(orig_w / aspect_ratio)
            else:
                new_h = orig_h
                new_w = int(orig_h * aspect_ratio)

            canvas = Image.new('RGB', (new_w, new_h), background_color)
            x_off = (new_w - orig_w) // 2
            y_off = (new_h - orig_h) // 2

            if img.mode == 'RGBA':
                canvas.paste(img, (x_off, y_off), img)
            else:
                canvas.paste(img, (x_off, y_off))

            if extend_border:
                self._extend_edges(canvas, x_off, y_off, orig_w, orig_h)

            output_path.parent.mkdir(parents=True, exist_ok=True)
            canvas.save(output_path, 'PNG')
            self.logger.info(
                f"Saved: {output_path.name} "
                f"({orig_w}x{orig_h} → {new_w}x{new_h})"
            )

    def process_image_fixed_size(
        self,
        input_path: Path,
        output_path: Path,
        target_width: int,
        target_height: int,
        background_color: Optional[Tuple[int, int, int]] = None,
        extend_border: bool = False,
    ) -> None:
        """
        Scale and pad a single image to exact pixel dimensions.

        The image is downscaled (preserving aspect ratio, using
        ``LANCZOS`` resampling) so it fits within the target canvas, then
        centred.  Remaining strips are filled with *background_color* or,
        when *extend_border* is ``True``, by replicating the outermost
        edge pixels — eliminating the colour-band artefact visible when the
        image border has a different hue than the detected corner colour.

        Args:
            input_path:       Path to the source image.
            output_path:      Destination path for the processed image.
            target_width:     Canvas width in pixels (e.g. ``1920``).
            target_height:    Canvas height in pixels (e.g. ``1080``).
            background_color: RGB padding colour.  Auto-detected from corners
                              when ``None``.
            extend_border:    Replicate edge pixels instead of a flat fill.

        Raises:
            Exception: If image processing fails.
        """
        self.logger.info(f"Processing: {input_path.name}")

        with Image.open(input_path) as img:
            if img.mode not in ('RGB', 'RGBA'):
                img = img.convert('RGB')

            orig_w, orig_h = img.size

            if orig_w == target_width and orig_h == target_height:
                img.save(output_path, 'PNG')
                self.logger.info(f"Image already in target size: {input_path.name}")
                return

            if background_color is None:
                background_color = self.get_background_color(img)

            scale = min(target_width / orig_w, target_height / orig_h)
            scaled_w = int(orig_w * scale)
            scaled_h = int(orig_h * scale)

            resized = img.resize((scaled_w, scaled_h), Image.Resampling.LANCZOS)

            canvas = Image.new('RGB', (target_width, target_height), background_color)
            x_off = (target_width - scaled_w) // 2
            y_off = (target_height - scaled_h) // 2

            if resized.mode == 'RGBA':
                canvas.paste(resized, (x_off, y_off), resized)
            else:
                canvas.paste(resized, (x_off, y_off))

            if extend_border:
                self._extend_edges(canvas, x_off, y_off, scaled_w, scaled_h)

            output_path.parent.mkdir(parents=True, exist_ok=True)
            canvas.save(output_path, 'PNG')
            self.logger.info(
                f"Saved: {output_path.name} "
                f"({orig_w}x{orig_h} → {target_width}x{target_height})"
            )

    def convert_single_to_1920x1080(
        self,
        input_path: Path,
        output_path: Path,
        background_color: Optional[Tuple[int, int, int]] = None,
        extend_border: bool = False,
    ) -> None:
        """
        Convert a *single* image to 1920 × 1080 pixels.

        Convenience wrapper around :meth:`process_image_fixed_size` with the
        target dimensions preset to 1920 × 1080.  Designed for quickly
        exporting a square illustration to the standard HD widescreen format
        without having to set up a whole folder batch.

        The image is scaled to fit within 1920 × 1080 (preserving aspect
        ratio) and centred.  For a 1:1 square source the result will have
        420 px padding strips on the left and right.

        Args:
            input_path:       Path to the source image (typically 1:1 square).
            output_path:      Destination path for the 1920 × 1080 image.
            background_color: RGB colour for the side padding.  Auto-detected
                              from corners when ``None``.
            extend_border:    When ``True``, fill the padding strips by
                              replicating the outermost edge pixels rather than
                              using a flat colour.  Strongly recommended for
                              illustrations whose border colour differs from
                              the detected background (e.g. a coloured ocean
                              edge on an iceberg graphic).

        Raises:
            FileNotFoundError: If *input_path* does not exist.
            Exception: If image processing fails.
        """
        if not input_path.exists():
            raise FileNotFoundError(f"Input image not found: {input_path}")
        self.process_image_fixed_size(
            input_path,
            output_path,
            target_width=1920,
            target_height=1080,
            background_color=background_color,
            extend_border=extend_border,
        )

    # ------------------------------------------------------------------
    # Batch-folder methods
    # ------------------------------------------------------------------

    def process_folder(
        self,
        source_folder: str,
        destination_folder: str,
        aspect_ratio: str,
        background_color: Optional[Tuple[int, int, int]] = None,
        extend_border: bool = False,
        progress_callback=None,
    ) -> ProcessingResult:
        """
        Process all images in a folder to the target aspect ratio.

        Images whose destination copy already has the correct ratio are
        skipped.

        Args:
            source_folder:      Path to the folder containing source images.
            destination_folder: Path where processed images are saved.
            aspect_ratio:       Target aspect ratio string (e.g. ``'4:5'``).
            background_color:   Optional RGB tuple for padding colour.
            extend_border:      Replicate edge pixels instead of a flat fill.
            progress_callback:  Optional ``fn(current, total, message)``
                                called after each image.

        Returns:
            :class:`ProcessingResult` with counts and any error details.
        """
        start_time = time.time()
        errors: List[Tuple[str, str]] = []

        source_path = Path(source_folder)
        dest_path = Path(destination_folder)

        if not source_path.exists():
            raise ValueError(f"Source folder does not exist: {source_folder}")
        if not source_path.is_dir():
            raise ValueError(f"Source path is not a directory: {source_folder}")

        try:
            ratio_float = self.parse_aspect_ratio(aspect_ratio)
        except ValueError as e:
            raise ValueError(f"Invalid aspect ratio: {e}")

        dest_path.mkdir(parents=True, exist_ok=True)

        image_files: List[Path] = []
        for ext in self.SUPPORTED_FORMATS:
            image_files.extend(source_path.glob(f'*{ext}'))
            image_files.extend(source_path.glob(f'*{ext.upper()}'))
        image_files = list(set(image_files))
        total = len(image_files)

        if total == 0:
            self.logger.warning(f"No supported image files found in: {source_folder}")
            return ProcessingResult(0, 0, 0, 0, 0.0, [])

        self.logger.info(f"Found {total} image(s) to process")
        successful = failed = skipped = 0

        for idx, img_path in enumerate(image_files, 1):
            try:
                out = dest_path / img_path.name
                if out.exists():
                    try:
                        with Image.open(out) as ex:
                            if abs(ex.width / ex.height - ratio_float) < 0.01:
                                self.logger.info(f"Skipping already processed: {out.name}")
                                skipped += 1
                                if progress_callback:
                                    progress_callback(idx, total, f"Skipping: {img_path.name}")
                                continue
                    except Exception as e:
                        self.logger.warning(
                            f"Could not verify {out.name}, re-processing. Error: {e}"
                        )
                self.process_image(
                    img_path, out, ratio_float, background_color, extend_border
                )
                successful += 1
                if progress_callback:
                    progress_callback(idx, total, f"Processing: {img_path.name}")
            except Exception as e:
                failed += 1
                self.logger.error(f"Failed to process {img_path.name}: {e}")
                errors.append((img_path.name, str(e)))
                if progress_callback:
                    progress_callback(idx, total, f"Error: {img_path.name}")

        elapsed = time.time() - start_time
        self.logger.info(
            f"\nProcessing complete! "
            f"Total: {total}  Successful: {successful}  "
            f"Skipped: {skipped}  Failed: {failed}  "
            f"Time: {elapsed:.2f}s"
        )
        return ProcessingResult(total, successful, failed, skipped, elapsed, errors)

    def process_folder_fixed_size(
        self,
        source_folder: str,
        destination_folder: str,
        target_width: int,
        target_height: int,
        background_color: Optional[Tuple[int, int, int]] = None,
        extend_border: bool = False,
        progress_callback=None,
    ) -> ProcessingResult:
        """
        Process all images in a folder to exact pixel dimensions.

        Images whose destination copy already matches the target size are
        skipped.

        Args:
            source_folder:      Path to the folder containing source images.
            destination_folder: Path where processed images are saved.
            target_width:       Canvas width in pixels.
            target_height:      Canvas height in pixels.
            background_color:   Optional RGB tuple for padding colour.
            extend_border:      Replicate edge pixels instead of a flat fill.
            progress_callback:  Optional ``fn(current, total, message)``
                                called after each image.

        Returns:
            :class:`ProcessingResult` with counts and any error details.
        """
        start_time = time.time()
        errors: List[Tuple[str, str]] = []

        source_path = Path(source_folder)
        dest_path = Path(destination_folder)

        if not source_path.exists():
            raise ValueError(f"Source folder does not exist: {source_folder}")
        if not source_path.is_dir():
            raise ValueError(f"Source path is not a directory: {source_folder}")

        dest_path.mkdir(parents=True, exist_ok=True)

        image_files: List[Path] = []
        for ext in self.SUPPORTED_FORMATS:
            image_files.extend(source_path.glob(f'*{ext}'))
            image_files.extend(source_path.glob(f'*{ext.upper()}'))
        image_files = list(set(image_files))
        total = len(image_files)

        if total == 0:
            self.logger.warning(f"No supported image files found in: {source_folder}")
            return ProcessingResult(0, 0, 0, 0, 0.0, [])

        self.logger.info(f"Found {total} image(s) to process")
        successful = failed = skipped = 0

        for idx, img_path in enumerate(image_files, 1):
            try:
                out = dest_path / img_path.name
                if out.exists():
                    try:
                        with Image.open(out) as ex:
                            if ex.width == target_width and ex.height == target_height:
                                self.logger.info(f"Skipping already processed: {out.name}")
                                skipped += 1
                                if progress_callback:
                                    progress_callback(idx, total, f"Skipping: {img_path.name}")
                                continue
                    except Exception as e:
                        self.logger.warning(
                            f"Could not verify {out.name}, re-processing. Error: {e}"
                        )
                self.process_image_fixed_size(
                    img_path, out,
                    target_width, target_height,
                    background_color, extend_border,
                )
                successful += 1
                if progress_callback:
                    progress_callback(idx, total, f"Processing: {img_path.name}")
            except Exception as e:
                failed += 1
                self.logger.error(f"Failed to process {img_path.name}: {e}")
                errors.append((img_path.name, str(e)))
                if progress_callback:
                    progress_callback(idx, total, f"Error: {img_path.name}")

        elapsed = time.time() - start_time
        self.logger.info(
            f"\nProcessing complete! "
            f"Total: {total}  Successful: {successful}  "
            f"Skipped: {skipped}  Failed: {failed}  "
            f"Time: {elapsed:.2f}s"
        )
        return ProcessingResult(total, successful, failed, skipped, elapsed, errors)


# ----------------------------------------------------------------------
# Utility
# ----------------------------------------------------------------------

def parse_color(color_str: str) -> Tuple[int, int, int]:
    """
    Parse a colour string to an ``(R, G, B)`` tuple.

    Accepts:
    * Hex strings with or without a leading ``#``: ``#F5A623``, ``f5a623``
    * Comma-separated integers: ``245,166,35``

    Raises:
        ValueError: If the format is not recognised.
    """
    if not color_str:
        return None
    color_str = color_str.strip().lstrip('#').lower()
    if len(color_str) == 6:
        try:
            r = int(color_str[0:2], 16)
            g = int(color_str[2:4], 16)
            b = int(color_str[4:6], 16)
            return (r, g, b)
        except ValueError:
            pass
    if ',' in color_str:
        try:
            parts = [int(x.strip()) for x in color_str.split(',')]
            if len(parts) == 3 and all(0 <= x <= 255 for x in parts):
                return tuple(parts)
        except ValueError:
            pass
    raise ValueError(f"Invalid color format: {color_str}")


# ----------------------------------------------------------------------
# CLI entry point
# ----------------------------------------------------------------------

def main():
    """Command-line interface for batch image formatting."""
    formatter = IllustrationsFormatter()
    config = formatter.config

    parser = argparse.ArgumentParser(
        description='Format images for Instagram or fixed dimensions (1920×1080)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s -s ./images -d ./output -r 4:5
  %(prog)s -s ./photos -d ./formatted -r 9:16 -c 255,255,255
  %(prog)s -s ./input  -d ./output  -r 0.8 -c "#FF5733"
  %(prog)s -s ./input  -d ./output  -r 3:4 --extend-border
        """,
    )
    parser.add_argument('-s', '--source',
                        default=config.get('source_folder', ''),
                        help='Source folder (default from config)')
    parser.add_argument('-d', '--destination',
                        default=config.get('destination_folder', ''),
                        help='Destination folder (default from config)')
    parser.add_argument('-r', '--ratio',
                        default=config.get('aspect_ratio', '3:4'),
                        help='Target aspect ratio, e.g. 4:5, 9:16, 0.8 (default from config)')
    parser.add_argument('-c', '--color',
                        default=config.get('background_color', ''),
                        help='Background colour as #RRGGBB or R,G,B (default from config)')
    parser.add_argument('--extend-border', action='store_true',
                        help=(
                            'Fill padding by replicating the outermost edge pixels '
                            'instead of a flat colour.  Useful when the image border '
                            'colour differs from the detected corner colour.'
                        ))
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='Enable verbose logging')

    args = parser.parse_args()

    if not args.source:
        print("Error: Source folder must be specified via -s or in config file",
              file=sys.stderr)
        sys.exit(1)
    if not args.destination:
        print("Error: Destination folder must be specified via -d or in config file",
              file=sys.stderr)
        sys.exit(1)

    if args.verbose:
        formatter.logger.setLevel(logging.DEBUG)

    bg_color = None
    if args.color:
        try:
            bg_color = parse_color(args.color)
        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)

    try:
        result = formatter.process_folder(
            args.source,
            args.destination,
            args.ratio,
            bg_color,
            extend_border=args.extend_border,
        )
        print(f"\nProcessing complete!")
        print(f"Total images: {result.total_images}")
        print(f"Successful:   {result.successful}")
        print(f"Skipped:      {result.skipped}")
        print(f"Failed:       {result.failed}")
        print(f"Time elapsed: {result.elapsed_time:.2f} seconds")
        if result.errors:
            print("\nErrors:")
            for filename, error in result.errors:
                print(f"  - {filename}: {error}")
        sys.exit(1 if result.failed > 0 else 0)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
