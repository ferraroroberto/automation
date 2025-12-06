import os
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageChops, ImageStat

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def unpack_gif() -> None:
    """
    Selects a GIF file via dialog and extracts all frames as high-quality JPEGs.
    Files are saved in a new folder named after the GIF file.
    Naming convention: [original_name]_[xxx]_of_[yyy].jpg
    """
    # Hide the main tkinter window
    root = tk.Tk()
    root.withdraw()

    # Ask for the GIF file
    file_path = filedialog.askopenfilename(
        title="Select GIF to unpack",
        filetypes=[("GIF files", "*.gif")]
    )

    if not file_path:
        logger.info("No file selected. Exiting.")
        return

    try:
        base_dir = os.path.dirname(file_path)
        filename = os.path.basename(file_path)
        name_no_ext = os.path.splitext(filename)[0]
        
        # Create a specific directory for the frames to avoid clutter
        output_dir = os.path.join(base_dir, f"{name_no_ext}_frames")
        os.makedirs(output_dir, exist_ok=True)

        with Image.open(file_path) as im:
            n_frames = getattr(im, 'n_frames', 1)
            
            logger.info(f"Extracting frames from '{filename}' (Total: {n_frames})...")
            
            previous_frame = None
            saved_count = 0

            for i in range(n_frames):
                im.seek(i)
                # Convert to RGB mode (required for JPEG)
                frame = im.convert('RGB')
                
                # Check for duplicate frames
                if previous_frame is not None:
                    # Use Mean Squared Error or simpler mean difference to handle slight compression artifacts
                    diff = ImageChops.difference(frame, previous_frame)
                    stat = ImageStat.Stat(diff)
                    # average mean of all bands (R, G, B)
                    diff_mean = sum(stat.mean) / len(stat.mean)
                    
                    # Threshold: if average pixel difference is less than 1.0 (out of 255), consider them identical
                    # strict equality (getbbox() is None) often fails due to GIF compression noise
                    if diff_mean < 1.0:
                        logger.info(f"Frame {i+1}/{n_frames} is visually identical to previous frame (diff: {diff_mean:.4f}). Skipping.")
                        continue

                # Format: original_name_xxx_of_yyy.jpg
                # xxx is 1-based index padded to 3 digits
                # yyy is total frames padded to 3 digits
                new_filename = f"{name_no_ext}_{i+1:03d}_of_{n_frames:03d}.jpg"
                save_path = os.path.join(output_dir, new_filename)
                
                # Save as high-quality JPEG
                # quality=100: Max quality
                # subsampling=0: 4:4:4 (no chroma subsampling) for best sharpness
                frame.save(save_path, "JPEG", quality=100, subsampling=0)
                
                previous_frame = frame
                saved_count += 1

        logger.info(f"Successfully extracted {saved_count} unique frames (scanned {n_frames}) to {output_dir}")
        messagebox.showinfo(
            "Success", 
            f"Successfully extracted {saved_count} unique frames (scanned {n_frames}).\nSaved in:\n{os.path.normpath(output_dir)}"
        )

    except Exception as e:
        logger.error(f"An error occurred: {e}")
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
    finally:
        root.destroy()

if __name__ == "__main__":
    unpack_gif()

