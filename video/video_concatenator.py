import os
import subprocess
from typing import List
import logging
import tkinter as tk
from tkinter import filedialog

# Configure logging following AGENTS.md standards
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class VideoConcatenator:
    """Class to handle video concatenation operations using ffmpeg."""
    
    def concatenate_videos(self, video_paths: List[str], output_path: str) -> bool:
        """
        Concatenate multiple video files into a single output video using ffmpeg.
        
        Args:
            video_paths: List of paths to video files to concatenate
            output_path: Path where the output video will be saved
        
        Returns:
            bool: True if successful, False otherwise
        """
        # Validate input files
        for path in video_paths:
            if not os.path.exists(path):
                logger.error(f"Video file not found: {path}")
                return False
        
        try:
            # Create a temporary file list for ffmpeg
            list_file = "temp_video_list.txt"
            
            # Sort video_paths alphabetically before writing
            sorted_video_paths = sorted(video_paths)

            # Write the list of videos to a text file
            with open(list_file, 'w') as f:
                for path in sorted_video_paths:
                    # Convert to absolute path and escape for ffmpeg
                    abs_path = os.path.abspath(path).replace('\\', '/')
                    f.write(f"file '{abs_path}'\n")
            
            logger.info(f"Concatenating {len(video_paths)} videos with ffmpeg...")
            
            # Updated ffmpeg command to preserve tracks without re-encoding
            cmd = [
                'ffmpeg',
                '-f', 'concat',
                '-safe', '0',
                '-i', list_file,
                '-c', 'copy',       # Copy streams without re-encoding
                '-map', '0',        # Map all streams (video, audio, subtitles, etc.)
                '-y',               # Overwrite output file
                output_path
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            # Clean up temporary file
            if os.path.exists(list_file):
                os.remove(list_file)
            
            if result.returncode == 0:
                logger.info(f"Successfully created video: {output_path}")
                return True
            else:
                logger.error(f"FFmpeg error: {result.stderr}")
                # Try alternative method if the first fails
                return self._concatenate_with_filter_complex(sorted_video_paths, output_path)
                
        except FileNotFoundError:
            logger.error("FFmpeg not found. Please install ffmpeg and add it to your PATH.")
            return False
        except Exception as e:
            logger.error(f"Error concatenating videos: {str(e)}")
            # Clean up temporary file
            if os.path.exists("temp_video_list.txt"):
                os.remove("temp_video_list.txt")
            return False
    
    def _concatenate_with_filter_complex(self, video_paths: List[str], output_path: str) -> bool:
        """
        Alternative concatenation method using filter_complex - only used as last resort.
        """
        try:
            logger.info("Trying alternative concatenation method (this will re-encode)...")
            
            # Build filter_complex command
            inputs = []
            filter_parts = []
            
            for i, path in enumerate(video_paths):
                inputs.extend(['-i', path])
                filter_parts.append(f'[{i}:v][{i}:a]')
            
            filter_complex = ''.join(filter_parts) + f'concat=n={len(video_paths)}:v=1:a=1[outv][outa]'
            
            cmd = [
                'ffmpeg',
                *inputs,
                '-filter_complex', filter_complex,
                '-map', '[outv]',
                '-map', '[outa]',
                '-c:v', 'libx264',
                '-c:a', 'aac',
                '-y',
                output_path
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                logger.info(f"Successfully created video with filter_complex: {output_path}")
                return True
            else:
                logger.error(f"Filter_complex method also failed: {result.stderr}")
                return False
                
        except Exception as e:
            logger.error(f"Error with filter_complex method: {str(e)}")
            return False

def select_video_files() -> List[str]:
    """
    Open a file dialog to select multiple video files.
    
    Returns:
        List[str]: List of selected video file paths
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    
    filetypes = [
        ('Video files', '*.mp4 *.avi *.mov *.mkv *.wmv *.flv *.webm'),
        ('MP4 files', '*.mp4'),
        ('AVI files', '*.avi'),
        ('MOV files', '*.mov'),
        ('All files', '*.*')
    ]
    
    file_paths = filedialog.askopenfilenames(
        title="Select video files to concatenate",
        filetypes=filetypes
    )
    
    root.destroy()
    return list(file_paths)

def select_output_file() -> str:
    """
    Open a file dialog to select output file location.
    
    Returns:
        str: Selected output file path
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    
    filetypes = [
        ('MP4 files', '*.mp4'),
        ('AVI files', '*.avi'),
        ('MOV files', '*.mov'),
        ('All files', '*.*')
    ]
    
    file_path = filedialog.asksaveasfilename(
        title="Save concatenated video as",
        defaultextension=".mp4",
        filetypes=filetypes
    )
    
    root.destroy()
    return file_path

def main():
    """Example usage of the VideoConcatenator class with GUI file selection."""
    logger.info("🎬 Video Concatenator - Select files using GUI")
    
    try:
        # Select input video files
        video_files = select_video_files()
        
        if not video_files:
            logger.warning("⚠️ No video files selected. Exiting.")
            return

        # Log video files in the order selected by the user
        logger.info(f"📋 Selected {len(video_files)} video files:")
        for i, file in enumerate(video_files, 1):
            logger.info(f"  {i}. {os.path.basename(file)}")
        
        # Select output file location
        output_file = select_output_file()
        
        if not output_file:
            logger.warning("⚠️ No output file selected. Exiting.")
            return

        logger.info(f"📁 Output file: {output_file}")
        
        # Concatenate videos
        concatenator = VideoConcatenator()
        success = concatenator.concatenate_videos(video_files, output_file)
        
        if success:
            logger.info("✅ Videos concatenated successfully!")
            logger.info(f"📁 Output saved to: {output_file}")
        else:
            logger.error("❌ Failed to concatenate videos.")
            logger.error("💡 Check the console output above for detailed error information.")
    
    except Exception as e:
        logger.error(f"💥 Unexpected error in main: {str(e)}")

    finally:
        logger.info("🏁 Script execution completed.")

if __name__ == "__main__":
    main()