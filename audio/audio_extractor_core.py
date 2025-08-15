import os
import subprocess
import json
import logging
from pathlib import Path
from typing import List, Dict, Optional
import tkinter as tk
from tkinter import filedialog, messagebox

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class AudioExtractor:
    """Extracts audio tracks from video files using FFmpeg."""
    
    def __init__(self, input_directory: str = None):
        """Initialize the audio extractor.
        
        Args:
            input_directory: Directory containing video files. Defaults to current directory.
        """
        self.input_directory = Path(input_directory) if input_directory else Path.cwd()
        self.supported_formats = {'.mkv', '.mp4'}
        self.supported_extensions = [('Video files', '*.mkv *.mp4'), ('All files', '*.*')]
        
    def check_ffmpeg(self) -> bool:
        """Check if FFmpeg and FFprobe are available."""
        try:
            subprocess.run(['ffmpeg', '-version'], capture_output=True, check=True)
            subprocess.run(['ffprobe', '-version'], capture_output=True, check=True)
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            logger.error("FFmpeg or FFprobe not found. Please install FFmpeg.")
            return False
    
    def get_audio_tracks(self, video_file: Path) -> List[Dict]:
        """Get information about audio tracks in a video file.
        
        Args:
            video_file: Path to the video file.
            
        Returns:
            List of dictionaries containing audio track information.
        """
        try:
            cmd = [
                'ffprobe',
                '-v', 'error',
                '-select_streams', 'a',
                '-show_entries', 'stream=index,codec_name,codec_type,tags',
                '-of', 'json',
                str(video_file)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            data = json.loads(result.stdout)
            
            audio_tracks = []
            for stream in data.get('streams', []):
                if stream.get('codec_type') == 'audio':
                    track_info = {
                        'index': stream['index'],
                        'codec': stream.get('codec_name', 'unknown'),
                        'language': stream.get('tags', {}).get('language', 'und'),
                        'title': stream.get('tags', {}).get('title', '')
                    }
                    audio_tracks.append(track_info)
            
            return audio_tracks
            
        except subprocess.CalledProcessError as e:
            logger.error(f"Error getting audio tracks from {video_file.name}: {e}")
            return []
        except json.JSONDecodeError as e:
            logger.error(f"Error parsing FFprobe output for {video_file.name}: {e}")
            return []
    
    def extract_audio_track(self, video_file: Path, track_index: int, output_file: Path) -> bool:
        """Extract a single audio track from a video file.
        
        Args:
            video_file: Path to the video file.
            track_index: Index of the audio track to extract.
            output_file: Path for the output audio file.
            
        Returns:
            True if extraction was successful, False otherwise.
        """
        try:
            cmd = [
                'ffmpeg',
                '-i', str(video_file),
                '-map', f'0:a:{track_index}',
                '-c:a', 'libmp3lame',
                '-b:a', '192k',
                '-y',
                str(output_file.with_suffix('.mp3'))
            ]
            
            subprocess.run(cmd, capture_output=True, check=True)
            logger.info(f"Extracted track {track_index} to {output_file.with_suffix('.mp3').name}")
            return True
            
        except subprocess.CalledProcessError as e:
            logger.error(f"Error extracting track {track_index} from {video_file.name}: {e}")
            return False
    
    def extract_all_audio_unified(self, video_file: Path, output_file: Path) -> bool:
        """Extract all audio tracks into a single unified file by mixing them together.
        
        Args:
            video_file: Path to the video file.
            output_file: Path for the output audio file.
            
        Returns:
            True if extraction was successful, False otherwise.
        """
        try:
            audio_tracks = self.get_audio_tracks(video_file)
            num_tracks = len(audio_tracks)
            
            if num_tracks == 1:
                # If only one track, just extract it
                cmd = [
                    'ffmpeg',
                    '-i', str(video_file),
                    '-map', '0:a:0',
                    '-c:a', 'libmp3lame',
                    '-b:a', '192k',
                    '-y',
                    str(output_file.with_suffix('.mp3'))
                ]
            else:
                # Mix multiple tracks together
                inputs = ' '.join([f'[0:a:{i}]' for i in range(num_tracks)])
                filter_complex = f'{inputs}amix=inputs={num_tracks}:duration=longest[out]'
                
                cmd = [
                    'ffmpeg',
                    '-i', str(video_file),
                    '-filter_complex', filter_complex,
                    '-map', '[out]',
                    '-c:a', 'libmp3lame',
                    '-b:a', '192k',
                    '-y',
                    str(output_file.with_suffix('.mp3'))
                ]
            
            subprocess.run(cmd, capture_output=True, check=True)
            logger.info(f"Extracted and mixed all audio tracks to {output_file.with_suffix('.mp3').name}")
            return True
            
        except subprocess.CalledProcessError as e:
            logger.error(f"Error extracting unified audio from {video_file.name}: {e}")
            return False
    
    def process_video_file(self, video_file: Path) -> None:
        """Process a single video file and extract all audio tracks.
        
        Args:
            video_file: Path to the video file.
        """
        logger.info(f"Processing {video_file.name}")
        
        audio_tracks = self.get_audio_tracks(video_file)
        
        if not audio_tracks:
            logger.warning(f"No audio tracks found in {video_file.name}")
            return
        
        base_name = video_file.stem
        output_dir = video_file.parent
        
        # Extract unified audio (mixed tracks)
        unified_output = output_dir / f"{base_name}_audio_unified"
        self.extract_all_audio_unified(video_file, unified_output)
        
        # Extract individual tracks if there are multiple
        if len(audio_tracks) > 1:
            for i, track in enumerate(audio_tracks):
                track_output = output_dir / f"{base_name}_audio_track{i + 1}"
                self.extract_audio_track(video_file, i, track_output)
        
        logger.info(f"Completed processing {video_file.name}")
    
    def process_directory(self) -> None:
        """Process all video files in the specified directory."""
        if not self.check_ffmpeg():
            return
        
        video_files = []
        for ext in self.supported_formats:
            video_files.extend(self.input_directory.glob(f'*{ext}'))
        
        if not video_files:
            logger.warning(f"No video files found in {self.input_directory}")
            return
        
        logger.info(f"Found {len(video_files)} video file(s) to process")
        
        for video_file in video_files:
            try:
                self.process_video_file(video_file)
            except Exception as e:
                logger.error(f"Unexpected error processing {video_file.name}: {e}")
                continue
        
        logger.info("Processing complete")
    
    def select_files_gui(self) -> List[Path]:
        """Open a GUI dialog to select video files.
        
        Returns:
            List of selected video file paths.
        """
        root = tk.Tk()
        root.withdraw()
        
        file_paths = filedialog.askopenfilenames(
            title="Select video files to extract audio",
            filetypes=self.supported_extensions,
            initialdir=self.input_directory
        )
        
        root.destroy()
        
        selected_files = []
        for file_path in file_paths:
            path = Path(file_path)
            if path.suffix.lower() in self.supported_formats:
                selected_files.append(path)
            else:
                logger.warning(f"Skipping unsupported file: {path}")
        
        return selected_files
    
    def process_selected_files(self, video_files: List[Path]) -> None:
        """Process a list of selected video files.
        
        Args:
            video_files: List of video file paths to process.
        """
        if not self.check_ffmpeg():
            return
        
        if not video_files:
            logger.warning("No video files selected")
            return
        
        logger.info(f"Processing {len(video_files)} selected file(s)")
        
        for video_file in video_files:
            try:
                self.process_video_file(video_file)
            except Exception as e:
                logger.error(f"Unexpected error processing {video_file.name}: {e}")
                continue
        
        logger.info("Processing complete")
        
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Processing Complete", 
                          f"Successfully processed {len(video_files)} file(s).\n"
                          "Check the console for detailed information.")
        root.destroy()

def main():
    """Main function to run the audio extractor."""
    import argparse
    
    parser = argparse.ArgumentParser(description='Extract audio tracks from video files')
    parser.add_argument(
        'directory',
        nargs='?',
        default='.',
        help='Directory containing video files (default: current directory)'
    )
    parser.add_argument(
        '--gui', '-g',
        action='store_true',
        help='Use GUI file selector to choose files manually'
    )
    
    args = parser.parse_args()
    
    extractor = AudioExtractor(args.directory)
    
    if args.gui:
        selected_files = extractor.select_files_gui()
        if selected_files:
            extractor.process_selected_files(selected_files)
        else:
            logger.info("No files selected")
    else:
        extractor.process_directory()

if __name__ == '__main__':
    main()
