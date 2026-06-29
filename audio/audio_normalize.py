"""
Audio Normalization Tool

This script normalizes audio files to maximum volume without distortion using
EBU R128 loudness normalization. It provides detailed before/after analysis
and user control over normalization parameters.

📚 AUDIO NORMALIZATION CONCEPTS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
What is Audio Normalization?
- Normalization adjusts audio levels so the loudest parts reach a target volume
- Prevents distortion by ensuring peaks don't exceed 0 dBFS (digital full scale)
- Uses EBU R128 standard for broadcast-quality loudness normalization

Why Normalize Audio?
- Ensures consistent volume across different audio sources
- Maximizes loudness without introducing distortion
- Professional standard for audio mastering

Key Concepts:
- LUFS (Loudness Units Full Scale): Modern loudness measurement
- Integrated Loudness: Overall perceived loudness of the entire track
- True Peak: Maximum level including inter-sample peaks
- Target Level: Desired loudness level (typically -16 to -23 LUFS)

Normalization Methods:
- Peak Normalization: Simple but can cause distortion
- RMS Normalization: Based on average levels, better but still limited
- EBU R128 Loudnorm: Advanced algorithm considering human perception

EBU R128 Benefits:
- Dual-pass processing for accurate loudness measurement
- Prevents distortion by limiting true peaks
- Maintains dynamic range while maximizing loudness
- Industry standard for streaming and broadcast
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import json
import logging
from pathlib import Path
from typing import Dict, Optional, Tuple

# Configure module-level logger
logger = logging.getLogger(__name__)


# -------------------------------------------------------
# Function: choose_file
# Purpose:  Open a file dialog to select an audio file
# Returns:  The full path of the selected file
# -------------------------------------------------------
def choose_file() -> str:
    """Open file dialog for audio file selection."""
    root = tk.Tk()
    root.withdraw()  # Hide the main Tk window
    file_path = filedialog.askopenfilename(
        title="Select audio file to normalize",
        filetypes=[("Audio files", "*.mp3 *.wav *.flac *.aac *.ogg *.m4a *.wma *.aiff")]
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        exit()
    return file_path


# -------------------------------------------------------
# Function: analyze_audio_levels
# Purpose:  Analyze current audio levels using ffmpeg
# Arguments:
#   - input_path (str): path to the audio file
# Returns:  Dictionary with loudness analysis results
# -------------------------------------------------------
def analyze_audio_levels(input_path: str) -> Optional[Dict]:
    """
    Analyze audio loudness using ffmpeg's ebur128 filter.

    Returns comprehensive loudness statistics including:
    - Integrated loudness (I)
    - Loudness range (LRA)
    - True peak levels
    - Sample peak levels
    """
    try:
        logger.info("🔍 Analyzing current audio levels...")

        # FFmpeg command for loudness analysis
        ffmpeg_cmd = [
            "ffmpeg",
            "-i", input_path,
            "-filter_complex", "ebur128=peak=true",
            "-f", "null",
            "-"
        ]

        result = subprocess.run(
            ffmpeg_cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )

        if result.returncode != 0:
            logger.error(f"❌ Analysis failed: {result.stderr}")
            return None

        # Parse ebur128 output from stderr
        # The output format is:
        # Integrated loudness:
        #   I:         -21.1 LUFS
        #   Threshold: -31.1 LUFS
        #
        # Loudness range:
        #   LRA:         0.0 LU
        #   ...
        #
        # True peak:
        #   Peak:      -17.5 dBFS
        #
        # Sample peak:
        #   Peak:      -17.5 dBFS

        analysis = {}
        lines = result.stderr.split('\n')

        for i, line in enumerate(lines):
            line = line.strip()

            # Find integrated loudness (I: value)
            if line.startswith('I:'):
                parts = line.split()
                if len(parts) >= 2:
                    try:
                        analysis['integrated_loudness'] = float(parts[1])
                    except (ValueError, IndexError):
                        logger.warning(f"Could not parse integrated loudness from: {line}")

            # Find loudness range (LRA: value)
            elif line.startswith('LRA:'):
                parts = line.split()
                if len(parts) >= 2:
                    try:
                        analysis['loudness_range'] = float(parts[1])
                    except (ValueError, IndexError):
                        logger.warning(f"Could not parse LRA from: {line}")

            # Find true peak (Peak: value under "True peak:" section)
            elif line.startswith('Peak:'):
                parts = line.split()
                if len(parts) >= 2:
                    try:
                        peak_value = float(parts[1])
                        # Check context by looking at previous lines
                        if i > 0 and 'True peak:' in lines[i-1]:
                            analysis['true_peak'] = peak_value
                        elif i > 0 and 'Sample peak:' in lines[i-1]:
                            analysis['sample_peak'] = peak_value
                        # If we can't determine context, assume it's true peak if not set
                        elif 'true_peak' not in analysis:
                            analysis['true_peak'] = peak_value
                        else:
                            analysis['sample_peak'] = peak_value
                    except (ValueError, IndexError):
                        logger.warning(f"Could not parse peak from: {line}")

        if analysis:
            logger.info("✅ Audio analysis completed")
            return analysis
        else:
            logger.warning("⚠️ Could not parse analysis results")
            return None

    except Exception as e:
        logger.error(f"❌ Analysis error: {e}")
        return None


# -------------------------------------------------------
# Function: create_normalization_dialog
# Purpose:  Create dialog for user to choose normalization settings
# Arguments:
#   - analysis (Dict): current audio analysis results
# Returns:  Tuple of (target_lufs, apply_normalization)
# -------------------------------------------------------
def create_normalization_dialog(analysis: Dict) -> Tuple[Optional[float], bool]:
    """
    Create dialog for user normalization settings.

    Shows current levels and allows user to choose:
    - Target loudness level
    - Whether to proceed with normalization
    """
    dialog = tk.Toplevel()
    dialog.title("Audio Normalization Settings")
    dialog.geometry("500x450")

    # Current levels display
    ttk.Label(dialog, text="📊 CURRENT AUDIO ANALYSIS", font=("Arial", 12, "bold")).pack(pady=10)

    analysis_frame = ttk.Frame(dialog)
    analysis_frame.pack(pady=5, padx=20, fill="x")

    ttk.Label(analysis_frame, text="Integrated Loudness (LUFS):").grid(row=0, column=0, sticky="w")
    integrated_val = analysis.get('integrated_loudness', 'N/A')
    integrated_text = f"{integrated_val:.1f}" if isinstance(integrated_val, (int, float)) else str(integrated_val)
    ttk.Label(analysis_frame, text=integrated_text).grid(row=0, column=1, sticky="e")

    ttk.Label(analysis_frame, text="Loudness Range (LRA):").grid(row=1, column=0, sticky="w")
    lra_val = analysis.get('loudness_range', 'N/A')
    lra_text = f"{lra_val:.1f} LU" if isinstance(lra_val, (int, float)) else str(lra_val)
    ttk.Label(analysis_frame, text=lra_text).grid(row=1, column=1, sticky="e")

    ttk.Label(analysis_frame, text="True Peak:").grid(row=2, column=0, sticky="w")
    true_peak_val = analysis.get('true_peak', 'N/A')
    true_peak_text = f"{true_peak_val:.1f} dBFS" if isinstance(true_peak_val, (int, float)) else str(true_peak_val)
    ttk.Label(analysis_frame, text=true_peak_text).grid(row=2, column=1, sticky="e")

    ttk.Label(analysis_frame, text="Sample Peak:").grid(row=3, column=0, sticky="w")
    sample_peak_val = analysis.get('sample_peak', 'N/A')
    sample_peak_text = f"{sample_peak_val:.1f} dBFS" if isinstance(sample_peak_val, (int, float)) else str(sample_peak_val)
    ttk.Label(analysis_frame, text=sample_peak_text).grid(row=3, column=1, sticky="e")

    # Settings section
    ttk.Label(dialog, text="\n🎛️ NORMALIZATION SETTINGS", font=("Arial", 12, "bold")).pack(pady=10)

    settings_frame = ttk.Frame(dialog)
    settings_frame.pack(pady=5, padx=20, fill="x")

    ttk.Label(settings_frame, text="Target Loudness (LUFS):").grid(row=0, column=0, sticky="w")
    target_var = tk.StringVar(value="-16.0")  # Default for streaming platforms
    target_combo = ttk.Combobox(settings_frame, textvariable=target_var, values=["-23.0", "-20.0", "-16.0", "-14.0"])
    target_combo.grid(row=0, column=1, sticky="ew")

    # Educational info
    info_text = """
🎓 RECOMMENDED TARGETS:
• -23 LUFS: Broadcast TV standard
• -20 LUFS: Music streaming (Spotify, Apple Music)
• -16 LUFS: Podcasts and YouTube
• -14 LUFS: Voice recordings

💡 Higher targets = louder audio, but may reduce dynamic range
"""
    info_label = ttk.Label(dialog, text=info_text, justify="left", font=("Arial", 9))
    info_label.pack(pady=10, padx=20)

    result = {"target": None, "proceed": False}

    def on_proceed():
        try:
            result["target"] = float(target_var.get())
            result["proceed"] = True
            dialog.destroy()
        except ValueError:
            messagebox.showerror("Error", "Invalid target loudness value")

    def on_cancel():
        dialog.destroy()

    # Buttons
    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=20)

    ttk.Button(button_frame, text="🚫 Cancel", command=on_cancel).pack(side="left", padx=10)
    ttk.Button(button_frame, text="✅ Normalize Audio", command=on_proceed).pack(side="right", padx=10)

    dialog.wait_window()
    return result["target"], result["proceed"]


# -------------------------------------------------------
# Function: normalize_audio
# Purpose:  Apply EBU R128 loudness normalization to audio file
# Arguments:
#   - input_path (str): path to the original audio file
#   - target_lufs (float): target loudness in LUFS
# Returns:  Path to the normalized output file
# -------------------------------------------------------
def normalize_audio(input_path: str, target_lufs: float) -> Optional[str]:
    """
    Apply EBU R128 loudness normalization using ffmpeg.

    The loudnorm filter performs dual-pass processing:
    1. First pass: measures loudness characteristics
    2. Second pass: applies normalization with measured parameters

    This ensures accurate normalization without distortion.
    """
    input_dir = os.path.dirname(input_path)
    filename = os.path.basename(input_path)
    name, ext = os.path.splitext(filename)

    # Create output filename with normalization info
    output_file = os.path.join(input_dir, f"{name}_normalized_{target_lufs}LUFS{ext}")

    try:
        logger.info(f"🔄 Normalizing audio to {target_lufs} LUFS...")

        # Two-pass loudnorm processing
        # First pass: measure loudness
        measure_cmd = [
            "ffmpeg",
            "-i", input_path,
            "-filter_complex", f"loudnorm=I={target_lufs}:TP=-1.5:LRA=11:print_format=json",
            "-f", "null",
            "-"
        ]

        logger.debug("📏 First pass: measuring loudness characteristics")
        measure_result = subprocess.run(
            measure_cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )

        if measure_result.returncode != 0:
            logger.error(f"❌ Measurement failed: {measure_result.stderr}")
            return None

        # Extract loudnorm parameters from JSON output
        loudnorm_params = None
        json_lines = []

        # Check both stdout and stderr for JSON
        output_to_check = measure_result.stderr + "\n" + measure_result.stdout

        for line in output_to_check.split('\n'):
            line = line.strip()
            if line == '{':
                # JSON starts on this line, collect until closing brace
                json_lines = [line]
                continue
            elif json_lines and line == '}':
                json_lines.append(line)
                json_text = '\n'.join(json_lines)
                try:
                    loudnorm_params = json.loads(json_text)
                    logger.debug(f"Successfully parsed loudnorm parameters: {loudnorm_params}")
                    break
                except json.JSONDecodeError:
                    json_lines = []
                    continue
            elif json_lines:
                json_lines.append(line)

        if not loudnorm_params:
            logger.error("❌ Could not extract loudnorm parameters from ffmpeg output")
            return None

        logger.debug("📏 Second pass: applying normalization")

        # Second pass: apply normalization with measured parameters
        normalize_cmd = [
            "ffmpeg",
            "-y",  # Overwrite output
            "-i", input_path,
            "-filter_complex", (
                f"loudnorm=I={target_lufs}:TP=-1.5:LRA=11:"
                f"measured_I={loudnorm_params['input_i']}:"
                f"measured_LRA={loudnorm_params['input_lra']}:"
                f"measured_TP={loudnorm_params['input_tp']}:"
                f"measured_thresh={loudnorm_params['input_thresh']}:"
                f"offset={loudnorm_params['target_offset']}:"
                "print_format=summary"
            ),
            "-c:a", "aac",  # AAC codec for compatibility
            "-b:a", "192k",  # High quality bitrate
            "-movflags", "+faststart",
            output_file
        ]

        result = subprocess.run(normalize_cmd, check=True, capture_output=True, text=True)

        logger.info("✅ Audio normalization completed successfully")
        return output_file

    except subprocess.CalledProcessError as e:
        logger.error(f"❌ Normalization failed: {e}")
        messagebox.showerror("Error", f"Normalization failed:\n{e}")
        return None
    except Exception as e:
        logger.error(f"❌ Unexpected error during normalization: {e}")
        messagebox.showerror("Error", f"Unexpected error:\n{e}")
        return None


# -------------------------------------------------------
# Function: show_results
# Purpose:  Display before/after comparison and normalization results
# Arguments:
#   - original_analysis (Dict): analysis of original file
#   - output_path (str): path to normalized file
# -------------------------------------------------------
def show_results(original_analysis: Dict, output_path: str) -> None:
    """Display comprehensive before/after normalization results."""

    # Analyze normalized file
    normalized_analysis = analyze_audio_levels(output_path)

    if not normalized_analysis:
        messagebox.showwarning("Warning", "Could not analyze normalized file, but it was created successfully.")
        return

    # Create results dialog
    results_dialog = tk.Toplevel()
    results_dialog.title("Normalization Results")
    results_dialog.geometry("600x500")

    # Title
    ttk.Label(results_dialog, text="🎵 AUDIO NORMALIZATION RESULTS",
              font=("Arial", 14, "bold")).pack(pady=15)

    # Create comparison table
    comparison_frame = ttk.Frame(results_dialog)
    comparison_frame.pack(pady=10, padx=20, fill="x")

    # Headers
    ttk.Label(comparison_frame, text="", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5)
    ttk.Label(comparison_frame, text="BEFORE", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5)
    ttk.Label(comparison_frame, text="AFTER", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5)
    ttk.Label(comparison_frame, text="CHANGE", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5)

    # Data rows
    metrics = [
        ("Integrated Loudness", "integrated_loudness", "LUFS"),
        ("Loudness Range", "loudness_range", "LU"),
        ("True Peak", "true_peak", "dBFS"),
        ("Sample Peak", "sample_peak", "dBFS")
    ]

    for i, (label, key, unit) in enumerate(metrics, 1):
        ttk.Label(comparison_frame, text=f"{label} ({unit}):").grid(row=i, column=0, sticky="w", padx=5, pady=2)

        before_val = original_analysis.get(key, 0)
        after_val = normalized_analysis.get(key, 0) if normalized_analysis else 0
        change = after_val - before_val

        # Format values safely
        before_text = f"{before_val:.1f}" if isinstance(before_val, (int, float)) else str(before_val)
        after_text = f"{after_val:.1f}" if isinstance(after_val, (int, float)) else str(after_val)
        change_text = f"{change:+.1f}" if isinstance(change, (int, float)) else str(change)

        ttk.Label(comparison_frame, text=before_text).grid(row=i, column=1, padx=5)
        ttk.Label(comparison_frame, text=after_text).grid(row=i, column=2, padx=5)

        # Color-code change (green for improvement, red for degradation)
        change_color = "green" if abs(change) < 0.5 else "blue"
        ttk.Label(comparison_frame, text=change_text, foreground=change_color).grid(row=i, column=3, padx=5)

    # Summary text
    summary_text = f"""
📁 Output File: {os.path.basename(output_path)}

🎯 NORMALIZATION ACHIEVED:
• Audio normalized to target loudness level
• True peaks controlled to prevent distortion
• Dynamic range preserved within safe limits

💡 INTERPRETING RESULTS:
• Integrated Loudness closer to target = successful normalization
• True Peak < 0 dBFS = no clipping/distortion
• Loudness Range indicates dynamic range preservation
"""

    summary_label = ttk.Label(results_dialog, text=summary_text, justify="left", font=("Arial", 9))
    summary_label.pack(pady=20, padx=20)

    # Close button
    ttk.Button(results_dialog, text="✅ Done", command=results_dialog.destroy).pack(pady=10)

    results_dialog.wait_window()


# -------------------------------------------------------
# Main program entry point
# -------------------------------------------------------
def main():
    """Main program execution flow."""
    # Set up basic logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

    logger.info("🎵 Audio Normalization Tool Started")

    try:
        # Check if test mode (no GUI)
        import sys
        if len(sys.argv) > 1 and sys.argv[1] == "--test":
            # Test mode - resolve file from argv[2] or env var
            test_file = (sys.argv[2] if len(sys.argv) > 2 else None) or os.environ.get("AUDIO_NORMALIZE_TEST_FILE")
            if not test_file:
                logger.error("❌ TEST MODE: No test file specified. Pass path as argv[2] or set AUDIO_NORMALIZE_TEST_FILE env var.")
                return
            input_path = test_file
            logger.info(f"🧪 TEST MODE: Using file: {os.path.basename(input_path)}")

            # Step 2: Analyze current levels
            analysis = analyze_audio_levels(input_path)
            if not analysis:
                logger.error("❌ Could not analyze test audio file")
                return

            logger.info(f"📊 Analysis Results: {analysis}")

            # Step 3: Use default target (-16 LUFS)
            target_lufs = -16.0
            logger.info(f"🎯 Using target loudness: {target_lufs} LUFS")

            # Step 4: Apply normalization
            output_path = normalize_audio(input_path, target_lufs)
            if not output_path:
                return

            # Step 5: Analyze result
            result_analysis = analyze_audio_levels(output_path)
            logger.info(f"✅ Normalization completed!")
            logger.info(f"📁 Output: {output_path}")
            if result_analysis:
                logger.info(f"📊 Final analysis: {result_analysis}")

            return

        # Normal GUI mode
        # Step 1: Choose input file
        input_path = choose_file()
        logger.info(f"📂 Selected file: {os.path.basename(input_path)}")

        # Step 2: Analyze current levels
        analysis = analyze_audio_levels(input_path)
        if not analysis:
            messagebox.showerror("Error", "Could not analyze audio file. Please check if it's a valid audio file.")
            return

        # Step 3: Show settings dialog
        target_lufs, proceed = create_normalization_dialog(analysis)
        if not proceed:
            logger.info("❌ Normalization canceled by user")
            return

        # Step 4: Apply normalization
        output_path = normalize_audio(input_path, target_lufs)
        if not output_path:
            return  # Error already shown

        # Step 5: Show results
        show_results(analysis, output_path)

        logger.info("🎉 Audio normalization process completed successfully")

    except Exception as e:
        logger.error(f"❌ Unexpected error: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")


if __name__ == "__main__":
    main()
