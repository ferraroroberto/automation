"""
Main launcher script for the voice transcription application.
This script provides a unified entry point for both console and GUI modes.
"""

import sys
import argparse


def main():
    """Main entry point for the transcription application."""
    parser = argparse.ArgumentParser(description="Voice recording and transcription tool")
    parser.add_argument("--gui", action="store_true", 
                       help="Launch with graphical interface")
    parser.add_argument("--launch_from_elgato", action="store_true",
                       help="Launch with GUI interface for Elgato Stream Deck")
    parser.add_argument("--file", type=str, 
                       help="Transcribe a specific audio file")
    parser.add_argument("--language", type=str, choices=["Spanish", "English"], 
                       default="Spanish", help="Language for transcription")
    parser.add_argument("--model", type=str, default="base",
                       choices=["tiny", "base", "small", "medium", "large"],
                       help="Whisper model size")
    args = parser.parse_args()
    
    # Determine which mode to use
    if args.gui or args.launch_from_elgato:
        # GUI mode
        from transcribe_voice_gui import main as gui_main
        gui_main()
    else:
        # Console mode
        from transcribe_voice_core import TranscriptionConfig, transcribe_file, run_console_transcription
        
        if args.file:
            # File mode
            config = TranscriptionConfig(
                language=args.language,
                translate=(args.language == "Spanish"),
                model_size=args.model
            )
            try:
                import pyperclip
                text = transcribe_file(args.file, config)
                print("\n=== TRANSCRIPTION ===\n")
                print(text)
                pyperclip.copy(text)
                print("\n✅ Text copied to clipboard.")
            except Exception as e:
                print(f"❌ ERROR: {e}")
                sys.exit(1)
        else:
            # Interactive recording mode
            try:
                config = TranscriptionConfig(model_size=args.model)
                exit_code = run_console_transcription(config)
                sys.exit(exit_code)
            except KeyboardInterrupt:
                print("\n⚠️  Interrupted by user")
                sys.exit(1)


if __name__ == "__main__":
    main()