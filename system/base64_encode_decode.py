import base64
import logging
from pathlib import Path
from typing import Iterable

import tkinter as tk
from tkinter import filedialog, messagebox

LOG_FORMAT = "%(levelname)s | %(message)s"
logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
logger = logging.getLogger(__name__)


def encode_path(source: Path, target: Path) -> None:
    with source.open("rb") as binary_file:
        encoded = base64.b64encode(binary_file.read()).decode("utf-8")
    with target.open("w", encoding="utf-8") as text_file:
        text_file.write(encoded)


def decode_path(source: Path, target: Path) -> None:
    with source.open("r", encoding="utf-8") as text_file:
        data = text_file.read()
    with target.open("wb") as binary_file:
        binary_file.write(base64.b64decode(data))


def encode_and_split_path(source: Path, target_base: Path, chunk_size_mb: int) -> int:
    """Encode file to Base64 and split into chunks of specified size in MB.

    Args:
        source: Source file path
        target_base: Base path for output files (without extension)
        chunk_size_mb: Size of each chunk in MB

    Returns:
        Number of chunks created
    """
    with source.open("rb") as binary_file:
        encoded = base64.b64encode(binary_file.read()).decode("utf-8")

    chunk_size_bytes = chunk_size_mb * 1024 * 1024  # Convert MB to bytes
    chunks = []

    for i in range(0, len(encoded), chunk_size_bytes):
        chunks.append(encoded[i:i + chunk_size_bytes])

    # Write chunks with sequential numbering
    for idx, chunk in enumerate(chunks, 1):
        chunk_filename = f"{target_base.stem}_{idx:03d}.txt"
        chunk_path = target_base.with_name(chunk_filename)
        with chunk_path.open("w", encoding="utf-8") as text_file:
            text_file.write(chunk)
        logger.info("✅ Created chunk %d: %s", idx, chunk_path.name)

    return len(chunks)


def join_and_decode_path(selected_file: Path, target: Path) -> None:
    """Find related split files, join them, and decode to target file.

    Args:
        selected_file: One of the split files (e.g., file_001.txt)
        target: Target file path for decoded output
    """
    import re

    # Extract base name pattern (everything before the _XXX.txt part)
    match = re.match(r"^(.+?)_\d{3}\.txt$", selected_file.name)
    if not match:
        raise ValueError(f"Selected file {selected_file.name} doesn't match split file pattern")

    base_name = match.group(1)
    parent_dir = selected_file.parent

    # Find all related files
    pattern = re.compile(rf"^{re.escape(base_name)}_\d{{3}}\.txt$")
    related_files = []

    for file_path in parent_dir.iterdir():
        if file_path.is_file() and pattern.match(file_path.name):
            # Extract the number from filename
            num_match = re.search(r"_(\d{3})\.txt$", file_path.name)
            if num_match:
                num = int(num_match.group(1))
                related_files.append((num, file_path))

    if not related_files:
        raise ValueError(f"No related split files found for {selected_file.name}")

    # Sort by number and read in order
    related_files.sort(key=lambda x: x[0])
    logger.info("Found %d related files: %s", len(related_files),
                [f.name for _, f in related_files])

    # Read and join all chunks
    encoded_data = ""
    for num, file_path in related_files:
        with file_path.open("r", encoding="utf-8") as f:
            encoded_data += f.read()
        logger.info("✅ Read chunk %d: %s", num, file_path.name)

    # Decode and write to target
    with target.open("wb") as binary_file:
        binary_file.write(base64.b64decode(encoded_data))


def iter_folder_files(folder: Path) -> Iterable[Path]:
    for path in folder.iterdir():
        if path.is_file():
            yield path


def encode_file_handler() -> None:
    file_path = filedialog.askopenfilename(title="Select the file to encode")
    if not file_path:
        messagebox.showwarning("Warning", "No file selected!")
        return
    output_path = filedialog.asksaveasfilename(
        title="Save encoded file as",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")],
    )
    if not output_path:
        messagebox.showwarning("Warning", "No output file selected!")
        return

    try:
        encode_path(Path(file_path), Path(output_path))
    except Exception as error:  # noqa: BLE001
        logger.error("❌ Failed to encode %s: %s", file_path, error)
        messagebox.showerror("Error", f"Encoding failed:\n{error}")
        return

    logger.info("✅ Encoded file saved to %s", output_path)
    messagebox.showinfo("Success", f"File encoded successfully!\nSaved at:\n{output_path}")


def decode_file_handler() -> None:
    file_path = filedialog.askopenfilename(
        title="Select the Base64 text file to decode", filetypes=[("Text files", "*.txt")]
    )
    if not file_path:
        messagebox.showwarning("Warning", "No file selected!")
        return
    output_path = filedialog.asksaveasfilename(
        title="Save decoded file as",
        defaultextension="",
        filetypes=[("All Files", "*.*")],
    )
    if not output_path:
        messagebox.showwarning("Warning", "No output file selected!")
        return

    try:
        decode_path(Path(file_path), Path(output_path))
    except Exception as error:  # noqa: BLE001
        logger.error("❌ Failed to decode %s: %s", file_path, error)
        messagebox.showerror("Error", f"Decoding failed:\n{error}")
        return

    logger.info("✅ Decoded file saved to %s", output_path)
    messagebox.showinfo("Success", f"File decoded successfully!\nSaved at:\n{output_path}")


def encode_folder_handler() -> None:
    folder_path = filedialog.askdirectory(title="Select folder to encode")
    if not folder_path:
        messagebox.showwarning("Warning", "No folder selected!")
        return

    folder = Path(folder_path)
    processed = skipped = errors = 0
    for file_path in iter_folder_files(folder):
        if file_path.suffix.lower() == ".txt":
            skipped += 1
            logger.info("ℹ️  Skipping text file %s", file_path.name)
            continue
        target_path = file_path.with_name(f"{file_path.name}.txt")
        if target_path.exists():
            skipped += 1
            logger.warning("⚠️  Target %s already exists. Skipping.", target_path.name)
            continue
        try:
            encode_path(file_path, target_path)
            processed += 1
            logger.info("✅ Encoded %s -> %s", file_path.name, target_path.name)
        except Exception as error:  # noqa: BLE001
            errors += 1
            logger.error("❌ Failed to encode %s: %s", file_path.name, error)

    messagebox.showinfo(
        "Folder Encode Complete",
        f"Processed: {processed}\nSkipped: {skipped}\nErrors: {errors}",
    )


def decode_folder_handler() -> None:
    folder_path = filedialog.askdirectory(title="Select folder to decode")
    if not folder_path:
        messagebox.showwarning("Warning", "No folder selected!")
        return

    folder = Path(folder_path)
    processed = skipped = errors = 0
    for file_path in iter_folder_files(folder):
        if file_path.suffix.lower() != ".txt":
            skipped += 1
            logger.info("ℹ️  Skipping non-text file %s", file_path.name)
            continue
        target_name = file_path.name[:-4]
        if not target_name:
            skipped += 1
            logger.warning("⚠️  Cannot derive output name from %s", file_path.name)
            continue
        target_path = file_path.with_name(target_name)
        if target_path.exists():
            skipped += 1
            logger.warning("⚠️  Target %s already exists. Skipping.", target_path.name)
            continue
        try:
            decode_path(file_path, target_path)
            processed += 1
            logger.info("✅ Decoded %s -> %s", file_path.name, target_path.name)
        except Exception as error:  # noqa: BLE001
            errors += 1
            logger.error("❌ Failed to decode %s: %s", file_path.name, error)

    messagebox.showinfo(
        "Folder Decode Complete",
        f"Processed: {processed}\nSkipped: {skipped}\nErrors: {errors}",
    )


def get_chunk_size(root: tk.Tk) -> int | None:
    """Get chunk size in MB from user input dialog."""
    dialog = tk.Toplevel(root)
    dialog.title("Chunk Size")
    dialog.geometry("300x120")
    dialog.resizable(False, False)

    tk.Label(dialog, text="Enter chunk size in MB:").pack(pady=10)

    size_var = tk.StringVar()
    entry = tk.Entry(dialog, textvariable=size_var, width=20)
    entry.pack(pady=5)
    entry.focus()

    result = None

    def on_ok():
        nonlocal result
        try:
            size = int(size_var.get())
            if size <= 0:
                raise ValueError("Size must be positive")
            result = size
            dialog.destroy()
        except ValueError as e:
            messagebox.showerror("Invalid Input", f"Please enter a valid positive integer.\n{str(e)}")

    def on_cancel():
        dialog.destroy()

    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Cancel", command=on_cancel, width=10).pack(side=tk.RIGHT, padx=5)

    dialog.transient(root)
    dialog.grab_set()
    dialog.wait_window()

    return result


def encode_and_split_handler(root: tk.Tk) -> None:
    file_path = filedialog.askopenfilename(title="Select the file to encode and split")
    if not file_path:
        messagebox.showwarning("Warning", "No file selected!")
        return

    chunk_size = get_chunk_size(root)
    if chunk_size is None:
        return

    # Ask for output directory and base filename
    output_path = filedialog.asksaveasfilename(
        title="Save split files as (base name)",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")],
    )
    if not output_path:
        messagebox.showwarning("Warning", "No output location selected!")
        return

    output_path_obj = Path(output_path)
    source_path_obj = Path(file_path)

    try:
        num_chunks = encode_and_split_path(source_path_obj, output_path_obj, chunk_size)
    except Exception as error:  # noqa: BLE001
        logger.error("❌ Failed to encode and split %s: %s", file_path, error)
        messagebox.showerror("Error", f"Encoding and splitting failed:\n{error}")
        return

    logger.info("✅ File split into %d chunks", num_chunks)
    messagebox.showinfo(
        "Success",
        f"File successfully encoded and split into {num_chunks} chunks!\n"
        f"Chunk size: {chunk_size} MB\n"
        f"Files saved in: {output_path_obj.parent}"
    )


def join_and_decode_handler() -> None:
    file_path = filedialog.askopenfilename(
        title="Select one of the split files to join and decode",
        filetypes=[("Text files", "*.txt")]
    )
    if not file_path:
        messagebox.showwarning("Warning", "No file selected!")
        return

    # Ask for output file
    output_path = filedialog.asksaveasfilename(
        title="Save decoded file as",
        defaultextension="",
        filetypes=[("All Files", "*.*")],
    )
    if not output_path:
        messagebox.showwarning("Warning", "No output file selected!")
        return

    try:
        join_and_decode_path(Path(file_path), Path(output_path))
    except Exception as error:  # noqa: BLE001
        logger.error("❌ Failed to join and decode: %s", error)
        messagebox.showerror("Error", f"Joining and decoding failed:\n{error}")
        return

    logger.info("✅ Joined and decoded file saved to %s", output_path)
    messagebox.showinfo("Success", f"Files joined and decoded successfully!\nSaved at:\n{output_path}")


def main() -> None:
    root = tk.Tk()
    root.title("Base64 Encoder/Decoder")
    root.geometry("360x320")

    tk.Button(root, text="Encode File", command=encode_file_handler, width=30).pack(pady=8)
    tk.Button(root, text="Decode File", command=decode_file_handler, width=30).pack(pady=8)
    tk.Button(root, text="Encode Folder", command=encode_folder_handler, width=30).pack(pady=8)
    tk.Button(root, text="Decode Folder", command=decode_folder_handler, width=30).pack(pady=8)
    tk.Button(root, text="Encode & Split File", command=lambda: encode_and_split_handler(root), width=30).pack(pady=8)
    tk.Button(root, text="Join & Decode File", command=join_and_decode_handler, width=30).pack(pady=8)

    root.mainloop()


if __name__ == "__main__":
    main()
