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


def main() -> None:
    root = tk.Tk()
    root.title("Base64 Encoder/Decoder")
    root.geometry("360x260")

    tk.Button(root, text="Encode File", command=encode_file_handler, width=30).pack(pady=8)
    tk.Button(root, text="Decode File", command=decode_file_handler, width=30).pack(pady=8)
    tk.Button(root, text="Encode Folder", command=encode_folder_handler, width=30).pack(pady=8)
    tk.Button(root, text="Decode Folder", command=decode_folder_handler, width=30).pack(pady=8)

    root.mainloop()


if __name__ == "__main__":
    main()
