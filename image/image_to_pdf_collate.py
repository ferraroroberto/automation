import logging
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional

import pillow_heif
from _image_to_jpg_converter import flatten_to_rgb
from PIL import Image

pillow_heif.register_heif_opener()

logger = logging.getLogger(__name__)

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp", ".heic", ".heif"}


def collect_images(folder: Path) -> list[Path]:
    files = sorted(
        (f for f in folder.iterdir() if f.is_file() and f.suffix.lower() in IMAGE_EXTENSIONS),
        key=lambda p: p.name.lower(),
    )
    return files


def images_to_pdf(image_paths: list[Path], output_path: Path) -> None:
    pages: list[Image.Image] = []
    for path in image_paths:
        img = Image.open(path)
        pages.append(flatten_to_rgb(img))

    if not pages:
        raise ValueError("No valid images to save.")

    first, rest = pages[0], pages[1:]
    first.save(output_path, "PDF", save_all=True, append_images=rest)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Image → PDF Collate")
        self.resizable(False, False)
        self._build_ui()

    def _build_ui(self) -> None:
        frame = tk.Frame(self, padx=20, pady=20)
        frame.pack()

        tk.Label(frame, text="Select a folder to collate its images into a PDF.", wraplength=320).pack(pady=(0, 12))

        tk.Button(frame, text="Choose Folder & Create PDF", width=30, command=self._run).pack()

        self._status = tk.StringVar(value="")
        tk.Label(frame, textvariable=self._status, wraplength=320, fg="gray").pack(pady=(12, 0))

    def _run(self) -> None:
        folder_str: Optional[str] = filedialog.askdirectory(title="Select folder with images")
        if not folder_str:
            return

        folder = Path(folder_str)
        self._status.set("⏳ Scanning…")
        self.update_idletasks()

        images = collect_images(folder)
        if not images:
            self._status.set("No image files found in the selected folder.")
            messagebox.showwarning("No images", "No image files found in the selected folder.")
            return

        output_path = folder / f"{folder.name}.pdf"
        self._status.set(f"⏳ Building PDF from {len(images)} image(s)…")
        self.update_idletasks()

        try:
            images_to_pdf(images, output_path)
        except Exception as exc:
            logger.error("PDF creation failed: %s", exc)
            self._status.set(f"❌ Error: {exc}")
            messagebox.showerror("Error", str(exc))
            return

        logger.info("✅ PDF saved: %s", output_path)
        self._status.set(f"✅ Saved: {output_path.name}")
        messagebox.showinfo("Done", f"PDF created:\n{output_path}")


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
