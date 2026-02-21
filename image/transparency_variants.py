#!/usr/bin/env python3
"""
transparency_variants.py - Create image variants with a given transparency %.

Launch the tool, select an image file, enter transparency % (0=opaque, 100=fully transparent),
then create a copy with "_transp_xx" suffix (saved as PNG for alpha support).
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    Image = None

# Supported input formats
SUPPORTED_EXT = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'}


def make_transparent_variant(source_path: Path, transparency_pct: int, out_path: Path) -> bool:
    """
    Create a copy of the image with the given transparency (0=opaque, 100=fully transparent).
    Saves as PNG to preserve alpha. Returns True on success.
    """
    if Image is None:
        raise RuntimeError("Pillow (PIL) is required. Install with: pip install Pillow")

    if not 0 <= transparency_pct <= 100:
        raise ValueError("Transparency must be between 0 and 100")

    img = Image.open(source_path).convert("RGBA")
    alpha = img.split()[3]

    # opacity = 1 - (transparency/100)  =>  alpha factor
    opacity = 1.0 - (transparency_pct / 100.0)
    new_alpha = alpha.point(lambda x: int(x * opacity))
    img.putalpha(new_alpha)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path, "PNG")
    return True


class TransparencyVariantsApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Transparency variants")
        self.root.geometry("420x200")
        self.root.resizable(True, False)

        self.source_path: Path | None = None
        self.transparency_var = tk.IntVar(value=50)

        self._build_ui()

    def _build_ui(self):
        pad = (8, 12)  # (vertical, horizontal) padding for ttk.Frame

        # File selection
        f_file = ttk.Frame(self.root, padding=pad)
        f_file.pack(fill=tk.X)
        ttk.Button(f_file, text="Select image…", command=self._choose_file).pack(side=tk.LEFT)
        self.lbl_file = ttk.Label(f_file, text="No file selected", foreground="gray")
        self.lbl_file.pack(side=tk.LEFT, padx=(8, 0))

        # Transparency
        f_transp = ttk.Frame(self.root, padding=pad)
        f_transp.pack(fill=tk.X)
        ttk.Label(f_transp, text="Transparency % (0=opaque, 100=invisible):").pack(anchor=tk.W)
        sub = ttk.Frame(f_transp)
        sub.pack(fill=tk.X)
        scale = ttk.Scale(
            sub,
            from_=0,
            to=100,
            variable=self.transparency_var,
            orient=tk.HORIZONTAL,
            command=lambda v: self.transparency_var.set(round(float(v))),
        )
        scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.lbl_pct = ttk.Label(sub, text="50%", width=4)
        self.lbl_pct.pack(side=tk.LEFT)
        self.transparency_var.trace_add("write", self._update_pct_label)

        # Create button
        f_btn = ttk.Frame(self.root, padding=pad)
        f_btn.pack(fill=tk.X)
        self.btn_create = ttk.Button(
            f_btn,
            text="Create variant",
            command=self._create_variant,
            state=tk.DISABLED,
        )
        self.btn_create.pack()

        # Status
        self.lbl_status = ttk.Label(self.root, text="", foreground="gray")
        self.lbl_status.pack(pady=(0, 8))
        if Image is None:
            self.lbl_status.config(text="Pillow required: pip install Pillow", foreground="red")

    def _update_pct_label(self, *_):
        try:
            v = self.transparency_var.get()
            self.lbl_pct.config(text=f"{v}%")
        except tk.TclError:
            pass

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="Select image",
            filetypes=[
                ("Images", " ".join("*" + ext for ext in SUPPORTED_EXT)),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.source_path = Path(path)
            name = self.source_path.name
            self.lbl_file.config(text=name if len(name) <= 50 else name[:47] + "...", foreground="")
            self.btn_create.config(state=tk.NORMAL)
            self.lbl_status.config(text="")

    def _create_variant(self):
        if self.source_path is None or not self.source_path.is_file():
            messagebox.showerror("Error", "Please select a valid image file.")
            return

        try:
            pct = int(self.transparency_var.get())
        except (tk.TclError, ValueError):
            messagebox.showerror("Error", "Transparency must be a number between 0 and 100.")
            return

        if not 0 <= pct <= 100:
            messagebox.showerror("Error", "Transparency must be between 0 and 100.")
            return

        if Image is None:
            messagebox.showerror("Error", "Pillow is required. Install with: pip install Pillow")
            return

        # Output: same dir, stem_transp_XX.png
        out_name = f"{self.source_path.stem}_transp_{pct:02d}.png"
        out_path = self.source_path.parent / out_name

        try:
            make_transparent_variant(self.source_path, pct, out_path)
            self.lbl_status.config(text=f"Saved: {out_path.name}", foreground="green")
            messagebox.showinfo("Done", f"Created:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.lbl_status.config(text="", foreground="red")


def main():
    root = tk.Tk()
    TransparencyVariantsApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
