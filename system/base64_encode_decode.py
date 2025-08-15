import base64
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def select_file(title="Select file"):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title)
    return file_path

def encode_file():
    file_path = select_file("Select the file to encode")
    if not file_path:
        messagebox.showwarning("Warning", "No file selected!")
        return

    output_path = filedialog.asksaveasfilename(
        title="Save encoded file as",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")]
    )
    if not output_path:
        messagebox.showwarning("Warning", "No output file selected!")
        return

    with open(file_path, 'rb') as binary_file:
        encoded_data = base64.b64encode(binary_file.read()).decode('utf-8')

    with open(output_path, 'w') as text_file:
        text_file.write(encoded_data)

    messagebox.showinfo("Success", f"File encoded successfully!\nSaved at:\n{output_path}")

def decode_file():
    file_path = select_file("Select the Base64 text file to decode")
    if not file_path:
        messagebox.showwarning("Warning", "No file selected!")
        return

    output_path = filedialog.asksaveasfilename(
        title="Save decoded file as",
        defaultextension="",
        filetypes=[("Executable", "*.exe"), ("All Files", "*.*")]
    )
    if not output_path:
        messagebox.showwarning("Warning", "No output file selected!")
        return

    with open(file_path, 'r') as text_file:
        encoded_data = text_file.read()

    with open(output_path, 'wb') as binary_file:
        binary_file.write(base64.b64decode(encoded_data))

    messagebox.showinfo("Success", f"File decoded successfully!\nSaved at:\n{output_path}")

def main():
    root = tk.Tk()
    root.title("Base64 Encoder/Decoder")
    root.geometry("350x150")

    encode_button = tk.Button(root, text="Encode File to Base64", command=encode_file, width=30)
    encode_button.pack(pady=10)

    decode_button = tk.Button(root, text="Decode Base64 to File", command=decode_file, width=30)
    decode_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
