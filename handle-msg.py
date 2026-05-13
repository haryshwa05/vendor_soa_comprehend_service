from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

from msg_to_excel import MsgConversionError, convert_msg_to_excel


class MsgConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("MSG to Excel")
        self.root.geometry("620x250")
        self.root.minsize(620, 250)

        self.selected_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Select a .msg file to begin.")

        self.build_ui()

    def build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)

        title = ttk.Label(
            container,
            text="Convert Outlook .msg email body table to Excel",
            font=("Segoe UI", 13, "bold"),
        )
        title.pack(anchor="w", pady=(0, 8))

        description = ttk.Label(
            container,
            text=(
                "Pick one .msg file. The script tries the email body first, then Excel and PDF "
                "attachments, extracts invoice number, invoice date, and outstanding amount, "
                "and writes an .xlsx file in this script folder."
            ),
            wraplength=580,
            justify="left",
        )
        description.pack(anchor="w", pady=(0, 14))

        path_frame = ttk.Frame(container)
        path_frame.pack(fill="x", pady=(0, 12))

        entry = ttk.Entry(path_frame, textvariable=self.selected_path, state="readonly")
        entry.pack(side="left", fill="x", expand=True)

        browse_btn = ttk.Button(path_frame, text="Browse", command=self.browse_file)
        browse_btn.pack(side="left", padx=(8, 0))

        convert_btn = ttk.Button(container, text="Convert", command=self.convert_file)
        convert_btn.pack(anchor="w", pady=(0, 12))

        status_label = ttk.Label(
            container,
            textvariable=self.status_text,
            wraplength=580,
            justify="left",
            foreground="#1f2937",
        )
        status_label.pack(anchor="w")

    def browse_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Outlook MSG file",
            filetypes=[("Outlook MSG files", "*.msg"), ("All files", "*.*")],
        )
        if file_path:
            self.selected_path.set(file_path)
            self.status_text.set("File selected. Click Convert to create the Excel file.")

    def convert_file(self) -> None:
        raw_path = self.selected_path.get().strip()
        if not raw_path:
            self.status_text.set("Choose a .msg file first.")
            return

        msg_path = Path(raw_path)
        if msg_path.suffix.lower() != ".msg":
            self.status_text.set("Selected file must have a .msg extension.")
            return

        try:
            output_path = convert_msg_to_excel(msg_path)
        except MsgConversionError as exc:
            self.status_text.set(f"Conversion failed: {exc}")
            return
        except Exception as exc:
            self.status_text.set(f"Unexpected error: {exc}")
            return

        self.status_text.set(f"Excel created: {output_path}")


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = MsgConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
