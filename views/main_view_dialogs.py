from tkinter import filedialog, messagebox
from typing import List, Optional

# AI WARNING: This module contains AI-assisted code.
# Review and test changes carefully before release.


def show_warning(title: str, message: str) -> None:
    messagebox.showwarning(title, message)


def show_error(title: str, message: str) -> None:
    messagebox.showerror(title, message)


def show_info(title: str, message: str) -> None:
    messagebox.showinfo(title, message)


def ask_confirm(title: str, message: str) -> bool:
    return messagebox.askyesno(title, message)


def ask_save_file(initialfile: str = "merged_presentation.pdf") -> Optional[str]:
    return filedialog.asksaveasfilename(
        title="Save Merged PDF As",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        initialfile=initialfile,
    )


def ask_open_files() -> List[str]:
    filetypes = [
        ("PowerPoint files", "*.pptx *.ppt"),
        ("PPTX files", "*.pptx"),
        ("PPT files", "*.ppt"),
        ("All files", "*.*"),
    ]
    return list(
        filedialog.askopenfilenames(
            title="Select PowerPoint Files",
            filetypes=filetypes,
        )
    )


def ask_directory() -> Optional[str]:
    return filedialog.askdirectory(title="Select Output Folder for PDFs")
