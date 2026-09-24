#!/usr/bin/env python3
"""Convert common document formats into Markdown files.

This script scans an input folder for .docx, .xlsx, .html, and .pdf files,
converts each file to Markdown, and writes the results to an output folder.
It includes a simple Tkinter GUI for selecting folders and a resilient
error-handling workflow to keep processing even when a file fails.

Requirements:
    pip install python-docx pandas openpyxl markdownify pypdf
"""

from __future__ import annotations

import logging
import os
import re
import threading
import traceback
from pathlib import Path
from typing import Iterable, List

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

TARGET_EXTENSIONS = {".docx", ".xlsx", ".html", ".pdf"}

try:
    from markdownify import markdownify
except ImportError:  # pragma: no cover
    markdownify = None

try:
    from docx import Document
except ImportError:  # pragma: no cover
    Document = None

try:
    import pandas as pd
except ImportError:  # pragma: no cover
    pd = None

try:
    from pypdf import PdfReader
except ImportError:  # pragma: no cover
    try:
        from PyPDF2 import PdfReader
    except ImportError:  # pragma: no cover
        PdfReader = None


def configure_logger(log_path: Path) -> logging.Logger:
    logger = logging.getLogger("md_converter")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger


def clean_text(value: str) -> str:
    """Normalize extracted text for Markdown output."""
    if value is None:
        return ""
    cleaned = value.replace("\r\n", "\n").replace("\r", "\n")
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    cleaned = re.sub(r"[ \t]+\n", "\n", cleaned)
    return cleaned.strip()


def escape_markdown_cell(value: object) -> str:
    text = "" if value is None else str(value)
    return text.replace("|", "\\|").replace("\n", "<br>").strip()


def table_to_markdown(rows: List[List[object]]) -> str:
    """Convert a 2D table into Markdown pipe syntax."""
    if not rows:
        return ""

    width = max(len(row) for row in rows)
    normalized = [list(row) + [""] * (width - len(row)) for row in rows]
    escaped = [[escape_markdown_cell(cell) for cell in row] for row in normalized]

    header = escaped[0]
    separator = ["---"] * width
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(separator) + " |",
    ]

    for row in escaped[1:]:
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines)


def docx_to_markdown(file_path: Path) -> str:
    """Convert a DOCX file to Markdown using python-docx."""
    if Document is None:
        raise RuntimeError("python-docx is not installed. Install with: pip install python-docx")

    document = Document(str(file_path))
    blocks: List[str] = []

    for paragraph in document.paragraphs:
        text = clean_text(paragraph.text)
        if text:
            blocks.append(text)

    for table in document.tables:
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        markdown_table = table_to_markdown(rows)
        if markdown_table:
            blocks.append(markdown_table)

    return "\n\n".join(blocks) if blocks else "# DOCX document\n\nNo readable content found."


def xlsx_to_markdown(file_path: Path) -> str:
    """Convert an XLSX workbook to Markdown, one sheet at a time."""
    if pd is None:
        raise RuntimeError("pandas/openpyxl are not installed. Install with: pip install pandas openpyxl")

    workbook_parts: List[str] = []
    xls_file = pd.ExcelFile(file_path, engine="openpyxl")

    for sheet_name in xls_file.sheet_names:
        dataframe = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
        if dataframe.empty:
            continue

        dataframe = dataframe.fillna("")
        column_names = ["" if col is None else str(col) for col in dataframe.columns]
        dataframe = dataframe.rename(columns=dict(zip(dataframe.columns, column_names)))

        rows: List[List[object]] = [column_names]
        for _, row in dataframe.iterrows():
            rows.append(["" if pd.isna(value) else str(value) for value in row.tolist()])

        sheet_markdown = table_to_markdown(rows)
        workbook_parts.append(f"## {sheet_name}\n\n{sheet_markdown}")

    if not workbook_parts:
        return "# Spreadsheet\n\nNo worksheet data found."

    return "\n\n".join(workbook_parts)


def html_to_markdown(file_path: Path) -> str:
    """Convert HTML content to Markdown using markdownify."""
    if markdownify is None:
        raise RuntimeError("markdownify is not installed. Install with: pip install markdownify")

    html_content = file_path.read_text(encoding="utf-8", errors="replace")
    markdown_text = markdownify(html_content, strip=["script", "style"], heading_style="ATX")
    cleaned = clean_text(markdown_text)
    return cleaned or "# HTML document\n\nNo readable content found."


def pdf_to_markdown(file_path: Path) -> str:
    """Convert a PDF file to Markdown by extracting text page by page."""
    if PdfReader is None:
        raise RuntimeError("pypdf is not installed. Install with: pip install pypdf")

    reader = PdfReader(str(file_path))
    pages: List[str] = []

    for page in reader.pages:
        text = page.extract_text() or ""
        cleaned = clean_text(text)
        if cleaned:
            pages.append(cleaned)

    if not pages:
        return "# PDF document\n\nNo extractable text found."

    return "\n\n---\n\n".join(pages)


def convert_file_to_markdown(file_path: Path) -> str:
    suffix = file_path.suffix.lower()

    if suffix == ".docx":
        return docx_to_markdown(file_path)
    if suffix == ".xlsx":
        return xlsx_to_markdown(file_path)
    if suffix == ".html":
        return html_to_markdown(file_path)
    if suffix == ".pdf":
        return pdf_to_markdown(file_path)

    raise ValueError(f"Unsupported file type: {suffix}")


def build_output_path(output_dir: Path, source_file: Path) -> Path:
    base_name = source_file.stem
    candidate = output_dir / f"{base_name}.md"
    counter = 1

    while candidate.exists():
        candidate = output_dir / f"{base_name}_{counter}.md"
        counter += 1

    return candidate


def find_target_files(input_dir: Path) -> List[Path]:
    matches: List[Path] = []
    for root, _, files in os.walk(input_dir):
        for file_name in files:
            path = Path(root) / file_name
            if path.suffix.lower() in TARGET_EXTENSIONS:
                matches.append(path)
    return sorted(matches)


def convert_folder(input_dir: Path, output_dir: Path, logger: logging.Logger) -> dict:
    input_dir = input_dir.expanduser().resolve()
    output_dir = output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    files = find_target_files(input_dir)
    if not files:
        raise FileNotFoundError(
            f"No supported files found in '{input_dir}'. Accepted extensions: {sorted(TARGET_EXTENSIONS)}"
        )

    stats = {"processed": 0, "errors": 0}

    for file_path in files:
        try:
            markdown_text = convert_file_to_markdown(file_path)
            output_path = build_output_path(output_dir, file_path)
            output_path.write_text(markdown_text, encoding="utf-8")
            stats["processed"] += 1
            logger.info("Converted %s -> %s", file_path.name, output_path.name)
            print(f"Converted: {file_path} -> {output_path}")
        except Exception as exc:  # noqa: BLE001
            stats["errors"] += 1
            logger.error("Failed to convert %s: %s", file_path, exc)
            logger.error(traceback.format_exc())
            print(f"ERROR: {file_path} -> {exc}")

    return stats


class FileToMarkdownApp(tk.Tk):
    """Tkinter GUI for selecting folders and starting conversion."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Document to Markdown Converter")
        self.geometry("760x420")
        self.minsize(650, 360)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=18)
        main.pack(fill="both", expand=True)

        title = ttk.Label(main, text="Convert files to Markdown", font=("Segoe UI", 14, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))

        ttk.Label(main, text="Input folder:", font=("Segoe UI", 10, "bold")).grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 8)
        )
        ttk.Entry(main, textvariable=self.input_var, width=62).grid(row=1, column=1, sticky="ew")
        ttk.Button(main, text="Browse", command=self.choose_input_folder, width=12).grid(
            row=1, column=2, padx=(10, 0), sticky="ew"
        )

        ttk.Label(main, text="Output folder:", font=("Segoe UI", 10, "bold")).grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=(10, 8)
        )
        ttk.Entry(main, textvariable=self.output_var, width=62).grid(row=2, column=1, sticky="ew")
        ttk.Button(main, text="Browse", command=self.choose_output_folder, width=12).grid(
            row=2, column=2, padx=(10, 0), sticky="ew"
        )

        self.run_button = ttk.Button(main, text="Convert files", command=self.start_conversion)
        self.run_button.grid(row=3, column=0, columnspan=3, pady=(18, 12), sticky="ew")

        ttk.Label(main, text="Status", font=("Segoe UI", 10, "bold")).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(0, 4)
        )

        self.status_label = ttk.Label(main, textvariable=self.status_var, foreground="#1f5f8b")
        self.status_label.grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 8))

        self.log_box = tk.Text(main, height=10, wrap="word", bg="#f8f8f8", relief="solid", borderwidth=1)
        self.log_box.grid(row=6, column=0, columnspan=3, sticky="nsew")
        self.log_box.insert("end", "Ready. Select an input folder and output folder to begin.\n")
        self.log_box.configure(state="disabled")

        main.columnconfigure(1, weight=1)
        main.rowconfigure(6, weight=1)

    def append_log(self, message: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def choose_input_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select input folder")
        if folder:
            self.input_var.set(folder)
            self.append_log(f"Input folder selected: {folder}")

    def choose_output_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_var.set(folder)
            self.append_log(f"Output folder selected: {folder}")

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def start_conversion(self) -> None:
        input_dir = self.input_var.get().strip()
        output_dir = self.output_var.get().strip()

        if not input_dir:
            messagebox.showerror("Input required", "Please select an input folder.")
            return

        if not output_dir:
            messagebox.showerror("Output required", "Please select an output folder.")
            return

        input_path = Path(input_dir).expanduser()
        output_path = Path(output_dir).expanduser()

        if not input_path.exists() or not input_path.is_dir():
            messagebox.showerror("Invalid input", f"The input folder does not exist: {input_path}")
            return

        output_path.mkdir(parents=True, exist_ok=True)
        self.set_status("Scanning files and converting...")
        self.append_log(f"Processing folder: {input_path}")
        self.run_button.config(state="disabled")

        thread = threading.Thread(
            target=self._conversion_worker,
            args=(input_path, output_path),
            daemon=True,
        )
        thread.start()

    def _conversion_worker(self, input_dir: Path, output_dir: Path) -> None:
        try:
            log_path = output_dir / "conversion_errors.log"
            logger = configure_logger(log_path)
            stats = convert_folder(input_dir, output_dir, logger)
            completed_summary = (
                f"Completed: {stats['processed']} files converted, "
                f"{stats['errors']} files failed. Errors are logged in {log_path}."
            )
            self.after(0, self.set_status, completed_summary)
            self.after(0, self.append_log, completed_summary)
            self.after(0, messagebox.showinfo, "Conversion complete", completed_summary)
        except Exception as exc:  # noqa: BLE001
            error_message = f"Conversion failed: {exc}"
            self.after(0, self.set_status, error_message)
            self.after(0, self.append_log, error_message)
            self.after(0, messagebox.showerror, "Conversion failed", error_message)
        finally:
            self.after(0, lambda: self.run_button.config(state="normal"))


if __name__ == "__main__":
    app = FileToMarkdownApp()
    app.mainloop()
