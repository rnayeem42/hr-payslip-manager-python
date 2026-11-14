"""Sanitized portfolio edition of an HR payslip workflow utility.

No employer, employee, payslip, or credential data is included.
"""

from __future__ import annotations

import csv
import re
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz
from openpyxl import load_workbook

try:
    import win32com.client as win32com
except ImportError:  # Allows non-Windows inspection and testing.
    win32com = None


APP_NAME = "HR Payslip Manager — Portfolio Edition"
EMPLOYEE_ID_RE = re.compile(r"\b(EMP\d{5,})\b", re.IGNORECASE)
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')
DRAFT_MARKER_NAME = "HRPortfolioDraft"
DRAFT_MARKER_VALUE = "HR_PAYSLIP_PORTFOLIO_EDITION"


def extract_employee_id(text: str) -> str | None:
    """Return a normalized employee ID found in extracted PDF text."""
    match = EMPLOYEE_ID_RE.search(text or "")
    return match.group(1).upper() if match else None


def unique_output_path(folder: Path, filename: str) -> Path:
    """Return a non-conflicting output path without overwriting a file."""
    target = folder / filename
    if not target.exists():
        return target
    for number in range(1, 10_000):
        candidate = folder / f"{target.stem}_{number}{target.suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError("Could not create a unique output filename.")


def split_payslips(pdf_path: Path, output_dir: Path) -> list[list[str | int]]:
    """Split a combined PDF and return auditable per-page results."""
    output_dir.mkdir(parents=True, exist_ok=True)
    results: list[list[str | int]] = []
    seen: set[str] = set()

    with fitz.open(pdf_path) as source:
        for index, page in enumerate(source):
            page_number = index + 1
            employee_id = extract_employee_id(page.get_text("text"))
            if not employee_id:
                results.append([page_number, "", "ERROR", "Employee ID not found"])
                continue
            if employee_id in seen:
                results.append([page_number, employee_id, "ERROR", "Duplicate employee ID"])
                continue

            safe_name = INVALID_FILENAME_CHARS.sub("_", employee_id) + ".pdf"
            target = unique_output_path(output_dir, safe_name)
            with fitz.open() as individual:
                individual.insert_pdf(source, from_page=index, to_page=index)
                individual.save(target)

            seen.add(employee_id)
            results.append([page_number, employee_id, "OK", target.name])

    return results


def load_employee_emails(excel_path: Path) -> dict[str, str]:
    """Load and validate Employee ID / Email mappings from an Excel sheet."""
    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        if "Employees" not in workbook.sheetnames:
            raise ValueError("Expected a sheet named 'Employees'.")
        sheet = workbook["Employees"]
        headers = {
            str(cell.value).strip().lower(): cell.column
            for cell in sheet[1]
            if cell.value is not None
        }
        id_col = headers.get("employee id")
        email_col = headers.get("email")
        if not id_col or not email_col:
            raise ValueError("Expected 'Employee ID' and 'Email' columns.")

        records: dict[str, str] = {}
        for row in range(2, sheet.max_row + 1):
            raw_id = sheet.cell(row=row, column=id_col).value
            raw_email = sheet.cell(row=row, column=email_col).value
            if raw_id is None:
                continue
            employee_id = str(raw_id).strip().upper()
            email = str(raw_email or "").strip()
            if employee_id in records:
                raise ValueError(f"Duplicate employee ID: {employee_id}")
            if not EMPLOYEE_ID_RE.fullmatch(employee_id):
                raise ValueError(f"Invalid employee ID format: {employee_id}")
            records[employee_id] = email
        return records
    finally:
        workbook.close()


def create_outlook_drafts(output_dir: Path, excel_path: Path) -> list[list[str]]:
    """Create tagged Outlook drafts. This function never sends email."""
    if win32com is None:
        raise RuntimeError("Outlook automation requires Windows and pywin32.")
    recipients = load_employee_emails(excel_path)
    outlook = win32com.Dispatch("Outlook.Application")
    results: list[list[str]] = []

    for pdf_path in sorted(output_dir.glob("EMP*.pdf")):
        employee_id = pdf_path.stem.split("_")[0].upper()
        email = recipients.get(employee_id, "")
        if not email:
            results.append([employee_id, "ERROR", "No email mapping found"])
            continue

        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = f"Payslip - {employee_id}"
        mail.HTMLBody = (
            "<p>Dear Colleague,</p>"
            "<p>Please find your current payslip attached.</p>"
            f"<p>Employee ID: <b>{employee_id}</b></p>"
            "<p>Kind regards,<br>HR Team</p>"
        )
        mail.Attachments.Add(str(pdf_path.resolve()))
        marker = mail.UserProperties.Add(DRAFT_MARKER_NAME, 1, False)
        marker.Value = DRAFT_MARKER_VALUE
        mail.Save()
        results.append([employee_id, "DRAFT CREATED", pdf_path.name])

    return results


def write_log(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        writer.writerows(rows)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("850x620")
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.status = tk.StringVar(value="Ready")
        self._build()

    def _build(self) -> None:
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=APP_NAME, font=("Segoe UI", 20, "bold")).pack(anchor="w")
        ttk.Label(
            frame,
            text="Split local payslip PDFs, validate employee IDs, and prepare Outlook drafts.",
        ).pack(anchor="w", pady=(2, 18))

        self._path_row(frame, "Combined PDF", self.pdf_path, self._pick_pdf)
        self._path_row(frame, "Output folder", self.output_dir, self._pick_output)
        self._path_row(frame, "Employee workbook", self.excel_path, self._pick_excel)

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=12)
        ttk.Button(actions, text="Split payslips", command=self._start_split).pack(side="left")
        ttk.Button(actions, text="Create Outlook drafts", command=self._start_drafts).pack(side="left", padx=8)

        self.table = ttk.Treeview(frame, columns=("id", "status", "details"), show="headings")
        for column, title, width in (
            ("id", "Employee ID", 150),
            ("status", "Status", 150),
            ("details", "Details", 450),
        ):
            self.table.heading(column, text=title)
            self.table.column(column, width=width, anchor="w")
        self.table.pack(fill="both", expand=True, pady=8)
        ttk.Label(frame, textvariable=self.status).pack(anchor="w")

    def _path_row(self, parent, label, variable, command) -> None:
        ttk.Label(parent, text=label).pack(anchor="w")
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(2, 10))
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Browse", command=command).pack(side="left", padx=(8, 0))

    def _pick_pdf(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf_path.set(path)

    def _pick_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def _pick_excel(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self.excel_path.set(path)

    def _clear_results(self) -> None:
        for item in self.table.get_children():
            self.table.delete(item)

    def _start_split(self) -> None:
        threading.Thread(target=self._run_split, daemon=True).start()

    def _run_split(self) -> None:
        try:
            pdf = Path(self.pdf_path.get())
            output = Path(self.output_dir.get())
            if not pdf.is_file() or not self.output_dir.get():
                raise ValueError("Select a valid PDF and output folder.")
            rows = split_payslips(pdf, output)
            write_log(output / "split_log.csv", ["Page", "Employee ID", "Status", "Details"], rows)
            self.after(0, lambda: self._show([[str(r[1]), str(r[2]), str(r[3])] for r in rows]))
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror(APP_NAME, str(exc)))

    def _start_drafts(self) -> None:
        if not messagebox.askyesno(APP_NAME, "Create Outlook drafts? No messages will be sent."):
            return
        threading.Thread(target=self._run_drafts, daemon=True).start()

    def _run_drafts(self) -> None:
        try:
            output = Path(self.output_dir.get())
            excel = Path(self.excel_path.get())
            if not output.is_dir() or not excel.is_file():
                raise ValueError("Select a valid output folder and employee workbook.")
            rows = create_outlook_drafts(output, excel)
            write_log(output / "draft_log.csv", ["Employee ID", "Status", "Details"], rows)
            self.after(0, lambda: self._show(rows))
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror(APP_NAME, str(exc)))

    def _show(self, rows: list[list[str]]) -> None:
        self._clear_results()
        for row in rows:
            self.table.insert("", "end", values=row)
        self.status.set(f"Completed: {len(rows)} record(s)")


if __name__ == "__main__":
    App().mainloop()


