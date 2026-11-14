# HR Payslip Manager — Portfolio Edition

A privacy-conscious Windows desktop utility that demonstrates an end-to-end HR document workflow:

- split a combined payslip PDF into one file per employee;
- extract and validate employee IDs from PDF text;
- match employee IDs to email addresses in an Excel workbook;
- create individually addressed Outlook drafts with the correct attachment;
- record processing outcomes in CSV audit logs;
- prevent duplicate employee IDs and avoid overwriting existing files.

This repository is a **sanitized portfolio edition** of a tool built from a real HR workflow. It contains no employer data, employee data, payslips, email addresses, branding, credentials, or internal documents. Independent software project; not an official product of any employer.

## Stack

- Python 3.11+
- Tkinter
- PyMuPDF
- openpyxl
- Microsoft Outlook COM automation (`pywin32`)

## Excel format

Create an `.xlsx` workbook with a sheet named `Employees` and these headers:

| Employee ID | Email |
|---|---|
| EMP10001 | employee@example.com |

## Run locally

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Outlook integration requires Windows and the classic Microsoft Outlook desktop client.

## Privacy and safety

- Input files remain on the local computer.
- The app creates drafts; it does not automatically send emails.
- Drafts are tagged so they can be distinguished from unrelated Outlook items.
- Do not publish real payslips or employee lists in this repository.

## Author

Nayeem Rahman — HR & Administration professional and independent software developer.


