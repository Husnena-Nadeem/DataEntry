import re
from io import BytesIO

from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from pydantic import BaseModel

app = FastAPI(title="Email to Excel API")

ALLOWED_ALIGNMENTS = ("left", "center", "right")


class Payload(BaseModel):
    records: list[dict]
    options: dict = {}


def clean(value) -> str:
    """Extra spaces hatao aur text ko saaf karo."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def validate(record: dict, row_no: int) -> list[str]:
    """Basic validation. Masail ki list return karta hai."""
    issues = []
    for key, value in record.items():
        if value == "":
            issues.append(f"Row {row_no}: '{key}' khali hai")

    phone = record.get("Phone")
    if phone:
        digits = re.sub(r"\D", "", phone)
        if len(digits) != 11:
            issues.append(f"Row {row_no}: Phone 11 digits ka nahi hai ({phone})")

    age = record.get("Age")
    if age:
        if not age.isdigit() or not (0 < int(age) < 120):
            issues.append(f"Row {row_no}: Age sahi nahi hai ({age})")

    return issues


def get_alignment(options: dict) -> str:
    """Sirf left/center/right allow karo, warna default left."""
    value = clean(options.get("alignment", "left")).lower()
    return value if value in ALLOWED_ALIGNMENTS else "left"


@app.get("/")
def health():
    return {"status": "ok"}


@app.post("/generate-excel")
def generate_excel(p: Payload):
    records = [{clean(k): clean(v) for k, v in r.items()} for r in p.records]

    headers = []
    for r in records:
        for k in r:
            if k not in headers:
                headers.append(k)

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(headers)

    all_issues = []
    for i, r in enumerate(records, start=2):
        row = {h: r.get(h, "") for h in headers}
        ws.append([row[h] for h in headers])
        all_issues.extend(validate(row, i))

    # Formatting: header bold + customer ki alignment
    align = get_alignment(p.options)
    for row_cells in ws.iter_rows():
        for cell in row_cells:
            cell.alignment = Alignment(horizontal=align)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Column width
    for col in ws.columns:
        width = max(len(str(c.value or "")) for c in col) + 3
        ws.column_dimensions[col[0].column_letter].width = width

    if all_issues:
        issues_ws = wb.create_sheet("Issues")
        issues_ws.append(["Issue"])
        issues_ws["A1"].font = Font(bold=True)
        for msg in all_issues:
            issues_ws.append([msg])
        issues_ws.column_dimensions["A"].width = 60

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=output.xlsx",
            "X-Issue-Count": str(len(all_issues)),
        },
    )