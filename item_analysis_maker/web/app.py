from __future__ import annotations

import re
import tempfile
import zipfile
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.background import BackgroundTask

from ..analysis import analyze, combine_datasets
from ..outputs import generate_outputs, safe_part
from ..parsing import InputFormatError, parse_csv_batch
from ..programs import PROGRAM_CHAIRS, PROGRAM_LONG_NAMES

WEB_DIR = Path(__file__).resolve().parent
PROJECT_DIR = WEB_DIR.parents[1]
TEMPLATE_PATH = PROJECT_DIR / "ITEM_ANALYSIS_FORMAT.xlsx"
MAX_INPUT_BYTES = 20 * 1024 * 1024
EXAM_TYPES = ("Prelims", "Midterms", "Finals")
SEMESTERS = ("1st Sem", "2nd Sem", "Term Break")

app = FastAPI(title="ItemAnalysisMaker")
app.mount("/static", StaticFiles(directory=WEB_DIR / "static"), name="static")
templates = Jinja2Templates(directory=WEB_DIR / "templates")


@app.middleware("http")
async def limit_request_size(request: Request, call_next):
    content_length = request.headers.get("content-length")
    if content_length:
        try:
            if int(content_length) > MAX_INPUT_BYTES + 1024 * 1024:
                return JSONResponse({"detail": "Request is too large. Combined CSV input is limited to 20 MB."}, status_code=413)
        except ValueError:
            return JSONResponse({"detail": "Invalid request size."}, status_code=400)
    return await call_next(request)


def _safe_filename(value: str) -> str:
    value = Path(value).name
    return re.sub(r"[^A-Za-z0-9_-]+", "_", Path(value).stem).strip("_-")[:60] or "Class"


def _render(request: Request, *, message: str = "", error: bool = False, values: dict | None = None) -> HTMLResponse:
    values = values or {}
    return templates.TemplateResponse(request=request, name="index.html", context={
        "message": message, "is_error": error, "values": values,
        "exam_types": EXAM_TYPES, "semesters": SEMESTERS,
        "program_chairs": PROGRAM_CHAIRS, "program_long_names": PROGRAM_LONG_NAMES,
    })


def _metadata(exam_type, academic_year, semester, subject, description, prepared_by, department_chairperson, program=""):
    if exam_type not in EXAM_TYPES:
        raise ValueError("Choose Prelims, Midterms, or Finals.")
    if semester not in SEMESTERS:
        raise ValueError("Choose a valid semester.")
    if not subject.strip():
        raise ValueError("Enter a subject.")
    return {
        "exam_type": exam_type, "academic_year": academic_year.strip(), "semester": semester,
        "subject": subject.strip(), "description": description.strip(),
        "prepared_by": prepared_by.strip(), "department_chairperson": department_chairperson.strip(),
        "program": program,
    }


async def _read_batch(files: list[UploadFile]):
    selected = [upload for upload in files if upload.filename]
    if not selected:
        raise InputFormatError("Choose at least one class LMS or ZipGrade CSV.")
    total = 0
    labels = []
    with tempfile.TemporaryDirectory(prefix="item-analysis-input-") as directory:
        paths = []
        for index, upload in enumerate(selected, start=1):
            filename = Path(upload.filename).name
            if Path(filename).suffix.casefold() != ".csv":
                raise InputFormatError(f"{filename}: only .csv files are supported.")
            content = await upload.read(MAX_INPUT_BYTES + 1)
            total += len(content)
            if total > MAX_INPUT_BYTES:
                raise InputFormatError("Combined CSV input is too large. The limit is 20 MB.")
            label = _safe_filename(filename)
            path = Path(directory) / f"{index:03d}_{label}.csv"
            path.write_bytes(content)
            paths.append(path)
            labels.append(filename)
        parsed = parse_csv_batch(paths)
        combined = combine_datasets(tuple(data for _, data in parsed))
    return combined, tuple(labels)


async def _close_uploads(files):
    for upload in files:
        await upload.close()


@app.get("/", response_class=HTMLResponse)
async def homepage(request: Request):
    return _render(request)


@app.post("/validate", response_class=HTMLResponse)
async def validate(
    request: Request,
    files: list[UploadFile] = File(default=[]),
    exam_type: str = Form("Prelims"),
    academic_year: str = Form(""),
    semester: str = Form("1st Sem"),
    subject: str = Form(""),
    description: str = Form(""),
    prepared_by: str = Form(""),
    department_chairperson: str = Form(""),
    program: str = Form(""),
):
    values = {"exam_type": exam_type, "academic_year": academic_year, "semester": semester,
              "subject": subject, "description": description, "prepared_by": prepared_by,
              "department_chairperson": department_chairperson, "program": program}
    try:
        if program and program not in PROGRAM_CHAIRS:
            raise ValueError("Choose a valid program.")
        if program and not department_chairperson.strip():
            department_chairperson = PROGRAM_CHAIRS[program]
            values["department_chairperson"] = department_chairperson
        _metadata(**values)
        data, labels = await _read_batch(files)
        result = analyze(data)
        message = (f"Validated {len(labels)} class CSVs for one course: {len(data.records)} eligible students, "
                   f"{data.question_count} items. Pooled high group: {len(result.high_group)}; "
                   f"low group: {len(result.low_group)}.")
        return _render(request, message=message, values=values)
    except (ValueError, InputFormatError) as exc:
        return _render(request, message=str(exc), error=True, values=values)
    except Exception:
        return _render(request, message="CSV validation failed. Check the selected exports and try again.",
                       error=True, values=values)
    finally:
        await _close_uploads(files)


@app.post("/generate")
async def generate(
    request: Request,
    files: list[UploadFile] = File(default=[]),
    exam_type: str = Form("Prelims"),
    academic_year: str = Form(""),
    semester: str = Form("1st Sem"),
    subject: str = Form(""),
    description: str = Form(""),
    prepared_by: str = Form(""),
    department_chairperson: str = Form(""),
    program: str = Form(""),
):
    temp_dir = None
    try:
        if program and program not in PROGRAM_CHAIRS:
            raise ValueError("Choose a valid program.")
        if program and not department_chairperson.strip():
            department_chairperson = PROGRAM_CHAIRS[program]
        metadata = _metadata(exam_type, academic_year, semester, subject, description,
                             prepared_by, department_chairperson, program)
        data, labels = await _read_batch(files)
        result = analyze(data)
        if not TEMPLATE_PATH.is_file():
            raise RuntimeError("The Excel template is unavailable on the server.")
        temp_dir = tempfile.TemporaryDirectory(prefix="item-analysis-output-")
        output_root = Path(temp_dir.name)
        xlsx_path, docx_path = generate_outputs(TEMPLATE_PATH, output_root, result, metadata)
        archive_path = output_root / "Item_Analysis_Package.zip"
        with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(xlsx_path, arcname=xlsx_path.name)
            archive.write(docx_path, arcname=docx_path.name)
        filename = f"Item_Analysis_{safe_part(metadata['subject'], 'Subject')}_{safe_part(metadata['exam_type'], 'Exam')}.zip"
        return FileResponse(archive_path, media_type="application/zip", filename=filename,
                            background=BackgroundTask(temp_dir.cleanup))
    except (ValueError, InputFormatError) as exc:
        if temp_dir is not None:
            temp_dir.cleanup()
        return _render(request, message=str(exc), error=True, values={
            "exam_type": exam_type, "academic_year": academic_year, "semester": semester,
            "subject": subject, "description": description, "prepared_by": prepared_by,
            "department_chairperson": department_chairperson,
            "program": program,
        })
    except Exception:
        if temp_dir is not None:
            temp_dir.cleanup()
        return _render(request, message="Generation failed. Check the CSVs and try again.", error=True, values={
            "exam_type": exam_type, "academic_year": academic_year, "semester": semester,
            "subject": subject, "description": description, "prepared_by": prepared_by,
            "department_chairperson": department_chairperson,
            "program": program,
        })
    finally:
        await _close_uploads(files)
