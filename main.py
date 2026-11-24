"""FinSight document extraction and local report storage."""

import csv
import json
import io
import os
import re
import shutil
import sqlite3
import zipfile
from pathlib import Path
from typing import Literal

import httpx
import fitz
import pytesseract
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from PIL import Image
from docx import Document as WordDocument
from docx.opc.exceptions import PackageNotFoundError
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from pydantic import BaseModel, Field

app = FastAPI(title="FinSight API", version="2.0.0")

STATIC_DIR = Path(__file__).parent / "static"
MAX_PDF_BYTES = 10 * 1024 * 1024
MAX_UPLOAD_BYTES = MAX_PDF_BYTES
DB_PATH = Path(os.environ.get("FINSIGHT_DB_PATH", Path(__file__).parent / "finsight.db"))
OLLAMA_URL = os.environ.get("OLLAMA_URL", "http://127.0.0.1:11434")
SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".xlsx", ".xlsm", ".csv"}

app.mount("/assets", StaticFiles(directory=STATIC_DIR), name="assets")


def database_connection() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def initialize_database() -> None:
    with database_connection() as connection:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                page_count INTEGER NOT NULL,
                source_size INTEGER NOT NULL,
                source_format TEXT NOT NULL DEFAULT 'pdf',
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS pages (
                document_id INTEGER NOT NULL,
                page_number INTEGER NOT NULL,
                text TEXT NOT NULL,
                PRIMARY KEY (document_id, page_number),
                FOREIGN KEY (document_id) REFERENCES documents(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS document_tables (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                table_number INTEGER NOT NULL,
                page_number INTEGER NOT NULL,
                title TEXT NOT NULL,
                rows_json TEXT NOT NULL,
                FOREIGN KEY (document_id) REFERENCES documents(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS annotations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                page_number INTEGER NOT NULL,
                kind TEXT NOT NULL CHECK (kind IN ('note', 'highlight')),
                selected_text TEXT NOT NULL DEFAULT '',
                body TEXT NOT NULL DEFAULT '',
                color TEXT NOT NULL DEFAULT 'yellow',
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (document_id) REFERENCES documents(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS ai_citations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                kind TEXT NOT NULL,
                prompt TEXT NOT NULL,
                answer TEXT NOT NULL,
                refs_json TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (document_id) REFERENCES documents(id) ON DELETE CASCADE
            );
            """
        )
        columns = {
            row["name"] for row in connection.execute("PRAGMA table_info(documents)")
        }
        if "source_format" not in columns:
            connection.execute(
                "ALTER TABLE documents ADD COLUMN source_format TEXT NOT NULL DEFAULT 'pdf'"
            )


initialize_database()


@app.get("/", include_in_schema=False)
def home() -> FileResponse:
    """Serve FinSight's web interface."""
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/health")
def health() -> dict[str, str]:
    """Simple check that the API process is responding."""
    return {"status": "ok"}


def _read_pdf(contents: bytes, ocr: bool) -> tuple[list[str], list[dict]]:
    try:
        document = fitz.open(stream=contents, filetype="pdf")
        pages = [page.get_text("text") for page in document]
        tables = []
        for page_number, page in enumerate(document, 1):
            found = page.find_tables()
            for table in found.tables:
                rows = table.extract()
                if rows:
                    tables.append({
                        "page_number": page_number,
                        "title": f"Page {page_number} table",
                        "rows": rows,
                    })
        if ocr and not shutil.which("tesseract"):
            raise HTTPException(
                status_code=503,
                detail="OCR is unavailable. Install Tesseract OCR and restart FinSight.",
            )
        if ocr:
            pages = []
            for page in document:
                pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                image = Image.open(io.BytesIO(pixmap.tobytes("png")))
                pages.append(pytesseract.image_to_string(image))
        document.close()
        return pages, tables
    except (fitz.FileDataError, ValueError, zipfile.BadZipFile):
        raise HTTPException(status_code=400, detail="The PDF could not be read.") from None


def _read_upload(filename: str, contents: bytes, ocr: bool) -> tuple[list[str], list[dict], str]:
    extension = Path(filename).suffix.lower()
    if extension == ".pdf":
        if not contents.startswith(b"%PDF-"):
            raise HTTPException(status_code=400, detail="The uploaded file is not a valid PDF.")
        pages, tables = _read_pdf(contents, ocr)
        return pages, tables, "pdf"
    if ocr:
        raise HTTPException(status_code=400, detail="OCR is only available for PDF files.")
    if extension == ".docx":
        try:
            document = WordDocument(io.BytesIO(contents))
        except (ValueError, zipfile.BadZipFile, PackageNotFoundError):
            raise HTTPException(status_code=400, detail="The Word document could not be read.") from None
        text = "\n".join(paragraph.text for paragraph in document.paragraphs if paragraph.text)
        tables = []
        for table_number, table in enumerate(document.tables, 1):
            rows = [[cell.text for cell in row.cells] for row in table.rows]
            tables.append({"page_number": 1, "title": f"Table {table_number}", "rows": rows})
            text = f"{text}\n\n" + "\n".join("\t".join(row) for row in rows)
        return [text], tables, "docx"
    if extension in {".xlsx", ".xlsm"}:
        try:
            workbook = load_workbook(io.BytesIO(contents), read_only=True, data_only=True)
        except (ValueError, zipfile.BadZipFile, InvalidFileException):
            raise HTTPException(status_code=400, detail="The spreadsheet could not be read.") from None
        pages = []
        tables = []
        for page_number, sheet in enumerate(workbook.worksheets, 1):
            rows = [[cell for cell in row] for row in sheet.iter_rows(values_only=True)]
            normalized = [[("" if cell is None else str(cell)) for cell in row] for row in rows]
            pages.append("\n".join("\t".join(row) for row in normalized))
            if normalized:
                tables.append({
                    "page_number": page_number,
                    "title": sheet.title,
                    "rows": normalized,
                })
        workbook.close()
        return pages, tables, extension.lstrip(".")
    if extension == ".csv":
        try:
            rows = list(csv.reader(io.StringIO(contents.decode("utf-8-sig"))))
        except UnicodeDecodeError:
            raise HTTPException(status_code=400, detail="The CSV file must use UTF-8 text.") from None
        text = "\n".join("\t".join(row) for row in rows)
        return [text], [{"page_number": 1, "title": "CSV data", "rows": rows}], "csv"
    raise HTTPException(
        status_code=415,
        detail="Upload a PDF, DOCX, XLSX, XLSM, or CSV file.",
    )


def _save_document(
    filename: str,
    contents: bytes,
    pages: list[str],
    tables: list[dict],
    source_format: str,
) -> int:
    with database_connection() as connection:
        cursor = connection.execute(
            "INSERT INTO documents (filename, page_count, source_size, source_format) "
            "VALUES (?, ?, ?, ?)",
            (filename, len(pages), len(contents), source_format),
        )
        document_id = cursor.lastrowid
        connection.executemany(
            "INSERT INTO pages (document_id, page_number, text) VALUES (?, ?, ?)",
            [(document_id, number, text) for number, text in enumerate(pages, 1)],
        )
        connection.executemany(
            "INSERT INTO document_tables (document_id, table_number, page_number, title, rows_json) "
            "VALUES (?, ?, ?, ?, ?)",
            [
                (document_id, number, table["page_number"], table["title"], json.dumps(table["rows"]))
                for number, table in enumerate(tables, 1)
            ],
        )
    return document_id


@app.post("/documents/extract")
async def extract_document(file: UploadFile = File(...), ocr: bool = False) -> dict:
    """Extract page/sheet text and tables from supported local document formats."""
    filename = Path(file.filename or "upload").name
    extension = Path(filename).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise HTTPException(status_code=415, detail="Upload a PDF, DOCX, XLSX, XLSM, or CSV file.")
    contents = await file.read(MAX_UPLOAD_BYTES + 1)
    if len(contents) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="Files must be 10 MB or smaller.")
    pages, tables, source_format = _read_upload(filename, contents, ocr)
    document_id = _save_document(filename, contents, pages, tables, source_format)
    extracted_pages = [
        {"page_number": number, "text": text} for number, text in enumerate(pages, 1)
    ]
    return {
        "document_id": document_id,
        "filename": filename,
        "source_format": source_format,
        "page_count": len(pages),
        "pages": extracted_pages,
        "tables": tables,
    }


@app.get("/documents")
def list_documents() -> dict:
    """List saved reports without loading their full extracted text."""
    with database_connection() as connection:
        rows = connection.execute(
            "SELECT id, filename, page_count, source_size, source_format, created_at "
            "FROM documents ORDER BY id DESC"
        ).fetchall()
    return {"documents": [dict(row) for row in rows]}


@app.get("/documents/{document_id:int}")
def get_document(document_id: int) -> dict:
    """Get a saved report and all of its source-page text."""
    with database_connection() as connection:
        document = connection.execute(
            "SELECT id, filename, page_count, source_size, source_format, created_at "
            "FROM documents WHERE id = ?",
            (document_id,),
        ).fetchone()
        if document is None:
            raise HTTPException(status_code=404, detail="Saved report not found.")
        page_rows = connection.execute(
            "SELECT page_number, text FROM pages "
            "WHERE document_id = ? ORDER BY page_number",
            (document_id,),
        ).fetchall()
    result = dict(document)
    result["pages"] = [dict(row) for row in page_rows]
    return result


@app.delete("/documents/{document_id}")
def delete_document(document_id: int) -> dict:
    """Permanently remove a saved report from this local workspace."""
    with database_connection() as connection:
        connection.execute("DELETE FROM pages WHERE document_id = ?", (document_id,))
        cursor = connection.execute("DELETE FROM documents WHERE id = ?", (document_id,))
    if cursor.rowcount == 0:
        raise HTTPException(status_code=404, detail="Saved report not found.")
    return {"deleted": True, "document_id": document_id}


@app.get("/documents/search")
def search_documents(
    q: str = "",
    document_id: int | None = None,
    page_number: int | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
) -> dict:
    """Search saved report names and extracted page text."""
    query = q.strip()
    if not query:
        return {"documents": [], "matches": []}
    filters = ["(d.filename LIKE ? OR d.id IN (SELECT document_id FROM pages WHERE text LIKE ?))"]
    values: list[object] = [f"%{query}%", f"%{query}%"]
    if document_id is not None:
        filters.append("d.id = ?")
        values.append(document_id)
    if start_date:
        filters.append("date(d.created_at) >= date(?)")
        values.append(start_date)
    if end_date:
        filters.append("date(d.created_at) <= date(?)")
        values.append(end_date)
    if page_number is not None:
        filters.append(
            "d.id IN (SELECT document_id FROM pages WHERE page_number = ? AND text LIKE ?)"
        )
        values.extend([page_number, f"%{query}%"])
    where_clause = " AND ".join(filters)
    with database_connection() as connection:
        documents = connection.execute(
            "SELECT d.id, d.filename, d.page_count, d.source_size, d.source_format, d.created_at "
            f"FROM documents d WHERE {where_clause} ORDER BY d.id DESC",
            values,
        ).fetchall()
        match_filters = ["p.text LIKE ?"]
        match_values: list[object] = [f"%{query}%"]
        if document_id is not None:
            match_filters.append("p.document_id = ?")
            match_values.append(document_id)
        if page_number is not None:
            match_filters.append("p.page_number = ?")
            match_values.append(page_number)
        if start_date:
            match_filters.append("date(d.created_at) >= date(?)")
            match_values.append(start_date)
        if end_date:
            match_filters.append("date(d.created_at) <= date(?)")
            match_values.append(end_date)
        matches = connection.execute(
            "SELECT p.document_id, d.filename, p.page_number, p.text FROM pages p "
            "JOIN documents d ON d.id = p.document_id WHERE "
            + " AND ".join(match_filters)
            + " ORDER BY p.document_id DESC, p.page_number",
            match_values,
        ).fetchall()
    return {
        "documents": [dict(row) for row in documents],
        "matches": [dict(row) for row in matches],
    }


def extract_financial_value(text: str, labels: tuple[str, ...]) -> float | None:
    ordered_labels = sorted(labels, key=len, reverse=True)
    label_pattern = "|".join(re.escape(label) for label in ordered_labels)
    amount_pattern = re.compile(
        r"([$£€]?\s*\(?[-+]?\d[\d,]*(?:\.\d+)?\)?)"
    )
    label_matcher = re.compile(rf"(?<!\w)(?:{label_pattern})\b", re.IGNORECASE)
    candidates = []
    lines = text.splitlines()
    for line_number, line in enumerate(lines):
        label_match = label_matcher.search(line)
        if not label_match:
            continue
        amount_match = amount_pattern.search(line[label_match.end():])
        if not amount_match:
            for following_line in lines[line_number + 1:line_number + 3]:
                if label_matcher.search(following_line):
                    break
                amount_match = amount_pattern.search(following_line)
                if amount_match and not following_line[:amount_match.start()].strip():
                    break
                amount_match = None
        if not amount_match:
            continue
        token = amount_match.group(1).strip()
        negative = token.startswith("(") and token.endswith(")")
        number = re.sub(r"[$£€,()\s]", "", token)
        try:
            value = float(number)
        except ValueError:
            continue
        if negative:
            value = -value
        formatted = int(any(char in token for char in ",.$£€"))
        candidates.append((formatted, abs(value), value))
    return max(candidates)[2] if candidates else None


@app.get("/documents/{document_id}/ratios")
def document_ratios(document_id: int) -> dict:
    """Calculate common ratios from recognizable labels in extracted text."""
    with database_connection() as connection:
        document = connection.execute(
            "SELECT filename FROM documents WHERE id = ?",
            (document_id,),
        ).fetchone()
        if document is None:
            raise HTTPException(status_code=404, detail="Saved report not found.")
        text = "\n".join(
            row["text"]
            for row in connection.execute(
                "SELECT text FROM pages WHERE document_id = ? ORDER BY page_number",
                (document_id,),
            )
        )
    values = {
        "assets": extract_financial_value(text, ("total assets", "assets")),
        "liabilities": extract_financial_value(text, ("total liabilities", "liabilities")),
        "equity": extract_financial_value(text, ("total equity", "shareholders equity", "stockholders equity")),
        "revenue": extract_financial_value(text, ("total revenue", "net revenue", "revenue", "sales")),
        "net_income": extract_financial_value(text, ("net income", "profit")),
        "current_assets": extract_financial_value(text, ("total current assets", "current assets")),
        "current_liabilities": extract_financial_value(text, ("total current liabilities", "current liabilities")),
        "gross_profit": extract_financial_value(text, ("gross profit",)),
        "operating_income": extract_financial_value(text, ("operating income", "income from operations")),
        "cash_from_operations": extract_financial_value(
            text, ("cash from operating activities", "operating cash flow", "cash flow from operations")
        ),
        "capital_expenditures": extract_financial_value(
            text, ("capital expenditures", "capital expenditure", "property plant and equipment purchases")
        ),
    }
    ratios = {}
    if values["liabilities"] is not None and values["assets"]:
        ratios["debt_to_assets"] = round(values["liabilities"] / values["assets"], 3)
    if values["net_income"] is not None and values["revenue"]:
        ratios["net_profit_margin"] = round(values["net_income"] / values["revenue"], 3)
    if values["equity"] is not None and values["assets"]:
        ratios["equity_to_assets"] = round(values["equity"] / values["assets"], 3)
    if values["gross_profit"] is not None and values["revenue"]:
        ratios["gross_margin"] = round(values["gross_profit"] / values["revenue"], 3)
    if values["operating_income"] is not None and values["revenue"]:
        ratios["operating_margin"] = round(values["operating_income"] / values["revenue"], 3)
    if values["current_assets"] is not None and values["current_liabilities"]:
        ratios["current_ratio"] = round(values["current_assets"] / values["current_liabilities"], 3)
    if values["cash_from_operations"] is not None and values["current_liabilities"]:
        ratios["operating_cash_flow_ratio"] = round(
            values["cash_from_operations"] / values["current_liabilities"], 3
        )
    if values["cash_from_operations"] is not None and values["capital_expenditures"] is not None:
        values["free_cash_flow"] = values["cash_from_operations"] - values["capital_expenditures"]
        if values["revenue"]:
            ratios["free_cash_flow_margin"] = round(values["free_cash_flow"] / values["revenue"], 3)
    explanations = {
        "debt_to_assets": "Liabilities as a share of assets; lower can indicate less balance-sheet leverage.",
        "net_profit_margin": "Net income retained from each unit of revenue.",
        "equity_to_assets": "The share of assets financed by shareholders' equity.",
        "gross_margin": "Revenue remaining after direct costs.",
        "operating_margin": "Operating income earned per unit of revenue.",
        "current_ratio": "Current assets available for each unit of current liabilities.",
        "operating_cash_flow_ratio": "Operating cash flow available for current liabilities.",
        "free_cash_flow_margin": "Operating cash after capital expenditures, as a share of revenue.",
    }
    return {
        "document_id": document_id,
        "filename": document["filename"],
        "values": values,
        "ratios": ratios,
        "explanations": {key: explanations[key] for key in ratios},
    }


def _document_pages(document_id: int) -> tuple[str, list[dict]]:
    with database_connection() as connection:
        document = connection.execute(
            "SELECT filename FROM documents WHERE id = ?",
            (document_id,),
        ).fetchone()
        if document is None:
            raise HTTPException(status_code=404, detail="Saved report not found.")
        page_rows = connection.execute(
            "SELECT page_number, text FROM pages "
            "WHERE document_id = ? ORDER BY page_number",
            (document_id,),
        ).fetchall()
    return document["filename"], [dict(row) for row in page_rows]


def _citation_refs(pages: list[dict], query: str) -> list[dict]:
    words = {
        word.lower()
        for word in re.findall(r"[A-Za-z0-9]+", query)
        if len(word) > 2
    }
    ranked = []
    for page in pages:
        text = page["text"]
        lowered = text.lower()
        score = sum(lowered.count(word) for word in words)
        if score:
            ranked.append((score, page))
    selected = [page for _, page in sorted(ranked, key=lambda item: item[0], reverse=True)[:4]]
    if not selected:
        selected = pages[: min(3, len(pages))]
    references = []
    for page in selected:
        lines = [line.strip() for line in page["text"].splitlines() if line.strip()]
        matching = [
            line for line in lines
            if not words or any(word in line.lower() for word in words)
        ]
        excerpt = " … ".join(matching[:3] or lines[:3])[:500]
        references.append({"page_number": page["page_number"], "excerpt": excerpt})
    return references


async def _ollama_answer(model: str, prompt: str) -> str:
    try:
        async with httpx.AsyncClient(timeout=120) as client:
            response = await client.post(
                f"{OLLAMA_URL.rstrip('/')}/api/chat",
                json={
                    "model": model,
                    "stream": False,
                    "messages": [{"role": "user", "content": prompt}],
                },
            )
    except httpx.RequestError:
        raise HTTPException(
            status_code=503,
            detail="Ollama is unavailable. Start Ollama and ensure the selected model is installed.",
        ) from None
    try:
        payload = response.json()
    except ValueError:
        raise HTTPException(status_code=502, detail="Ollama returned an invalid response.") from None
    if response.status_code >= 400:
        detail = payload.get("error") if isinstance(payload, dict) else None
        raise HTTPException(status_code=502, detail=detail or "Ollama could not answer the question.")
    message = payload.get("message") if isinstance(payload, dict) else None
    content = message.get("content", "") if isinstance(message, dict) else ""
    answer = content.strip() if isinstance(content, str) else ""
    return answer or "Ollama returned no answer."


def _save_ai_citation(
    document_id: int,
    kind: str,
    prompt: str,
    answer: str,
    references: list[dict],
) -> None:
    with database_connection() as connection:
        connection.execute(
            "INSERT INTO ai_citations (document_id, kind, prompt, answer, refs_json) "
            "VALUES (?, ?, ?, ?, ?)",
            (document_id, kind, prompt, answer, json.dumps(references)),
        )


class QuestionRequest(BaseModel):
    question: str = Field(min_length=1, max_length=1000)
    model: str = Field(default="llama3.2", min_length=1, max_length=100)


@app.post("/documents/{document_id}/ask")
async def ask_document(document_id: int, request: QuestionRequest) -> dict:
    """Answer from a saved report and return page-level source citations."""
    _, pages = _document_pages(document_id)
    question = request.question.strip()
    references = _citation_refs(pages, question)
    context = "\n\n".join(
        f"[Page {ref['page_number']}]\n{ref['excerpt']}" for ref in references
    )
    prompt = (
        "Answer only from the report context. Cite factual statements with the exact source "
        "marker [Page N] using the page numbers shown. If the answer is not present, say so. "
        "Do not invent figures or page references.\n\n"
        f"{context}\n\nQuestion: {question}"
    )
    answer = await _ollama_answer(request.model, prompt)
    cited_numbers = {
        int(number) for number in re.findall(r"\[Page\s+(\d+)\]", answer, flags=re.IGNORECASE)
    }
    cited_refs = [ref for ref in references if ref["page_number"] in cited_numbers] or references
    _save_ai_citation(document_id, "answer", question, answer, cited_refs)
    return {"answer": answer, "model": request.model, "citations": cited_refs}


@app.post("/documents/{document_id}/summary")
async def summarize_document(document_id: int, request: QuestionRequest) -> dict:
    """Generate a cited executive summary, key figures, risks, and outlook."""
    _, pages = _document_pages(document_id)
    references = _citation_refs(
        pages,
        "revenue assets liabilities income cash risk risks outlook management",
    )
    context = "\n\n".join(
        f"[Page {ref['page_number']}]\n{ref['excerpt']}" for ref in references
    )
    prompt = (
        "Create a concise financial report review with four headings: Executive summary, Key "
        "figures, Risks, and Outlook. Use only the source context; label unavailable items "
        "clearly. Cite every factual statement with [Page N] from the provided page numbers. "
        "Do not give investment advice.\n\n"
        f"{context}"
    )
    answer = await _ollama_answer(request.model, prompt)
    cited_numbers = {
        int(number) for number in re.findall(r"\[Page\s+(\d+)\]", answer, flags=re.IGNORECASE)
    }
    cited_refs = [ref for ref in references if ref["page_number"] in cited_numbers] or references
    _save_ai_citation(document_id, "summary", "Executive summary", answer, cited_refs)
    return {"summary": answer, "model": request.model, "citations": cited_refs}


@app.get("/documents/{document_id}/tables")
def document_tables(document_id: int) -> dict:
    with database_connection() as connection:
        if connection.execute("SELECT id FROM documents WHERE id = ?", (document_id,)).fetchone() is None:
            raise HTTPException(status_code=404, detail="Saved report not found.")
        rows = connection.execute(
            "SELECT table_number, page_number, title, rows_json FROM document_tables "
            "WHERE document_id = ? ORDER BY table_number",
            (document_id,),
        ).fetchall()
    return {
        "tables": [
            {
                "table_number": row["table_number"],
                "page_number": row["page_number"],
                "title": row["title"],
                "rows": json.loads(row["rows_json"]),
            }
            for row in rows
        ]
    }


@app.get("/documents/{document_id}/tables/export")
def export_document_tables(document_id: int):
    result = document_tables(document_id)
    output = io.StringIO()
    writer = csv.writer(output)
    for table in result["tables"]:
        writer.writerow([table["title"]])
        writer.writerows(table["rows"])
        writer.writerow([])
    return PlainTextResponse(
        output.getvalue(),
        media_type="text/csv",
        headers={"Content-Disposition": 'attachment; filename="finsight-tables.csv"'},
    )


class AnnotationRequest(BaseModel):
    page_number: int = Field(ge=1)
    kind: Literal["note", "highlight"]
    selected_text: str = Field(default="", max_length=5000)
    body: str = Field(default="", max_length=5000)
    color: Literal["yellow", "green", "blue"] = "yellow"


@app.get("/documents/{document_id}/annotations")
def list_annotations(document_id: int) -> dict:
    with database_connection() as connection:
        if connection.execute("SELECT id FROM documents WHERE id = ?", (document_id,)).fetchone() is None:
            raise HTTPException(status_code=404, detail="Saved report not found.")
        rows = connection.execute(
            "SELECT id, page_number, kind, selected_text, body, color, created_at "
            "FROM annotations WHERE document_id = ? ORDER BY page_number, id",
            (document_id,),
        ).fetchall()
    return {"annotations": [dict(row) for row in rows]}


@app.post("/documents/{document_id}/annotations")
def create_annotation(document_id: int, request: AnnotationRequest) -> dict:
    if request.kind == "note" and not request.body.strip():
        raise HTTPException(status_code=422, detail="A note cannot be empty.")
    if request.kind == "highlight" and not request.selected_text.strip():
        raise HTTPException(status_code=422, detail="Select text to highlight.")
    with database_connection() as connection:
        if connection.execute(
            "SELECT 1 FROM documents WHERE id = ?", (document_id,)
        ).fetchone() is None:
            raise HTTPException(status_code=404, detail="Saved report not found.")
        if connection.execute(
            "SELECT 1 FROM pages WHERE document_id = ? AND page_number = ?",
            (document_id, request.page_number),
        ).fetchone() is None:
            raise HTTPException(status_code=404, detail="Report page not found.")
        cursor = connection.execute(
            "INSERT INTO annotations (document_id, page_number, kind, selected_text, body, color) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (
                document_id, request.page_number, request.kind, request.selected_text,
                request.body.strip(), request.color,
            ),
        )
        row = connection.execute(
            "SELECT id, page_number, kind, selected_text, body, color, created_at "
            "FROM annotations WHERE id = ?",
            (cursor.lastrowid,),
        ).fetchone()
    return {"annotation": dict(row)}


@app.delete("/documents/{document_id}/annotations/{annotation_id}")
def delete_annotation(document_id: int, annotation_id: int) -> dict:
    with database_connection() as connection:
        cursor = connection.execute(
            "DELETE FROM annotations WHERE id = ? AND document_id = ?",
            (annotation_id, document_id),
        )
    if cursor.rowcount == 0:
        raise HTTPException(status_code=404, detail="Annotation not found.")
    return {"deleted": True}


@app.get("/documents/compare")
def compare_documents(first_id: int, second_id: int) -> dict:
    if first_id == second_id:
        raise HTTPException(status_code=400, detail="Choose two different reports to compare.")
    first = document_ratios(first_id)
    second = document_ratios(second_id)
    metrics = {}
    for key in sorted(set(first["values"]) | set(second["values"])):
        before, after = first["values"].get(key), second["values"].get(key)
        if before is None and after is None:
            continue
        delta = after - before if before is not None and after is not None else None
        change_percent = (delta / abs(before) * 100) if delta is not None and before else None
        metrics[key] = {
            "first": before,
            "second": after,
            "change": delta,
            "change_percent": round(change_percent, 1) if change_percent is not None else None,
        }
    return {
        "first": {"id": first_id, "filename": first["filename"], "ratios": first["ratios"]},
        "second": {"id": second_id, "filename": second["filename"], "ratios": second["ratios"]},
        "metrics": metrics,
    }


@app.post("/documents/demo")
def create_demo_reports() -> dict:
    """Add two fictional annual reports so key workspace features can be explored."""
    samples = [
        ("FinSight demo annual report 2024.pdf", [
            "FINSIGHT DEMO COMPANY — FICTIONAL SAMPLE — 2024",
            "Revenue: 1200\nGross profit: 480\nOperating income: 144\nNet income: 96",
            "Total assets: 900\nTotal current assets: 360\nTotal liabilities: 450\nTotal current liabilities: 180\nTotal equity: 450",
            "Cash from operating activities: 130\nCapital expenditures: 50\nRevenue increased as subscription renewals improved. Risks include customer concentration and foreign exchange uncertainty. Management expects measured expansion.",
        ], [["Metric", "2024"], ["Revenue", "1200"], ["Gross profit", "480"], ["Net income", "96"]]),
        ("FinSight demo annual report 2025.pdf", [
            "FINSIGHT DEMO COMPANY — FICTIONAL SAMPLE — 2025",
            "Revenue: 1500\nGross profit: 660\nOperating income: 210\nNet income: 150",
            "Total assets: 1100\nTotal current assets: 440\nTotal liabilities: 500\nTotal current liabilities: 200\nTotal equity: 600",
            "Cash from operating activities: 190\nCapital expenditures: 65\nRevenue increased with stronger recurring sales. Risks include customer concentration and supplier costs. Management expects continued investment in product development.",
        ], [["Metric", "2025"], ["Revenue", "1500"], ["Gross profit", "660"], ["Net income", "150"]]),
    ]
    result = []
    for filename, pages, rows in samples:
        with database_connection() as connection:
            existing = connection.execute(
                "SELECT id FROM documents WHERE filename = ?", (filename,)
            ).fetchone()
        if existing:
            result.append({"document_id": existing["id"], "filename": filename})
            continue
        table_data = [{"page_number": 2, "title": "Sample financial figures", "rows": rows}]
        document_id = _save_document(
            filename, "\n".join(pages).encode(), pages, table_data, "demo"
        )
        result.append({"document_id": document_id, "filename": filename})
    return {"documents": result}


@app.get("/documents/{document_id}/export")
def export_document(
    document_id: int,
    format: Literal["txt", "json", "csv"] = "txt",
    pages: str | None = None,
    include_annotations: bool = False,
):
    """Export selected pages and optionally include notes, citations, and tables."""
    with database_connection() as connection:
        document = connection.execute(
            "SELECT id, filename, page_count, source_size, source_format, created_at "
            "FROM documents WHERE id = ?",
            (document_id,),
        ).fetchone()
        if document is None:
            raise HTTPException(status_code=404, detail="Saved report not found.")
        page_rows = connection.execute(
            "SELECT page_number, text FROM pages "
            "WHERE document_id = ? ORDER BY page_number",
            (document_id,),
        ).fetchall()
        annotations = connection.execute(
            "SELECT id, page_number, kind, selected_text, body, color, created_at "
            "FROM annotations WHERE document_id = ? ORDER BY page_number, id",
            (document_id,),
        ).fetchall()
        citations = connection.execute(
            "SELECT kind, prompt, answer, refs_json, created_at FROM ai_citations "
            "WHERE document_id = ? ORDER BY id",
            (document_id,),
        ).fetchall()
        tables = connection.execute(
            "SELECT table_number, page_number, title, rows_json FROM document_tables "
            "WHERE document_id = ? ORDER BY table_number",
            (document_id,),
        ).fetchall()
    selected_pages = None
    if pages is not None:
        try:
            selected_pages = {int(page.strip()) for page in pages.split(",") if page.strip()}
        except ValueError:
            raise HTTPException(
                status_code=400,
                detail="Page selection must be comma-separated numbers.",
            ) from None
        if not selected_pages or any(
            page < 1 or page > document["page_count"] for page in selected_pages
        ):
            raise HTTPException(status_code=400, detail="Select valid report pages to export.")
    if selected_pages is not None:
        page_rows = [row for row in page_rows if row["page_number"] in selected_pages]
    document_data = dict(document)
    document_data["pages"] = [dict(row) for row in page_rows]
    document_data["annotations"] = [
        dict(row) for row in annotations
        if include_annotations
        and (selected_pages is None or row["page_number"] in selected_pages)
    ]
    document_data["citations"] = []
    if include_annotations:
        for row in citations:
            refs = json.loads(row["refs_json"])
            if selected_pages is not None:
                refs = [ref for ref in refs if ref["page_number"] in selected_pages]
            document_data["citations"].append({
                "kind": row["kind"],
                "prompt": row["prompt"],
                "answer": row["answer"],
                "citations": refs,
                "created_at": row["created_at"],
            })
    document_data["tables"] = [
        {
            "table_number": row["table_number"],
            "page_number": row["page_number"],
            "title": row["title"],
            "rows": json.loads(row["rows_json"]),
        }
        for row in tables
        if selected_pages is None or row["page_number"] in selected_pages
    ]
    if format == "json":
        return PlainTextResponse(
            json.dumps(document_data, indent=2),
            media_type="application/json",
            headers={"Content-Disposition": f'attachment; filename="{document["filename"]}.json"'},
        )
    if format == "csv":
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(["page_number", "text"])
        writer.writerows((row["page_number"], row["text"]) for row in page_rows)
        if include_annotations:
            writer.writerow([])
            writer.writerow(["kind", "page_number", "selected_text", "body"])
            writer.writerows(
                (row["kind"], row["page_number"], row["selected_text"], row["body"])
                for row in document_data["annotations"]
            )
            for citation in document_data["citations"]:
                writer.writerow(["citation", citation["prompt"], citation["answer"]])
                for ref in citation["citations"]:
                    writer.writerow(["source", ref["page_number"], ref["excerpt"]])
            for table in document_data["tables"]:
                writer.writerow([])
                writer.writerow([table["title"]])
                writer.writerows(table["rows"])
        return PlainTextResponse(
            output.getvalue(),
            media_type="text/csv",
            headers={"Content-Disposition": f'attachment; filename="{document["filename"]}.csv"'},
        )
    report = "\n\n".join(f"Page {row['page_number']}\n{row['text']}" for row in page_rows)
    if include_annotations:
        if document_data["annotations"]:
            report += "\n\nNOTES AND HIGHLIGHTS\n" + "\n".join(
                f"Page {row['page_number']} [{row['kind']}]: {row['selected_text']} {row['body']}"
                for row in document_data["annotations"]
            )
        if document_data["citations"]:
            report += "\n\nAI ANSWERS AND SOURCES\n" + "\n".join(
                f"{item['kind']}: {item['answer']}\n"
                + "\n".join(
                    f"Page {ref['page_number']}: {ref['excerpt']}"
                    for ref in item["citations"]
                )
                for item in document_data["citations"]
            )
        if document_data["tables"]:
            report += "\n\nEXTRACTED TABLES\n" + "\n".join(
                f"{table['title']}\n" + "\n".join("\t".join(row) for row in table["rows"])
                for table in document_data["tables"]
            )
    return PlainTextResponse(
        report,
        media_type="text/plain",
        headers={"Content-Disposition": f'attachment; filename="{document["filename"]}.txt"'},
    )
