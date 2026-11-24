# FinSight — Final

FinSight is a local financial-document workspace. It extracts text and tables
from PDF, DOCX, XLSX, XLSM, and CSV files, saves reports and page annotations,
compares financial figures with grouped-bar and ratio-comparison charts,
calculates common ratios, exports selected pages, and asks local Ollama models
for cited answers and report summaries.

## Run it in VS Code

Open the `finsight` folder in VS Code. In **Terminal → New Terminal**, run:

```sh
cd backend
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
uvicorn main:app --reload
```

Leave that terminal open while using the site. Open **http://127.0.0.1:8000**
for the FinSight workspace. The API's interactive documentation is at
**http://127.0.0.1:8000/docs**.

If you already created and activated `.venv`, skip the environment creation
and activation commands, then run the install and server commands.

## Use the workspace

1. Choose or drop a PDF, DOCX, XLSX, XLSM, or CSV file that is 10 MB or smaller.
2. Select **Extract text**.
3. Review the extracted text and tables page by page; add notes or highlight text.
4. Ask the local model questions or generate a cited summary; select a page citation to open its source.
5. Compare two reports with financial charts and exact-value tables; search with report/page/date filters and review expanded ratios.
6. Export selected pages, notes, citations, and tables, or try the two fictional sample reports.
7. Reopen or delete previously extracted reports in **Saved reports**.

OCR requires Tesseract to be installed on the host and available on `PATH`.
AI features require Ollama running locally at `http://127.0.0.1:11434` and the
selected model to be installed; set `OLLAMA_URL` to use another address.
Reports, extracted tables, notes, highlights, and AI source citations are
stored in SQLite and are not sent to an external AI service.

## Run the tests

From `backend`, with `.venv` activated:

```sh
python -m pip install -r requirements-dev.txt
pytest
```

The tests cover the homepage, static assets, health check, supported document
imports, table extraction, annotations, search filters, report comparison,
citations, saved-report listing and deletion, and export behavior.

## Why this structure

- **FastAPI** serves both the website and the PDF extraction endpoint, so the
  browser can call the backend without a separate frontend server.
- **HTML, CSS, and JavaScript** keep this first interface simple to run and
  edit in VS Code; no Node.js setup is needed yet.
- **PyMuPDF**, **python-docx**, and **openpyxl** extract content and tables from supported formats.
- **Page-numbered results** keep the location of each extracted passage for
  later evidence citations.
- **A 10 MB upload limit** keeps this starter's document handling bounded.
