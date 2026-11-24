import io
import os
import tempfile
from pathlib import Path

import fitz
from docx import Document as WordDocument
from openpyxl import Workbook
from fastapi.testclient import TestClient

TEST_DB = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
TEST_DB.close()
os.environ["FINSIGHT_DB_PATH"] = TEST_DB.name

from main import app

client = TestClient(app)


def setup_function():
    Path(TEST_DB.name).unlink(missing_ok=True)
    import main

    main.initialize_database()


def pdf_bytes_with_text(text: str) -> bytes:
    document = fitz.open()
    page = document.new_page()
    page.insert_text((72, 72), text)
    return document.tobytes()


def test_health_returns_ok():
    response = client.get("/health")

    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_homepage_serves_the_fin_sight_workspace():
    response = client.get("/")

    assert response.status_code == 200
    assert "See the story" in response.text
    assert "Extract text" in response.text


def test_frontend_assets_are_available():
    css = client.get("/assets/styles.css")
    javascript = client.get("/assets/app.js")

    assert css.status_code == 200
    assert "Monochrome glass interface" in css.text
    assert javascript.status_code == 200
    assert "renderResults" in javascript.text


def test_extract_returns_text_by_page():
    response = client.post(
        "/documents/extract",
        files={"file": ("annual-report.pdf", pdf_bytes_with_text("Revenue: 100"), "application/pdf")},
    )

    assert response.status_code == 200
    result = response.json()
    assert result["filename"] == "annual-report.pdf"
    assert result["page_count"] == 1
    assert result["document_id"] > 0
    assert result["pages"] == [{"page_number": 1, "text": "Revenue: 100\n"}]


def test_extract_keeps_text_attached_to_each_source_page():
    document = fitz.open()
    first_page = document.new_page()
    first_page.insert_text((72, 72), "Revenue: 100")
    second_page = document.new_page()
    second_page.insert_text((72, 72), "Net income: 12")

    response = client.post(
        "/documents/extract",
        files={"file": ("two-pages.pdf", document.tobytes(), "application/pdf")},
    )

    assert response.status_code == 200
    assert response.json()["pages"] == [
        {"page_number": 1, "text": "Revenue: 100\n"},
        {"page_number": 2, "text": "Net income: 12\n"},
    ]


def test_extract_rejects_non_pdf_upload():
    response = client.post(
        "/documents/extract",
        files={"file": ("notes.txt", b"not a PDF", "text/plain")},
    )

    assert response.status_code == 415


def test_extracts_docx_and_preserves_its_tables():
    document = WordDocument()
    document.add_paragraph("Annual revenue: 240")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Metric"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Revenue"
    table.cell(1, 1).text = "240"
    contents = io.BytesIO()
    document.save(contents)

    response = client.post(
        "/documents/extract",
        files={"file": ("report.docx", contents.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
    )

    assert response.status_code == 200
    result = response.json()
    assert result["source_format"] == "docx"
    assert "Annual revenue: 240" in result["pages"][0]["text"]
    assert result["tables"][0]["rows"][1] == ["Revenue", "240"]


def test_extracts_spreadsheet_sheets_and_exports_tables():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Summary"
    sheet.append(["Metric", "Value"])
    sheet.append(["Revenue", 240])
    contents = io.BytesIO()
    workbook.save(contents)

    response = client.post(
        "/documents/extract",
        files={"file": ("report.xlsx", contents.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    document_id = response.json()["document_id"]
    tables = client.get(f"/documents/{document_id}/tables")
    exported = client.get(f"/documents/{document_id}/tables/export")

    assert response.status_code == 200
    assert response.json()["source_format"] == "xlsx"
    assert tables.json()["tables"][0]["rows"][1] == ["Revenue", "240"]
    assert "Revenue,240" in exported.text


def test_saved_report_can_be_listed_opened_and_deleted():
    extracted = client.post(
        "/documents/extract",
        files={"file": ("saved-report.pdf", pdf_bytes_with_text("Cash: 500"), "application/pdf")},
    )
    document_id = extracted.json()["document_id"]

    listing = client.get("/documents")
    assert listing.status_code == 200
    assert listing.json()["documents"][0]["filename"] == "saved-report.pdf"

    detail = client.get(f"/documents/{document_id}")
    assert detail.status_code == 200
    assert detail.json()["pages"] == [{"page_number": 1, "text": "Cash: 500\n"}]

    deleted = client.delete(f"/documents/{document_id}")
    assert deleted.status_code == 200
    assert deleted.json() == {"deleted": True, "document_id": document_id}
    assert client.get(f"/documents/{document_id}").status_code == 404


def test_document_search_lists_matching_saved_report():
    extracted = client.post(
        "/documents/extract",
        files={"file": ("quarterly-report.pdf", pdf_bytes_with_text("Revenue increased"), "application/pdf")},
    )
    assert extracted.status_code == 200

    response = client.get("/documents/search?q=Revenue")

    assert response.status_code == 200
    assert response.json()["documents"][0]["filename"] == "quarterly-report.pdf"
    assert response.json()["matches"][0]["text"] == "Revenue increased\n"


def test_search_can_filter_by_page_and_document():
    first = client.post(
        "/documents/extract",
        files={"file": ("first.pdf", pdf_bytes_with_text("Revenue increased"), "application/pdf")},
    ).json()["document_id"]
    second = client.post(
        "/documents/extract",
        files={"file": ("second.pdf", pdf_bytes_with_text("Costs increased"), "application/pdf")},
    ).json()["document_id"]

    response = client.get(
        f"/documents/search?q=increased&document_id={first}&page_number=1"
    )

    assert response.status_code == 200
    assert [row["document_id"] for row in response.json()["matches"]] == [first]
    assert second not in [row["id"] for row in response.json()["documents"]]


def test_ratios_are_calculated_from_extracted_labels():
    extracted = client.post(
        "/documents/extract",
        files={"file": ("ratios.pdf", pdf_bytes_with_text(
            "Total assets: 1000\nTotal liabilities: 400\nTotal equity: 600\nRevenue: 500\nNet income: 50"
        ), "application/pdf")},
    )
    document_id = extracted.json()["document_id"]

    response = client.get(f"/documents/{document_id}/ratios")

    assert response.status_code == 200
    assert response.json()["ratios"]["debt_to_assets"] == 0.4
    assert response.json()["ratios"]["net_profit_margin"] == 0.1


def test_financial_value_parser_ignores_small_contents_page_numbers():
    import main

    text = "Contents\nRevenue ........ 3\nFinancial statements\nRevenue\n1,200"

    assert main.extract_financial_value(text, ("revenue",)) == 1200


def test_export_returns_saved_report():
    extracted = client.post(
        "/documents/extract",
        files={"file": ("export.pdf", pdf_bytes_with_text("Cash: 125"), "application/pdf")},
    )
    document_id = extracted.json()["document_id"]

    response = client.get(f"/documents/{document_id}/export?format=txt")

    assert response.status_code == 200
    assert "Page 1" in response.text
    assert "Cash: 125" in response.text


def test_annotations_are_page_bound_and_can_be_included_in_selected_export():
    document = fitz.open()
    document.new_page().insert_text((72, 72), "Revenue: 120")
    document.new_page().insert_text((72, 72), "Net income: 24")
    extracted = client.post(
        "/documents/extract",
        files={"file": ("annotated.pdf", document.tobytes(), "application/pdf")},
    )
    document_id = extracted.json()["document_id"]
    note = client.post(
        f"/documents/{document_id}/annotations",
        json={"page_number": 2, "kind": "note", "body": "Check the net margin."},
    )
    highlight = client.post(
        f"/documents/{document_id}/annotations",
        json={"page_number": 2, "kind": "highlight", "selected_text": "Net income: 24"},
    )
    exported = client.get(
        f"/documents/{document_id}/export?format=json&pages=2&include_annotations=true"
    )

    assert note.status_code == 200
    assert highlight.status_code == 200
    payload = exported.json()
    assert [page["page_number"] for page in payload["pages"]] == [2]
    assert {item["kind"] for item in payload["annotations"]} == {"note", "highlight"}


def test_demo_reports_support_comparison_and_financial_ratios():
    demo = client.post("/documents/demo")
    first_id, second_id = [item["document_id"] for item in demo.json()["documents"]]

    comparison = client.get(
        f"/documents/compare?first_id={first_id}&second_id={second_id}"
    )
    ratios = client.get(f"/documents/{second_id}/ratios")

    assert comparison.status_code == 200
    assert comparison.json()["metrics"]["revenue"]["change_percent"] == 25.0
    assert ratios.json()["ratios"]["gross_margin"] == 0.44
    assert "gross_margin" in ratios.json()["explanations"]


def test_ai_answer_returns_and_persists_page_citations(monkeypatch):
    extracted = client.post(
        "/documents/extract",
        files={"file": ("cited.pdf", pdf_bytes_with_text("Revenue grew 15 percent"), "application/pdf")},
    )
    document_id = extracted.json()["document_id"]

    async def fake_answer(model, prompt):
        assert "[Page 1]" in prompt
        return "Revenue grew 15 percent. [Page 1]"

    import main

    monkeypatch.setattr(main, "_ollama_answer", fake_answer)
    response = client.post(
        f"/documents/{document_id}/ask",
        json={"question": "What happened to revenue?", "model": "test-model"},
    )
    summary = client.post(
        f"/documents/{document_id}/summary",
        json={"question": "summary", "model": "test-model"},
    )
    exported = client.get(
        f"/documents/{document_id}/export?format=json&include_annotations=true"
    )

    assert response.status_code == 200
    assert response.json()["citations"][0]["page_number"] == 1
    assert summary.status_code == 200
    assert summary.json()["citations"][0]["page_number"] == 1
    assert len(exported.json()["citations"]) == 2
    assert exported.json()["citations"][0]["citations"][0]["page_number"] == 1
