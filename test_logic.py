"""Ad-hoc verification script for app.py's non-Streamlit logic. Not a
formal pytest suite -- just exercises each pure function with mocked
inputs so we don't need a live Lusha API key to sanity-check the app."""

import io
import pandas as pd
import app


class UploadedFileStub(io.BytesIO):
    """Mimics the bits of Streamlit's UploadedFile that app.py touches:
    a real file-like object (so pd.read_excel works unmodified) plus a
    .name attribute."""

    def __init__(self, path, name=None):
        with open(path, "rb") as fh:
            super().__init__(fh.read())
        self.name = name or path


def test_parse_valid_file():
    df = app.parse_uploaded_file(UploadedFileStub("sample_leads.xlsx"))
    assert list(df.columns) == ["Full Name", "Company Name", "Role"], df.columns
    assert len(df) == 4, f"expected 4 usable rows, got {len(df)}\n{df}"
    print("PASS: parse_uploaded_file (valid file) ->", len(df), "rows")


def test_parse_missing_columns():
    bad_df = pd.DataFrame({"Name": ["x"], "Company": ["y"]})
    bad_df.to_excel("bad_columns.xlsx", index=False)
    try:
        app.parse_uploaded_file(UploadedFileStub("bad_columns.xlsx"))
        raise AssertionError("expected FileValidationError")
    except app.FileValidationError as e:
        assert "missing required column" in str(e).lower()
        print("PASS: parse_uploaded_file (missing columns) ->", e)


def test_parse_bad_extension():
    class FakeUploadTxt:
        name = "leads.txt"

    try:
        app.parse_uploaded_file(FakeUploadTxt())
        raise AssertionError("expected FileValidationError")
    except app.FileValidationError as e:
        print("PASS: parse_uploaded_file (bad extension) ->", e)


def test_parse_no_file():
    try:
        app.parse_uploaded_file(None)
        raise AssertionError("expected FileValidationError")
    except app.FileValidationError as e:
        print("PASS: parse_uploaded_file (no file) ->", e)


def test_split_full_name():
    assert app.split_full_name("Jane Smith") == ("Jane", "Smith")
    assert app.split_full_name("Madonna") == ("Madonna", "")
    assert app.split_full_name("Mary Jane Watson") == ("Mary", "Jane Watson")
    assert app.split_full_name("   ") == ("", "")
    print("PASS: split_full_name")


def test_parse_lusha_response_v3_shape():
    data = {
        "results": [
            {
                "clientReferenceId": "row-1",
                "emails": [{"email": "jane@acme.com", "type": "work"}],
                "phones": [{"number": "+1-555-0100"}],
            }
        ]
    }
    result = app.parse_lusha_response(data)
    assert result == {"phone": "+1-555-0100", "email": "jane@acme.com"}, result
    print("PASS: parse_lusha_response (v3 shape) ->", result)


def test_parse_lusha_response_not_found():
    data = {"results": [{"clientReferenceId": "row-1", "error": {"code": "NOT_FOUND"}}]}
    result = app.parse_lusha_response(data)
    assert result == {"phone": app.NOT_FOUND, "email": app.NOT_FOUND}, result
    print("PASS: parse_lusha_response (not found) ->", result)


def test_parse_lusha_response_empty():
    assert app.parse_lusha_response({}) == {"phone": app.NOT_FOUND, "email": app.NOT_FOUND}
    assert app.parse_lusha_response({"results": []}) == {"phone": app.NOT_FOUND, "email": app.NOT_FOUND}
    print("PASS: parse_lusha_response (empty/garbage input)")


def test_enrich_leads_handles_api_failure(monkeypatch):
    df = pd.DataFrame(
        {
            "Full Name": ["Jane Smith", "No Company"],
            "Company Name": ["Acme Corp", ""],
            "Role": ["VP Sales", "Analyst"],
        }
    )

    def fake_call(first, last, company, key):
        raise app.LushaAPIError("simulated auth failure")

    monkeypatch_target = app.call_lusha_api
    app.call_lusha_api = fake_call
    try:
        leads = app.enrich_leads(df, "fake-key")
    finally:
        app.call_lusha_api = monkeypatch_target

    assert len(leads) == 2
    assert leads[0].phone_number.startswith("Error:")
    assert leads[1].phone_number == app.NOT_FOUND  # no company -> skipped, not even called
    print("PASS: enrich_leads handles API failures without crashing")


def test_csv_export():
    leads = [
        app.EnrichedLead("Jane Smith", "Acme Corp", "VP Sales", "+1-555-0100", "jane@acme.com"),
        app.EnrichedLead("No Match", "Nowhere Inc", "Intern", app.NOT_FOUND, app.NOT_FOUND),
    ]
    csv_bytes = app.leads_to_csv_bytes(leads)
    text = csv_bytes.decode("utf-8")
    assert "Full Name,Company Name,Role,Phone Number,Email" in text
    assert "jane@acme.com" in text
    assert "Not found" in text
    print("PASS: leads_to_csv_bytes\n---\n" + text + "---")


def test_card_html_renders_and_escapes():
    lead = app.EnrichedLead("<script>alert(1)</script>", "Acme & Co", "R&D", "555", "a@b.com")
    html = app.render_lead_card_html(lead)
    assert "<script>alert(1)</script>" not in html  # must be escaped
    assert "&amp;" in html
    assert "555" in html and "a@b.com" in html
    print("PASS: render_lead_card_html escapes and includes fields")


class _FakeMonkeypatch:
    def setattr(self, *a, **k):
        pass


if __name__ == "__main__":
    test_parse_valid_file()
    test_parse_missing_columns()
    test_parse_bad_extension()
    test_parse_no_file()
    test_split_full_name()
    test_parse_lusha_response_v3_shape()
    test_parse_lusha_response_not_found()
    test_parse_lusha_response_empty()
    test_enrich_leads_handles_api_failure(_FakeMonkeypatch())
    test_csv_export()
    test_card_html_renders_and_escapes()
    print("\nALL TESTS PASSED")
