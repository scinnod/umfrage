"""Tests for the umfrage web server.

Requires the optional web dependencies (umfrage[web]).  Tests are
automatically skipped when flask or flask_limiter are not installed.
"""

from __future__ import annotations

import io
import zipfile

import pytest
import yaml

flask = pytest.importorskip("flask")
pytest.importorskip("flask_limiter")


# ── fixtures ──────────────────────────────────────────────────────────────────


@pytest.fixture(scope="module")
def client():
    """Flask test client with rate-limiting disabled."""
    from umfrage.server.app import create_app

    app = create_app()
    app.config["TESTING"] = True
    app.config["RATELIMIT_ENABLED"] = False
    with app.test_client() as c:
        yield c


@pytest.fixture()
def valid_yaml(sample_questionnaire):
    """Serialise the shared sample questionnaire fixture to a YAML string."""
    return yaml.dump(
        sample_questionnaire.model_dump(mode="json"), allow_unicode=True
    )


# ── static routes ─────────────────────────────────────────────────────────────


def test_index_returns_html(client):
    r = client.get("/")
    assert r.status_code == 200
    assert b"umfrage" in r.data.lower()


def test_index_has_security_headers(client):
    r = client.get("/")
    assert r.headers.get("X-Content-Type-Options") == "nosniff"
    assert r.headers.get("X-Frame-Options") == "DENY"
    assert "Content-Security-Policy" in r.headers
    # Nonce-based CSP must not contain 'unsafe-inline' for scripts
    csp = r.headers["Content-Security-Policy"]
    assert "nonce-" in csp
    assert "'unsafe-inline'" not in csp.split("script-src")[1].split(";")[0]


def test_schema_json_is_served(client):
    r = client.get("/static/schema.json")
    assert r.status_code == 200
    data = r.get_json()
    assert "$schema" in data
    assert data.get("title") == "umfrage Questionnaire Configuration"


# ── /api/generate — success paths ─────────────────────────────────────────────


def test_generate_valid_yaml_returns_zip(client, valid_yaml):
    r = client.post("/api/generate", data={"yaml_content": valid_yaml})
    assert r.status_code == 200
    assert "application/zip" in r.content_type
    buf = io.BytesIO(r.data)
    with zipfile.ZipFile(buf) as zf:
        names = zf.namelist()
    assert any(n.endswith(".xlsx") for n in names)
    assert any(n.endswith(".yaml") for n in names)
    assert any(n.endswith(".txt") for n in names)


def test_generate_zip_contains_correct_filenames(client, valid_yaml):
    r = client.post("/api/generate", data={"yaml_content": valid_yaml})
    assert r.status_code == 200
    buf = io.BytesIO(r.data)
    with zipfile.ZipFile(buf) as zf:
        names = zf.namelist()
    # All three files should share the same slug prefix
    slugs = {n.split("_questionnaire")[0] for n in names if "_questionnaire" in n}
    slugs |= {n.split("_metadata")[0] for n in names if "_metadata" in n}
    slugs |= {n.split("_generation_log")[0] for n in names if "_generation_log" in n}
    assert len(slugs) == 1, f"Expected one slug, got: {slugs}"


def test_generate_with_style_file_returns_zip(client, valid_yaml):
    style_yaml = "warning_color: 'FFD700'\n"
    r = client.post(
        "/api/generate",
        content_type="multipart/form-data",
        data={
            "yaml_content": valid_yaml,
            "style_file": (io.BytesIO(style_yaml.encode()), "style.yaml"),
        },
    )
    assert r.status_code == 200
    assert "application/zip" in r.content_type


def test_generate_zip_contains_log_file(client, valid_yaml):
    r = client.post("/api/generate", data={"yaml_content": valid_yaml})
    assert r.status_code == 200
    buf = io.BytesIO(r.data)
    with zipfile.ZipFile(buf) as zf:
        log_names = [n for n in zf.namelist() if n.endswith(".txt")]
        assert len(log_names) == 1
        log_text = zf.read(log_names[0]).decode()
    assert "umfrage Generation Log" in log_text
    assert "Test Survey 2024" in log_text
    assert "ZIP contents" in log_text
    assert "_questionnaire_" in log_text
    assert "_metadata_" in log_text


def test_generate_security_headers_on_zip_response(client, valid_yaml):
    r = client.post("/api/generate", data={"yaml_content": valid_yaml})
    assert r.status_code == 200
    assert r.headers.get("X-Content-Type-Options") == "nosniff"


# ── /api/generate — validation errors ─────────────────────────────────────────


def test_generate_invalid_yaml_syntax_returns_400(client):
    r = client.post("/api/generate", data={"yaml_content": "key: [unclosed"})
    assert r.status_code == 400
    assert "error" in r.get_json()


def test_generate_non_mapping_yaml_returns_400(client):
    r = client.post("/api/generate", data={"yaml_content": "- item1\n- item2\n"})
    assert r.status_code == 400
    assert "mapping" in r.get_json()["error"].lower()


def test_generate_missing_required_fields_returns_400(client):
    # Valid YAML but not a valid Questionnaire (missing organizer, sections, etc.)
    r = client.post(
        "/api/generate", data={"yaml_content": "title: Only a title\n"}
    )
    assert r.status_code == 400
    assert "error" in r.get_json()


def test_generate_empty_content_returns_400(client):
    r = client.post("/api/generate", data={"yaml_content": ""})
    assert r.status_code == 400


def test_generate_oversized_yaml_returns_400(client, valid_yaml):
    big_yaml = valid_yaml + "\n# " + "x" * 200_000
    r = client.post("/api/generate", data={"yaml_content": big_yaml})
    assert r.status_code == 400
    assert "100 KB" in r.get_json()["error"]


def test_generate_invalid_style_syntax_returns_400(client, valid_yaml):
    r = client.post(
        "/api/generate",
        content_type="multipart/form-data",
        data={
            "yaml_content": valid_yaml,
            "style_file": (io.BytesIO(b"key: [unclosed"), "style.yaml"),
        },
    )
    assert r.status_code == 400
    assert "Style" in r.get_json()["error"]


def test_generate_invalid_style_structure_returns_400(client, valid_yaml):
    # column_widths expects a mapping; passing a bare string cannot be coerced
    bad_style = "column_widths: this_is_not_a_mapping\n"
    r = client.post(
        "/api/generate",
        content_type="multipart/form-data",
        data={
            "yaml_content": valid_yaml,
            "style_file": (io.BytesIO(bad_style.encode()), "style.yaml"),
        },
    )
    assert r.status_code == 400


def test_generate_oversized_style_file_returns_400(client, valid_yaml):
    big_style = "warning_color: 'FFD700'\n# " + "x" * 60_000
    r = client.post(
        "/api/generate",
        content_type="multipart/form-data",
        data={
            "yaml_content": valid_yaml,
            "style_file": (io.BytesIO(big_style.encode()), "style.yaml"),
        },
    )
    assert r.status_code == 400
    assert "50 KB" in r.get_json()["error"]


# ── index.html template integrity ─────────────────────────────────────────────


def test_index_contains_required_element_ids(client):
    """All IDs referenced by applyTranslations() must be present in the page."""
    r = client.get("/")
    html = r.data.decode()
    for eid in ("llm-box", "collect-box", "validation-note", "editor-container"):
        assert f'id="{eid}"' in html, f"Missing element: id=\"{eid}\""


def test_index_inline_script_has_no_double_commas(client):
    """A double comma (,,) in JS is a syntax error that silently kills the editor."""
    r = client.get("/")
    html = r.data.decode()
    # Isolate the inline <script> block so we don't false-positive on HTML content
    start = html.find("<script ")
    end = html.rfind("</script>")
    assert start != -1 and end != -1
    script = html[start:end]
    assert ",," not in script, "Double comma found in inline script — JS syntax error"


def test_index_both_i18n_locales_define_collect_box(client):
    """Both 'en' and 'de' translation objects must define collectBoxHtml."""
    r = client.get("/")
    script_block = r.data.decode()
    assert script_block.count("collectBoxHtml:") == 2
