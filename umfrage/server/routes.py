"""Route handlers for the umfrage web server."""

from __future__ import annotations

import datetime
import io
import secrets
import tempfile
import zipfile
from pathlib import Path

import yaml
from flask import Flask, g, jsonify, render_template, request, send_file
from pydantic import ValidationError

from umfrage.checker import check_questionnaire
from umfrage.generator import generate_questionnaire, write_metadata_file
from umfrage.models import Questionnaire, StyleConfig

_MAX_YAML = 100_000   # bytes
_MAX_STYLE = 50_000   # bytes


def _fmt_pydantic(exc: ValidationError) -> str:
    return "; ".join(
        f"{'.'.join(str(loc) for loc in e['loc'])}: {e['msg']}"
        for e in exc.errors()
    )


def _generation_log(
    questionnaire: Questionnaire,
    timestamp: str,
    xlsx_name: str,
    meta_name: str,
    log_name: str,
    *,
    style_used: bool,
) -> str:
    """Return a plain-text summary written into the ZIP as a human-readable receipt."""
    sections = len(questionnaire.sections)
    questions = len(questionnaire.all_questions())
    org = questionnaire.organizer
    style_note = "provided" if style_used else "not provided (built-in defaults used)"
    iso_ts = datetime.datetime.strptime(timestamp, "%Y%m%d_%H%M%S").isoformat()
    lines = [
        "umfrage Generation Log",
        "======================",
        "",
        f"Generated : {iso_ts}",
        f"Title     : {questionnaire.title}",
        f"ID        : {questionnaire.questionnaire_id()}",
        f"Version   : {questionnaire.version}",
        f"Language  : {questionnaire.language}",
        f"Organizer : {org.name}, {org.institution}",
        f"Sections  : {sections}",
        f"Questions : {questions}",
        f"Style     : {style_note}",
        "",
        "ZIP contents",
        "------------",
        f"{xlsx_name}  — questionnaire form for respondents",
        f"{meta_name}  — required by 'umfrage collect'",
        f"{log_name}  — this file",
    ]
    return "\n".join(lines) + "\n"


def register_routes(app: Flask, limiter) -> None:

    @app.route("/")
    def index():
        g.csp_nonce = secrets.token_urlsafe(16)
        return render_template("index.html", nonce=g.csp_nonce)

    @app.route("/api/generate", methods=["POST"])
    @limiter.limit("20 per minute")
    def api_generate():
        # ── size guards (executed before any parsing) ──────────────────────
        yaml_content = request.form.get("yaml_content", "")
        if len(yaml_content.encode()) > _MAX_YAML:
            return jsonify({"error": "YAML content exceeds 100 KB limit."}), 400

        style = StyleConfig()
        style_file = request.files.get("style_file")
        if style_file and style_file.filename:
            style_bytes = style_file.read(_MAX_STYLE + 1)
            if len(style_bytes) > _MAX_STYLE:
                return jsonify({"error": "Style file exceeds 50 KB limit."}), 400
            try:
                raw_style = yaml.safe_load(style_bytes.decode("utf-8", errors="replace"))
            except yaml.YAMLError as exc:
                return jsonify({"error": f"Style YAML syntax error: {exc}"}), 400
            if not isinstance(raw_style, dict):
                return jsonify({"error": "Style YAML must be a mapping."}), 400
            try:
                style = StyleConfig.model_validate(raw_style)
            except ValidationError as exc:
                return jsonify({"error": f"Invalid style config: {_fmt_pydantic(exc)}"}), 400

        # ── parse and validate questionnaire ───────────────────────────────
        try:
            raw = yaml.safe_load(yaml_content)
        except yaml.YAMLError as exc:
            return jsonify({"error": f"YAML syntax error: {exc}"}), 400

        if not isinstance(raw, dict):
            return jsonify({"error": "Questionnaire YAML must be a mapping at the top level."}), 400

        try:
            questionnaire = Questionnaire.model_validate(raw)
        except ValidationError as exc:
            return jsonify({"error": f"Invalid questionnaire: {_fmt_pydantic(exc)}"}), 400

        check_result = check_questionnaire(questionnaire)
        if not check_result.is_valid:
            bullets = "\n• ".join(check_result.errors)
            return jsonify({"error": f"Questionnaire check failed:\n• {bullets}"}), 400

        # ── generate in isolated temp directory ────────────────────────────
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        qid = questionnaire.questionnaire_id()

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp = Path(tmpdir)
                xlsx_path = tmp / f"{qid}_questionnaire_{timestamp}.xlsx"
                meta_path = tmp / f"{qid}_metadata_{timestamp}.yaml"

                generate_questionnaire(questionnaire, style, xlsx_path)
                write_metadata_file(questionnaire, meta_path)

                log_path = tmp / f"{qid}_generation_log_{timestamp}.txt"
                log_path.write_text(_generation_log(
                    questionnaire, timestamp, xlsx_path.name,
                    meta_path.name, log_path.name,
                    style_used=(style_file is not None and style_file.filename != ""),
                ), encoding="utf-8")

                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    zf.write(xlsx_path, xlsx_path.name)
                    zf.write(meta_path, meta_path.name)
                    zf.write(log_path, log_path.name)
                zip_buf.seek(0)

        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": f"Generation failed: {exc}"}), 500

        return send_file(
            zip_buf,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"{qid}_questionnaire_{timestamp}.zip",
        )
