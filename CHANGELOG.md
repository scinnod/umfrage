# Changelog

All notable changes to **umfrage** will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [0.4.0] – 2026-08-27

### Added
- **`umfrage serve`** — new optional web interface for browser-based questionnaire
  generation.  Install with `pip install 'umfrage[web]'` (adds Flask, Flask-Limiter,
  Waitress); then run `umfrage serve` and open `http://127.0.0.1:5000`.
  Features: CodeMirror 5 YAML editor with live syntax checking, optional style-file
  upload (persisted across Generate calls), one-click ZIP download, bilingual UI
  (en/de, auto-detected from browser language), link to LLM authoring guide.
  Survey collection is intentionally not exposed — use `umfrage collect` on the CLI.
- **Generation log in ZIP** — every `umfrage generate` call via the web interface
  now includes a `*_generation_log_*.txt` in the download ZIP with a human-readable
  summary (title, organizer, section/question counts, style status, file list).
- **Security** — nonce-based `Content-Security-Policy`, per-IP rate limiting
  (20 req/min on `/api/generate`), 100 KB YAML / 50 KB style size guards,
  `tempfile.TemporaryDirectory` isolation per request.
- **16 new pytest tests** covering all web server routes, error paths, ZIP contents,
  security headers, and the generation log.  Tests skip automatically when `[web]`
  is not installed.
- **`docs/web_server.md`** — full documentation for the web interface including
  shared-server deployment guidance.

---

## [0.3.1] – 2026-08-25

### Added
- **`examples/` directory** — self-contained demo questionnaire
  ("Survey Tools in Research — A Meta-Survey") showcasing all four answer
  types (`scale`, `yes_no`, `choices`, `freetext`), named `choice_lists` with
  `choices_ref`, and the full generate → distribute → collect workflow with
  three pre-filled response files.

---

## [0.3.0] – 2026-08-17

### Added
- **`umfrage list RESPONSES_DIR`** — new read-only subcommand that scans a
  folder and prints a table of all questionnaire groups found (ID slug, title,
  config hash prefix, file count) without writing any output files.  Use it
  to discover what is in a folder before running `umfrage collect`.
- **`umfrage collect --survey TOKEN`** — optional filter flag (repeatable) that
  restricts collection to groups whose questionnaire ID (slug) or config hash
  prefix matches TOKEN.  Passing a slug that maps to more than one group
  (title collision) is a hard error with a hint to use the hash prefix shown
  by `umfrage list`.  Ignored when `--config` is also given.
- `GroupInfo` dataclass and `list_questionnaire_groups()` function added to
  the public API of `umfrage.collector`.

### Fixed
- **`umfrage collect` could not auto-discover metadata for most generated
  questionnaires**: `resolve_config()` scanned for `*_metadata.yaml` but
  `umfrage generate` writes `*_metadata_{timestamp}.yaml`.  The glob is now
  `*_metadata*.yaml`, which matches both the timestamped CLI output and the
  old timestamp-free format used in tests and manual workflows.

---

## [0.2.1] – 2026-08-17

### Added
- **Footer row on the questionnaire sheet**: each generated `.xlsx` now ends
  with a small, locked footer row reading "Generated with umfrage" (or
  "Erzeugt mit umfrage" in German) hyperlinked to the project repository.
  Opt out by setting `show_footer: false` in `style.yaml`.
- **`project_url` in `_meta` sheet**: the hidden metadata sheet now stores
  `project_url = https://github.com/scinnod/umfrage` for human reference.
- **`COLLABORATORS.md`**: documents project contributor David Kleinhans
  (Jade University of Applied Sciences, Germany).
- LICENSE updated with the GitHub project URL and a reference to `COLLABORATORS.md`.
- New i18n keys `footer_generated_with` in `en.yaml` / `de.yaml`.

### Fixed
- **`_meta` sheet was unprotected**: the hidden sheet storing structural
  metadata (used by `umfrage collect`) was not write-protected, so it could
  be edited accidentally after unhiding, breaking result collection.
  It is now protected with the same password as the questionnaire sheet.

---

## [0.2.0] – 2026-08-14

### Added
- **`choices` answer type**: user-defined dropdown with an arbitrary list of
  options (2–N items); Excel data-validation enforced identically to `yes_no`.
- **`choice_lists` (top-level YAML key)**: define named option lists once and
  reference them with `answer.choices_ref: <name>` in any number of questions.
  One-off lists can still be defined inline with `answer.choices: [...]`.
  Questions sharing the same resolved list receive a single consolidated
  `DataValidation` object (same safe deferred-registration pattern as `yes_no`).
- **`answer.show_choices_in_comment`** (default `true`): set to `false` to
  suppress listing options in the Scale/Comment column when the list is long.
- **CLI style auto-detection with feedback**: all three commands now print an
  `[INFO]` line reporting which style source is used.
  Priority order: explicit `--style PATH` → `style.yaml` in CWD → `config/style.yaml`
  → built-in defaults.
- New i18n keys in `en.yaml` / `de.yaml`: `hint_choices`, `dv_choices_error`.
- Checker checks 11–16 covering the new `choices` constraints (missing source,
  both sources set, unknown ref, <2 options, >255-char formula, duplicate options).
- JSON Schema updated: `choices` type enum value, `choices`, `choices_ref`,
  `show_choices_in_comment` properties in `answer`, `choice_lists` at root level.
- Sample config (`config/questionnaire_sample.yaml`) extended with two
  `choice_lists` entries and three `choices`-type questions demonstrating
  `choices_ref`, inline `choices`, and `show_choices_in_comment: false`.
- LLM guide (`docs/llm_guide.md`) updated with `§3.8 choice_lists`, `§4.4 choices`,
  extended validation rules, worked example, and two new authoring tips.

---

## [0.1.0] – 2026-08-11

### Added
- `umfrage validate` command: syntax and completeness check for questionnaire YAML configs
- `umfrage generate` command: creates a protected Excel questionnaire from a YAML config
  - Worksheet protection with optional password
  - Data validation dropdowns for scale (1–N) and yes/no answer types
  - Hidden `_meta` sheet storing structural metadata for later validation
  - Optional `--metadata-file` flag: writes a `*_metadata.yaml` companion file
    embedding the full questionnaire model (enables config-free collection)
- `umfrage collect` command: aggregates returned Excel files into a single results spreadsheet
  - Automatic grouping of files by questionnaire identity (config hash)
  - Supports multiple different questionnaires in one folder
  - Config auto-discovery from `*_metadata.yaml` files; `--config` always optional
  - Warning color highlighting for missing or invalid answers in the result file
- YAML-based questionnaire config (`config/questionnaire_sample.yaml`)
- YAML-based style/appearance config (`config/style.yaml`) with optional protection password
- JSON Schema for questionnaire YAML (`docs/questionnaire.schema.json`)
- LLM/AI authoring guide (`docs/llm_guide.md`)
- Full unit test suite (`tests/`) with pytest
- Apache 2.0 license
