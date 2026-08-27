"""Command-line interface for the umfrage questionnaire tool.

Commands
--------
``umfrage validate <config.yaml>``
    Check a questionnaire YAML for syntax and completeness without generating
    any files. Exits with code 1 if errors are found.

``umfrage generate <config.yaml>``
    Validate the config and generate a protected ``.xlsx`` questionnaire file.
    Optionally write a companion ``*_metadata.yaml`` (``--metadata-file``).

``umfrage collect <responses-dir>``
    Scan a folder for returned ``.xlsx`` files, group them by questionnaire
    identity, validate each, and aggregate into one ``results_*.xlsx`` per
    questionnaire found.
"""

from __future__ import annotations

import datetime
import sys
from pathlib import Path
from typing import Callable

import click

from umfrage.checker import check_questionnaire
from umfrage.collector import GroupInfo, collect_all, list_questionnaire_groups
from umfrage.config_loader import ConfigError, load_questionnaire, load_style
from umfrage.generator import generate_questionnaire, write_metadata_file
from umfrage.models import StyleConfig

# Auto-detection order when --style is not given (both relative to CWD).
_STYLE_CWD = Path("style.yaml")
_STYLE_CONFIG = Path("config/style.yaml")


def _resolve_style(style_path: Path | None) -> tuple[StyleConfig, str]:
    """Return (StyleConfig, info_message) for the appropriate style source.

    Priority:
      1. Explicit ``--style`` path.
      2. ``style.yaml`` in the current working directory.
      3. ``config/style.yaml`` relative to the current working directory.
      4. Built-in defaults.
    """
    if style_path is not None:
        return load_style(style_path), f"[INFO] Using style: {style_path}"
    if _STYLE_CWD.exists():
        return load_style(_STYLE_CWD), f"[INFO] Using style: {_STYLE_CWD}"
    if _STYLE_CONFIG.exists():
        return load_style(_STYLE_CONFIG), f"[INFO] Using style: {_STYLE_CONFIG}"
    return StyleConfig(), "[INFO] No style file found; using built-in defaults."


@click.group()
@click.version_option(package_name="umfrage")
def main() -> None:
    """umfrage — Excel-based questionnaire generation and collection tool.

    Typical workflow:

    \b
      1. Author a questionnaire YAML (use config/questionnaire_sample.yaml
         as a template or docs/llm_guide.md as an LLM prompt).
      2. Validate:   umfrage validate config/questionnaire.yaml
      3. Generate:   umfrage generate config/questionnaire.yaml
      4. Distribute the .xlsx to respondents.
      5. Collect:    umfrage collect responses/
    """


# ── validate ──────────────────────────────────────────────────────────────────

@main.command("validate")
@click.argument(
    "config",
    metavar="CONFIG",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
)
@click.option(
    "--style",
    "style_path",
    metavar="STYLE",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=None,
    help=(
        "Path to style.yaml. If omitted, style.yaml in the current directory is "
        "tried first, then config/style.yaml, then built-in defaults."
    ),
)
def cmd_validate(config: Path, style_path: Path | None) -> None:
    """Validate a questionnaire YAML config for syntax and completeness.

    CONFIG is the path to the questionnaire YAML file to check.

    Runs Pydantic structural validation followed by completeness checks
    (unique IDs, slug-safe IDs, scale min/max, email format, etc.).
    Exits with code 1 if any errors are found; warnings are printed but do
    not cause a non-zero exit.
    """
    try:
        questionnaire = load_questionnaire(config)
    except ConfigError as exc:
        click.echo(f"[ERROR] {exc}", err=True)
        sys.exit(1)

    try:
        _, style_msg = _resolve_style(style_path)
    except ConfigError as exc:
        click.echo(f"[ERROR] Style config: {exc}", err=True)
        sys.exit(1)
    click.echo(style_msg)

    check_result = check_questionnaire(questionnaire)

    for warning in check_result.warnings:
        click.echo(f"[WARNING] {warning}")

    if not check_result.is_valid:
        click.echo(
            f"\n[FAIL] {len(check_result.errors)} error(s) found in '{config}':",
            err=True,
        )
        for error in check_result.errors:
            click.echo(f"  • {error}", err=True)
        sys.exit(1)

    sections = len(questionnaire.sections)
    questions = len(questionnaire.all_questions())
    click.echo(
        f"[OK] '{config}' is valid — "
        f"{sections} section(s), {questions} question(s)."
    )


# ── generate ──────────────────────────────────────────────────────────────────

@main.command("generate")
@click.argument(
    "config",
    metavar="CONFIG",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
)
@click.option(
    "--output-dir", "-o",
    metavar="DIR",
    type=click.Path(file_okay=False, path_type=Path),
    default=None,
    help="Directory where the .xlsx file is written. Defaults to the current directory.",
)
@click.option(
    "--style",
    "style_path",
    metavar="STYLE",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=None,
    help=(
        "Path to style.yaml. If omitted, style.yaml in the current directory is "
        "tried first, then config/style.yaml, then built-in defaults."
    ),
)
@click.option(
    "--no-metadata-file",
    "metadata_file",
    is_flag=True,
    flag_value=False,
    default=True,
    help=(
        "Skip writing the companion *_metadata.yaml file. "
        "By default the metadata file is always written alongside the "
        "questionnaire so that 'umfrage collect' can be run without the "
        "original questionnaire.yaml being present."
    ),
)
def cmd_generate(
    config: Path,
    output_dir: Path | None,
    style_path: Path | None,
    metadata_file: bool,
) -> None:
    """Generate a protected Excel questionnaire file from a YAML config.

    CONFIG is the path to the questionnaire YAML file.

    The config is validated before generation. Use 'umfrage validate' to check
    the config without generating any files.

    The output filename is derived from the questionnaire title:
    ``{slug}_questionnaire.xlsx``.
    """
    try:
        questionnaire = load_questionnaire(config)
    except ConfigError as exc:
        click.echo(f"[ERROR] {exc}", err=True)
        sys.exit(1)

    check_result = check_questionnaire(questionnaire)
    if not check_result.is_valid:
        click.echo(
            "[ERROR] Config validation failed. Fix the following errors before generating:",
            err=True,
        )
        for error in check_result.errors:
            click.echo(f"  • {error}", err=True)
        sys.exit(1)

    for warning in check_result.warnings:
        click.echo(f"[WARNING] {warning}")

    style, style_msg = _resolve_style(style_path)
    click.echo(style_msg)
    out_dir = output_dir or Path(".")
    qid = questionnaire.questionnaire_id()
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx_path = out_dir / f"{qid}_questionnaire_{timestamp}.xlsx"

    try:
        generate_questionnaire(questionnaire, style, xlsx_path, config_file=config.name)
    except Exception as exc:
        click.echo(f"[ERROR] Failed to generate Excel file: {exc}", err=True)
        sys.exit(1)

    click.echo(f"[OK] Questionnaire generated: {xlsx_path}")

    if metadata_file:
        meta_path = out_dir / f"{qid}_metadata_{timestamp}.yaml"
        try:
            write_metadata_file(questionnaire, meta_path)
            click.echo(f"[OK] Metadata file written:  {meta_path}")
        except Exception as exc:
            click.echo(f"[WARNING] Could not write metadata file: {exc}")


# ── list ─────────────────────────────────────────────────────────────────────

@main.command("list")
@click.argument(
    "responses_dir",
    metavar="RESPONSES_DIR",
    type=click.Path(exists=True, file_okay=False, path_type=Path),
)
@click.option(
    "--config", "config_path",
    metavar="CONFIG",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=None,
    help="Path to questionnaire YAML for resolving group titles and IDs.",
)
@click.option(
    "--files", "-f",
    "show_files",
    is_flag=True,
    default=False,
    help="List individual response filenames under each group.",
)
def cmd_list(responses_dir: Path, config_path: Path | None, show_files: bool) -> None:
    """List questionnaire groups found in RESPONSES_DIR without collecting.

    RESPONSES_DIR is the folder containing returned .xlsx files.

    Prints one row per questionnaire group with its ID (slug), title, and
    config hash prefix.  Use the ID or hash prefix with
    'umfrage collect --survey TOKEN' to process only selected groups.
    """
    config_override = None
    if config_path:
        try:
            config_override = load_questionnaire(config_path)
        except ConfigError as exc:
            click.echo(f"[ERROR] {exc}", err=True)
            sys.exit(1)

    groups: list[GroupInfo] = list_questionnaire_groups(responses_dir, config_override)

    if not groups:
        click.echo("[INFO] No questionnaire response files found.")
        return

    # Dynamic column widths — slugs never truncated; titles/config capped.
    _TITLE_MAX = 48
    _CFG_MAX = 32
    _HASH_W = 13  # "xxxxxxxxxxxx…"
    cfg_w = max(len("Config file"), min(_CFG_MAX, max(len(g.config_file or "-") for g in groups)))
    id_w = max(len("ID (slug)"), max(len(g.questionnaire_id) for g in groups))
    title_w = max(len("Title"), min(_TITLE_MAX, max(len(g.title) for g in groups)))

    click.echo(f"Found {len(groups)} questionnaire group(s) in '{responses_dir}':\n")
    click.echo(
        f"  {'#':<4}  {'Config file':<{cfg_w}}  {'ID (slug)':<{id_w}}  "
        f"{'Title':<{title_w}}  {'Hash prefix':<{_HASH_W}}  Files"
    )
    click.echo("  " + "-" * (4 + 2 + cfg_w + 2 + id_w + 2 + title_w + 2 + _HASH_W + 2 + 5))
    for i, g in enumerate(groups, start=1):
        cfg_val = g.config_file or "-"
        cfg_display = cfg_val if len(cfg_val) <= cfg_w else cfg_val[:cfg_w - 1] + "…"
        title_display = (
            g.title if len(g.title) <= title_w else g.title[:title_w - 1] + "…"
        )
        flag = "  [!] no metadata" if g.unresolvable else ""
        click.echo(
            f"  {i:<4}  {cfg_display:<{cfg_w}}  {g.questionnaire_id:<{id_w}}  "
            f"{title_display:<{title_w}}  {g.config_hash[:12]}…  {g.file_count}{flag}"
        )
        if show_files:
            for fp in g.files:
                click.echo(f"         {fp.name}")

    if not show_files:
        click.echo("\n  Tip: run with --files (-f) to list individual response filenames.")


# ── collect helpers ──────────────────────────────────────────────────────────

def _make_invalid_prompt() -> Callable[[Path, list[str]], str]:
    """Return a stateful interactive callback for handling invalid files.

    The returned callable prompts the user once per file unless the user
    chose *All* or *None*, which short-circuits all subsequent calls.

    Choices:
      ``I`` / Enter  — Include this file (default).
      ``S``          — Skip this file.
      ``A``          — Include this AND all remaining invalid files.
      ``N``          — Skip this AND all remaining invalid files.

    When stdin is not a tty the prompt is skipped and ``"include"`` is
    returned (matching the documented default behaviour).

    Returns:
        A callable ``(path, errors) -> "include" | "skip"``.
    """
    _forced: list[str | None] = [None]  # shared mutable cell for forced decision

    def _prompt(path: Path, errors: list[str]) -> str:
        if _forced[0] is not None:
            return _forced[0]

        click.echo(f"\n[WARN] '{path.name}' failed validation:")
        for err in errors:
            click.echo(f"  • {err}")

        if not sys.stdin.isatty():
            click.echo(
                "  (non-interactive mode: including file with missing answers marked)"
            )
            return "include"

        while True:
            raw = click.prompt(
                "  Include? [I]nclude / [S]kip / include [A]ll / skip [N]one",
                default="I",
            ).strip().upper()
            if raw in ("I", ""):
                return "include"
            if raw == "S":
                return "skip"
            if raw == "A":
                _forced[0] = "include"
                return "include"
            if raw == "N":
                _forced[0] = "skip"
                return "skip"
            click.echo("  Please enter I, S, A, or N.")

    return _prompt


# ── collect ───────────────────────────────────────────────────────────────────

@main.command("collect")
@click.argument(
    "responses_dir",
    metavar="RESPONSES_DIR",
    type=click.Path(exists=True, file_okay=False, path_type=Path),
)
@click.option(
    "--config", "config_path",
    metavar="CONFIG",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=None,
    help=(
        "Path to questionnaire YAML. Optional if *_metadata.yaml files are "
        "present in RESPONSES_DIR. If provided, overrides metadata discovery "
        "for all questionnaire groups."
    ),
)
@click.option(
    "--style",
    "style_path",
    metavar="STYLE",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=None,
    help=(
        "Path to style.yaml. If omitted, style.yaml in the current directory is "
        "tried first, then config/style.yaml, then built-in defaults."
    ),
)
@click.option(
    "--output-dir", "-o",
    metavar="DIR",
    type=click.Path(file_okay=False, path_type=Path),
    default=None,
    help=(
        "Directory where result files are written. "
        "Defaults to RESPONSES_DIR."
    ),
)
@click.option(
    "--skip-invalid",
    is_flag=True,
    default=False,
    help=(
        "Silently skip files that fail validation instead of prompting. "
        "Useful for non-interactive/CI use."
    ),
)
@click.option(
    "--survey",
    "survey_filter",
    metavar="TOKEN",
    multiple=True,
    help=(
        "Process only the group matching TOKEN. TOKEN may be the questionnaire "
        "ID (slug) or a config hash prefix (\u22658 hex chars, shown by 'umfrage list'). "
        "Repeat to include multiple groups. "
        "Ignored when --config is given."
    ),
)
def cmd_collect(
    responses_dir: Path,
    config_path: Path | None,
    style_path: Path | None,
    output_dir: Path | None,
    skip_invalid: bool,
    survey_filter: tuple[str, ...],
) -> None:
    """Collect and aggregate returned questionnaire files into result spreadsheets.

    RESPONSES_DIR is the folder containing the returned .xlsx files.

    Response files are automatically grouped by questionnaire identity
    (config hash stored in each file's hidden _meta sheet). One
    ``results_*.xlsx`` file is produced per questionnaire group found.

    For each file that fails validation the tool interactively asks whether
    to include it anyway (default) or skip it.  Missing answers in
    force-included files are filled with the configured
    ``missing_answer_marker`` (default: ``XXXXX``) and highlighted in the
    warning colour.  Pass ``--skip-invalid`` to suppress prompting and always
    skip such files.

    Config discovery order:

    \b
      1. --config flag (overrides everything)
      2. *_metadata.yaml files in RESPONSES_DIR (auto-discovered by hash)
    """
    config_override = None
    if config_path:
        try:
            config_override = load_questionnaire(config_path)
        except ConfigError as exc:
            click.echo(f"[ERROR] {exc}", err=True)
            sys.exit(1)

    style, style_msg = _resolve_style(style_path)
    click.echo(style_msg)
    out_dir = output_dir or responses_dir
    on_invalid = None if skip_invalid else _make_invalid_prompt()

    # --survey is meaningless when --config maps every group to one fixed config.
    effective_survey_filter: list[str] | None = None
    if survey_filter:
        if config_override is not None:
            click.echo(
                "[INFO] --survey is ignored when --config is given "
                "(all groups use the supplied config)."
            )
        else:
            effective_survey_filter = list(survey_filter)

    try:
        summaries = collect_all(
            responses_dir, style, out_dir, config_override,
            on_invalid=on_invalid,
            survey_filter=effective_survey_filter,
        )
    except ValueError as exc:
        click.echo(f"[ERROR] {exc}", err=True)
        sys.exit(1)
    except Exception as exc:
        click.echo(f"[ERROR] Collection failed unexpectedly: {exc}", err=True)
        sys.exit(1)

    if not summaries:
        click.echo(
            "[WARNING] No processable questionnaire groups found.\n"
            "  • Ensure response files are .xlsx format.\n"
            "  • Ensure *_metadata.yaml files are present, or pass --config."
        )
        return

    for s in summaries:
        has_issues = s.skipped_count > 0 or s.force_included_count > 0
        status = "WARN" if has_issues else "OK"
        included = s.valid_count + s.force_included_count
        parts = [
            f"[{status}] '{s.questionnaire_title}': "
            f"{included}/{s.total_files} included"
        ]
        if s.force_included_count:
            parts.append(f" ({s.force_included_count} with issues)")
        if s.skipped_count:
            parts.append(f", {s.skipped_count} skipped")
        if s.output_path:
            parts.append(f" → {s.output_path}")
        click.echo("".join(parts))

        for force_path, errors in s.force_included_files:
            click.echo(f"       Included with issues: {force_path.name}")
            for error in errors:
                click.echo(f"         • {error}")

        for skipped_path, errors in s.skipped_files:
            click.echo(f"       Skipped: {skipped_path.name}")
            for error in errors:
                click.echo(f"         • {error}")


# ── serve ─────────────────────────────────────────────────────────────────────

@main.command("serve")
@click.option(
    "--host",
    default="127.0.0.1",
    show_default=True,
    help="Host address to bind.",
)
@click.option(
    "--port",
    default=5000,
    show_default=True,
    help="TCP port to listen on.",
)
def cmd_serve(host: str, port: int) -> None:
    """Start the web interface for browser-based questionnaire generation.

    Opens a local web server where questionnaire YAML can be written or
    loaded, optionally paired with a style file, and compiled to a
    downloadable ZIP (questionnaire .xlsx + metadata .yaml).

    Requires the optional web dependencies:

    \b
      pip install 'umfrage[web]'

    Collection (umfrage collect) is intentionally not exposed via the web
    interface — use the CLI for that step.
    """
    try:
        from waitress import serve as waitress_serve  # type: ignore[import-untyped]

        from umfrage.server.app import create_app
    except ImportError:
        click.echo(
            "[ERROR] Web server dependencies not installed.\n"
            "        Run: pip install 'umfrage[web]'",
            err=True,
        )
        sys.exit(1)

    app = create_app()
    url = f"http://{host}:{port}"
    click.echo(f"[INFO] umfrage web server listening at {url}")
    click.echo(f"[INFO] Open {url} in your browser.")
    click.echo("[INFO] Press Ctrl+C to stop.")
    waitress_serve(app, host=host, port=port, threads=4)
