# umfrage Web Server

The `umfrage serve` command starts a local web interface for questionnaire
generation.  It is a thin browser front-end for the same `generate` pipeline
available on the CLI — it does **not** expose survey collection.

---

## Installation

The web server depends on optional packages that are not installed with the
base `umfrage` package.  Install them with the `[web]` extra:

```bash
pip install 'umfrage[web]'
```

This adds **Flask**, **Flask-Limiter**, and **Waitress** to your environment.

---

## Starting the server

```bash
umfrage serve
# [INFO] umfrage web server listening at http://127.0.0.1:5000
# [INFO] Open http://127.0.0.1:5000 in your browser.
# [INFO] Press Ctrl+C to stop.
```

Options:

| Option | Default | Description |
|---|---|---|
| `--host` | `127.0.0.1` | Bind address.  Use `0.0.0.0` to allow LAN access. |
| `--port` | `5000` | TCP port. |

Example — shared departmental server:

```bash
umfrage serve --host 0.0.0.0 --port 8080
```

---

## Using the web interface

### Editor

The YAML editor (CodeMirror 6) provides:

- **Syntax highlighting** for YAML.
- **Inline error markers** as you type — syntax errors from js-yaml and
  structural warnings from the embedded questionnaire JSON Schema (Ajv).
- **600 ms debounce** so validation runs only after you pause typing.

### File operations

| Control | Action |
|---|---|
| **⬆ Load YAML** | Open a local `.yaml` / `.yml` file and load it into the editor. |
| **⬇ Save YAML** | Save the current editor content as `questionnaire.yaml` to your local machine. |
| **Style (optional) → Choose…** | Upload a `style.yaml` customisation file.  The selection persists across multiple Generate calls until the page is closed or a new file is selected. |

### Generating a questionnaire

1. Write or paste a questionnaire YAML in the editor (or load from file).
2. Optionally select a style file.
3. Click **⚡ Generate & Download ZIP**.
4. If the YAML is valid the browser downloads a ZIP containing:
   - `{slug}_questionnaire_{timestamp}.xlsx` — the protected Excel form.
   - `{slug}_metadata_{timestamp}.yaml` — the companion metadata file needed later by `umfrage collect`.
5. If the YAML is invalid, an error message is shown below the Generate button.

### LLM authoring assistance

A ready-to-paste prompt for AI-assisted YAML authoring is available in the
[LLM authoring guide](https://github.com/scinnod/umfrage/blob/main/docs/llm_guide.md).
The link is also shown in the web interface itself.

---

## Security

### What is safe

- **YAML parsing**: uses `yaml.safe_load` (no arbitrary object construction).
- **Validation**: all inputs pass through the same Pydantic models used by the CLI.
- **File isolation**: generation runs inside a `tempfile.TemporaryDirectory` that
  is deleted immediately after the ZIP is assembled.  No user data is persisted on
  the server.
- **Size limits**: YAML content ≤ 100 KB, style file ≤ 50 KB, total request ≤ 200 KB.
- **Rate limiting**: `/api/generate` is limited to 20 requests per minute per IP.
- **Security headers**: every response carries
  `X-Content-Type-Options: nosniff`, `X-Frame-Options: DENY`,
  `Referrer-Policy: strict-origin-when-cross-origin`, and a per-request
  nonce-based `Content-Security-Policy` that allows scripts only from
  `self` and `cdn.jsdelivr.net`.

### What is intentionally excluded

**Survey collection** (`umfrage collect`) is not exposed via the web server.
Collection processes files uploaded by third parties and involves more complex
validation logic; the risk profile is higher and it is better handled on the
CLI by the survey organiser who controls the environment.

### Running on a shared server

The default bind address `127.0.0.1` limits access to the local machine.
For a departmental server:

1. Use `--host 0.0.0.0` or bind to the specific interface.
2. Place the server behind a TLS-terminating reverse proxy
   (nginx, Caddy, or similar) to encrypt traffic.
3. Optionally add HTTP Basic Authentication at the proxy layer if the
   tool should not be accessible to everyone on the network.

**Waitress** (the WSGI server used by `umfrage serve`) runs a thread pool
(4 threads by default) and handles concurrent requests without the
limitations of Flask's development server.

---

## Architecture notes

The web server lives entirely inside the `umfrage.server` subpackage:

```
umfrage/server/
    __init__.py
    app.py          # Flask app factory, security headers, rate limiter
    routes.py       # /api/generate endpoint
    static/
        schema.json # questionnaire JSON Schema served to the browser editor
    templates/
        index.html  # single-page UI (CodeMirror 6, no build step)
```

`umfrage serve` is a Click subcommand in `umfrage/cli.py` that lazy-imports
`flask`, `flask_limiter`, and `waitress`.  If the `[web]` extras are not
installed the command prints a helpful error and exits.
