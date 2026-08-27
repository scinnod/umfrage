"""Flask application factory for the umfrage web server."""

from __future__ import annotations

import secrets
from pathlib import Path

from flask import Flask, g

_STATIC = Path(__file__).parent / "static"
_TEMPLATES = Path(__file__).parent / "templates"

_BASE_HEADERS = {
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "DENY",
    "Referrer-Policy": "strict-origin-when-cross-origin",
}


def _make_csp(nonce: str | None = None) -> str:
    script_src = "'self' https://cdn.jsdelivr.net"
    if nonce:
        script_src = f"'self' 'nonce-{nonce}' https://cdn.jsdelivr.net"
    return (
        f"default-src 'self'; "
        f"script-src {script_src}; "
        f"style-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net; "
        f"connect-src 'self'"
    )


def create_app() -> Flask:
    """Create and return the configured Flask application."""
    from flask_limiter import Limiter
    from flask_limiter.util import get_remote_address

    app = Flask(
        __name__,
        static_folder=str(_STATIC),
        static_url_path="/static",
        template_folder=str(_TEMPLATES),
    )
    # Hard cap on total upload size (YAML + style file)
    app.config.setdefault("MAX_CONTENT_LENGTH", 200 * 1024)

    limiter = Limiter(
        app=app,
        key_func=get_remote_address,
        default_limits=[],
        storage_uri="memory://",
    )

    from umfrage.server.routes import register_routes
    register_routes(app, limiter)

    @app.after_request
    def _add_security_headers(response):
        for header, value in _BASE_HEADERS.items():
            response.headers[header] = value
        nonce = getattr(g, "csp_nonce", None)
        response.headers["Content-Security-Policy"] = _make_csp(nonce)
        return response

    @app.errorhandler(413)
    def _payload_too_large(_e):
        from flask import jsonify
        return jsonify({"error": "Request payload too large (max 200 KB total)."}), 413

    @app.errorhandler(429)
    def _rate_limit(_e):
        from flask import jsonify
        return jsonify({"error": "Too many requests — please wait a moment."}), 429

    return app
