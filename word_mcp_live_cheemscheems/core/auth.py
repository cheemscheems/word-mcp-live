"""Bearer Token authentication for HTTP/SSE transport modes.

Usage::

    # Require authentication (default for HTTP/SSE):
    WORD_MCP_LIVE_API_KEY=sk-your-secret-key python -m word_mcp_live_cheemscheems

    # Explicitly allow no-auth (NOT RECOMMENDED for remote access):
    WORD_MCP_LIVE_INSECURE=true python -m word_mcp_live_cheemscheems

The key can also be placed in a ``.env`` file in the project root
(``load_dotenv()`` is called at startup in *main.py*).
"""

import json
import os

from starlette.types import ASGIApp, Receive, Scope, Send

WORD_MCP_LIVE_API_KEY: str | None = os.environ.get("WORD_MCP_LIVE_API_KEY")
_API_KEY_SET = bool(WORD_MCP_LIVE_API_KEY)

# Explicit opt-in for no-auth mode (for local/dev use only)
WORD_MCP_LIVE_INSECURE: str | None = os.environ.get("WORD_MCP_LIVE_INSECURE")
_INSECURE = bool(WORD_MCP_LIVE_INSECURE and WORD_MCP_LIVE_INSECURE.lower() in ("true", "1", "yes"))


def is_auth_enabled() -> bool:
    """Return ``True`` if Bearer token authentication is active."""
    return _API_KEY_SET


def is_insecure_mode() -> bool:
    """Return ``True`` if the user has explicitly opted into no-auth mode."""
    return not _API_KEY_SET and _INSECURE


def auth_required_for_http() -> bool:
    """Return ``True`` if HTTP/SSE transport must have authentication."""
    return not _API_KEY_SET and not _INSECURE


class BearerTokenMiddleware:
    """Pure ASGI middleware validating ``Authorization: Bearer <key>``.

    This is a low-level ASGI middleware (not Starlette's
    ``BaseHTTPMiddleware``) so it works correctly with SSE streaming
    responses and other non-standard ASGI message sequences.
    """

    def __init__(self, app: ASGIApp) -> None:
        self.app = app

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        # Only intercept HTTP requests; let WebSocket/other pass through
        if not _API_KEY_SET or scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        # Extract Authorization header from ASGI scope
        auth_header = ""
        for name, value in scope.get("headers", []):
            if name == b"authorization":
                auth_header = value.decode("utf-8")
                break

        if auth_header != f"Bearer {WORD_MCP_LIVE_API_KEY}":
            body = json.dumps(
                {"error": "未授权。请提供有效的 WORD_MCP_LIVE_API_KEY。"}
            ).encode("utf-8")
            await send({
                "type": "http.response.start",
                "status": 401,
                "headers": [(b"content-type", b"application/json")],
            })
            await send({
                "type": "http.response.body",
                "body": body,
            })
            return

        await self.app(scope, receive, send)
