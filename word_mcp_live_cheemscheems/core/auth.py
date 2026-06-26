"""Bearer Token authentication for HTTP/SSE transport modes.

Usage::

    # Require authentication (default for HTTP/SSE):
    WORD_MCP_LIVE_API_KEY=sk-your-secret-key python -m word_mcp_live_cheemscheems

    # Explicitly allow no-auth (NOT RECOMMENDED for remote access):
    WORD_MCP_LIVE_INSECURE=true python -m word_mcp_live_cheemscheems

The key can also be placed in a ``.env`` file in the project root
(``load_dotenv()`` is called at startup in *main.py*).
"""

import os

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import JSONResponse

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
    """Return ``True`` if HTTP/SSE transport must have authentication.

    When ``True``, the server should refuse to start in HTTP/SSE mode
    without a valid ``WORD_MCP_LIVE_API_KEY`` (unless insecure mode is
    explicitly enabled).
    """
    return not _API_KEY_SET and not _INSECURE


class BearerTokenMiddleware(BaseHTTPMiddleware):
    """Starlette ASGI middleware validating ``Authorization: Bearer <key>``.

    Checks:
    - ``WORD_MCP_LIVE_API_KEY`` set → requires matching Bearer token
    - ``WORD_MCP_LIVE_INSECURE=true`` → passes all requests through
    - Neither → 401 (backstop; server startup should already reject this)
    """

    async def dispatch(self, request: Request, call_next):
        if not _API_KEY_SET:
            return await call_next(request)

        auth_header = request.headers.get("Authorization", "")
        expected = f"Bearer {WORD_MCP_LIVE_API_KEY}"

        if auth_header != expected:
            return JSONResponse(
                {"error": f"未授权。请提供有效的 WORD_MCP_LIVE_API_KEY。"},
                status_code=401,
            )

        return await call_next(request)
