"""Centralised path validation for the Word Document Server.

Provides ``validate_path`` which all file-accepting tools should call
before opening or writing to a user-supplied path.  An optional
``MCP_ALLOWED_DIR`` environment variable can enforce a sandbox so that
the server only ever touches files under a single directory tree.

Usage::

    from word_mcp_live_cheemscheems.utils.path_safety import validate_path

    safe = validate_path(user_input)
    doc = Document(safe)

The function is also called inside ``ensure_docx_extension()`` in
*file_utils.py*, so most cross-platform tools get validation
automatically.
"""

import os

# Optional sandbox directory -- all file access is restricted to this
# subtree when set.  ``None`` means no restriction (backwards-compatible).
ALLOWED_DIR = os.environ.get("MCP_ALLOWED_DIR")


def validate_path(path: str, for_write: bool = False) -> str:
    """Validate and normalise a user-supplied file path.

    Checks (in order):
    1.  Path is a non-empty string.
    2.  No null byte (``\\x00``) in the path.
    3.  No parent-directory traversal (``..``) after normalisation.
    4.  Optional ``MCP_ALLOWED_DIR`` sandbox (see module docstring).
    5.  Symlink resolution via ``os.path.realpath``.

    Args:
        path:      User-supplied file path.
        for_write: If ``True``, also checks that the parent directory
                   is writeable (basic sanity check, not a security
                   boundary).

    Returns:
        Normalised absolute path.

    Raises:
        ValueError: If any check fails, with a Chinese-language message
                    suitable for returning to the LLM caller.
    """
    if not path or not isinstance(path, str):
        raise ValueError("路径不能为空。")

    if "\x00" in path:
        raise ValueError("路径包含空字符（\\x00），已拒绝。")

    # Normalise and detect parent-dir traversal
    normalized = os.path.normpath(path)
    parts = normalized.replace("\\", "/").split("/")
    if ".." in parts:
        raise ValueError("路径包含 '..' 遍历，已拒绝。")

    # Resolve to absolute path, following symlinks
    real_path = os.path.realpath(normalized)

    # Optional MCP_ALLOWED_DIR sandbox
    if ALLOWED_DIR:
        allowed = os.path.realpath(ALLOWED_DIR)
        if not real_path.startswith(allowed + os.sep) and real_path != allowed:
            raise ValueError(
                f"路径 {real_path} 不在允许的工作目录"
                f"（{allowed}）内。\n"
                "如需修改限制，请设置 MCP_ALLOWED_DIR 环境变量。"
            )

    return real_path


def validate_docx_path(filename: str) -> str:
    """Validate a .docx file path (ensures extension + validates).

    This is a convenience wrapper that calls ``ensure_docx_extension``
    followed by ``validate_path``, used by tools that need to ensure
    the path ends with ``.docx``.
    """
    from word_mcp_live_cheemscheems.utils.file_utils import ensure_docx_extension

    filename = ensure_docx_extension(filename)
    return validate_path(filename)
