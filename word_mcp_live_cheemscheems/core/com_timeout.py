"""COM call timeout wrapper for pywin32.

All COM calls (``win32com.client``) run synchronously in the calling
thread and cannot be interrupted from Python.  This module provides a
utility that runs a COM operation in a **background thread** (where
``CoInitialize`` is called first) with an ``asyncio`` timeout.

Usage::

    from word_mcp_live_cheemscheems.core.com_timeout import com_call, ComTimeoutError

    # Wrap a single COM method call (use a lambda)
    result = await com_call(lambda: doc.Save(), timeout=30)

    # Wrap a multi-step worker that re-connects to Word internally
    result = await com_call(_do_replace_all, filename, find_text,
                            replace_text, timeout=60)

.. note::
    The background thread **cannot be cancelled** mid-way.  If the
    timeout expires, this module raises ``ComTimeoutError`` and the
    caller gets an error response, but the thread continues running
    in the background until its COM call completes.  Its result is
    discarded.
"""

import asyncio
from typing import Any, Callable

import pythoncom


class ComTimeoutError(Exception):
    """Raised when a COM call exceeds its configured time budget."""


async def com_call(func: Callable, *args: Any,
                   timeout: int = 30, **kwargs: Any) -> Any:
    """Execute *func* in a background thread with COM initialized.

    The target thread calls ``pythoncom.CoInitialize()`` before running
    *func*, so it can obtain its own ``Word.Application`` reference via
    ``win32com.client.GetActiveObject()``.

    Args:
        func:    Callable that performs COM operations.  If you need
                 to call a bound COM method (e.g. ``doc.Save()``),
                 wrap it in a ``lambda``.
        timeout: Max seconds to wait.  When exceeded, the async call
                 raises ``ComTimeoutError``.  *func* is **not**
                 interrupted — it continues in the orphan thread.

    Returns:
        Whatever *func* returns.

    Raises:
        ComTimeoutError: If *timeout* seconds elapse.
    """
    def _worker() -> Any:
        pythoncom.CoInitialize()
        try:
            return func(*args, **kwargs)
        finally:
            pythoncom.CoUninitialize()

    loop = asyncio.get_event_loop()
    try:
        return await asyncio.wait_for(
            loop.run_in_executor(None, _worker),
            timeout=timeout,
        )
    except asyncio.TimeoutError:
        raise ComTimeoutError(
            f"COM 操作超过 {timeout} 秒仍未返回，已超时放弃。"
            " 操作仍在后台运行，完成后的结果将被忽略。"
        )
