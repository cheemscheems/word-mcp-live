"""Live editing tools for Microsoft Word via COM automation.

These tools operate on documents that are currently open in Word,
providing real-time editing capabilities with optional tracked changes.
"""

import json
import os
import re
import sys
import time

from word_document_server.defaults import DEFAULT_AUTHOR
# macOS JXA dispatch
_MAC_AVAILABLE = __import__('sys').platform == 'darwin'


# Word COM constants
WD_STORY = 6

# Word COM InsertBefore/InsertAfter limit (~32K chars).
# We use 30000 as safe margin below 2^15-1 = 32767.
_INSERT_CHUNK_SIZE = 30000

# Safety limits for find-replace operations.
_MAX_REPLACEMENTS = 50_000   # hard ceiling on number of replacements
_MAX_REPLACE_SEC = 30        # time budget (seconds) before truncation


async def word_live_insert_text(
    filename: str = None,
    text: str = "",
    position: str = "end",
    bookmark: str = None,
    track_changes: bool = False,
) -> str:
    """Insert text into an open Word document.

    Automatically chunks large text (>30K chars) to avoid Word COM limits.

    Args:
        filename: Document name or path (None = active document).
        text: Text to insert (no length limit — auto-chunked if needed).
        position: "start", "end", "cursor", or character offset as string.
        bookmark: Insert after a named bookmark (overrides position).
        track_changes: Track the insertion as a revision.

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_insert_text
        return mac_insert_text(filename=filename, text=text, position=position, bookmark=bookmark, track_changes=track_changes)

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        result = await com_call(
            _do_replace_all,
            filename, find_text, replace_text,
            match_case, match_whole_word, use_wildcards,
            replace_all, track_changes,
            timeout=60,
        )
        count = result["count"]
        count_truncated = result["truncated"]

        return json.dumps({
            "success": True,
            "document": filename,
            "find_text": find_text,
            "replace_text": replace_text,
            "replacements": count,
            "replace_all": replace_all,
            "tracked": track_changes,
            "truncated": count_truncated,
        }, ensure_ascii=False)

    except ComTimeoutError as e:
        return json.dumps({"error": str(e), "timeout": True, "truncated": True})
    except Exception as e:
        return json.dumps({"error": str(e)})

async def word_live_insert_paragraphs(
    filename: str = None,
    paragraphs: list = None,
    target_text: str = None,
    target_paragraph_index: int = None,
    position: str = "after",
    style: str = None,
    track_changes: bool = False,
) -> str:
    """[Windows only] Insert one or more paragraphs near a target paragraph in an open Word document.

    Targets by text match or paragraph index (0-based, matching word_live_get_text output).
    Inserts all paragraphs in a single undo record.

    Args:
        filename: Document name or path (None = active document).
        paragraphs: List of paragraph texts to insert. Each string becomes one Word paragraph.
        target_text: Text to search for (first matching paragraph). Mutually exclusive with target_paragraph_index.
        target_paragraph_index: 0-based paragraph index (as returned by word_live_get_text).
        position: 'before' or 'after' the target paragraph (default 'after').
        style: Style name for inserted paragraphs. None = "Normal" (avoids inheriting heading styles).
        track_changes: Track insertions as revisions.

    Returns:
        JSON with result info including count of paragraphs inserted.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_insert_paragraphs is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not paragraphs or not isinstance(paragraphs, list):
        return json.dumps({"error": "paragraphs must be a non-empty list of strings"})

    if target_text is None and target_paragraph_index is None:
        return json.dumps({"error": "Provide either target_text or target_paragraph_index"})

    if target_text is not None and target_paragraph_index is not None:
        return json.dumps({"error": "Provide target_text or target_paragraph_index, not both"})

    if position not in ("before", "after"):
        return json.dumps({"error": f"position must be 'before' or 'after', got '{position}'"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Find the target paragraph
        total_paras = doc.Paragraphs.Count
        target_para = None

        if target_paragraph_index is not None:
            com_index = target_paragraph_index + 1  # 0-based API → 1-based COM
            if com_index < 1 or com_index > total_paras:
                return json.dumps({
                    "error": f"target_paragraph_index {target_paragraph_index} out of range "
                    f"(0-{total_paras - 1})"
                })
            target_para = doc.Paragraphs(com_index)
        else:
            for i in range(1, total_paras + 1):
                para = doc.Paragraphs(i)
                para_text = para.Range.Text.rstrip("\r\x07")
                if target_text in para_text:
                    target_para = para
                    break
            if target_para is None:
                return json.dumps({"error": f"No paragraph found containing '{target_text}'"})

        resolved_style = style if style else "Normal"

        with undo_record(app, "MCP: Insert Paragraphs"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                inserted = 0

                if position == "after":
                    rng = target_para.Range.Duplicate
                    rng.Collapse(0)  # wdCollapseEnd
                    for para_text in paragraphs:
                        rng.InsertParagraphAfter()
                        rng.Collapse(0)  # wdCollapseEnd
                        rng.InsertAfter(para_text)
                        try:
                            rng.Style = resolved_style
                        except Exception:
                            pass
                        rng.Collapse(0)  # wdCollapseEnd
                        inserted += 1
                else:  # "before"
                    for para_text in reversed(paragraphs):
                        rng = target_para.Range.Duplicate
                        rng.Collapse(1)  # wdCollapseStart
                        rng.InsertParagraphBefore()
                        rng.Collapse(1)  # wdCollapseStart
                        rng.InsertAfter(para_text)
                        try:
                            rng.Style = resolved_style
                        except Exception:
                            pass
                        inserted += 1
            finally:
                doc.TrackRevisions = prev_tracking
                if track_changes:
                    app.UserName = prev_author

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_inserted": inserted,
            "position": position,
            "style": resolved_style,
            "tracked": track_changes,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_table(
    filename: str = None,
    rows: int = 2,
    cols: int = 2,
    position: str = "end",
    data: list = None,
    style: str = "Table Grid",
    autofit: str = "window",
    track_changes: bool = False,
) -> str:
    """Add a table to an open Word document.

    Args:
        filename: Document name or path.
        rows: Number of rows.
        cols: Number of columns.
        position: "start", "end", or character offset.
        data: Optional 2D list of cell data.
        style: Table style name. Default "Table Grid" (bordered).
            Use None or "" for no style.
        autofit: "window" (fit page width, default), "content" (fit cell content),
            "fixed" (fixed widths), or None for legacy behavior (no autofit).
        track_changes: Track as revision.

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_table
        return mac_add_table(filename=filename, rows=rows, cols=cols, position=position, data=data, track_changes=track_changes)

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if position == "start":
            rng = doc.Range(0, 0)
        elif position == "end":
            end_pos = doc.Content.End - 1
            rng = doc.Range(end_pos, end_pos)
        else:
            try:
                offset = int(position)
            except ValueError:
                return json.dumps({"error": f"Invalid position: {position}"})

            # Reject offsets that would land the new table inside an
            # existing table's range — Word would silently merge the
            # new structure into the old, breaking both.
            for t in doc.Tables:
                try:
                    ts, te = t.Range.Start, t.Range.End
                except Exception:
                    continue
                if ts <= offset <= te:
                    return json.dumps({
                        "error": (
                            f"position offset {offset} falls within an existing "
                            f"table at range [{ts}, {te}]. Choose an offset "
                            f"outside any table, or use position='end'/'start'."
                        )
                    })

            # Reject offsets immediately after an orphan cell separator
            # (residue from a prior Table.Delete with scrub disabled);
            # adding a table at such a point fuses it with the residue.
            if offset > 0:
                try:
                    probe = doc.Range(offset - 1, offset).Text or ""
                except Exception:
                    probe = ""
                if probe == "\x07":
                    return json.dumps({
                        "error": (
                            f"position offset {offset} sits immediately after "
                            f"an orphan cell separator (\\x07). Run "
                            f"word_live_modify_table operation='delete_table' "
                            f"with scrub_orphans=True (the default) on the "
                            f"prior table, or use word_live_diagnose_layout "
                            f"to locate and clean separators."
                        )
                    })

            rng = doc.Range(offset, offset)

        with undo_record(app, "MCP: Add Table"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                # AutoFit behavior constants
                AUTOFIT_MAP = {
                    "window": (1, 2),   # wdWord9TableBehavior, wdAutoFitWindow
                    "content": (1, 1),  # wdWord9TableBehavior, wdAutoFitContent
                    "fixed": (0, 0),    # wdWord8TableBehavior, wdAutoFitFixed
                }

                if autofit and autofit.lower() in AUTOFIT_MAP:
                    default_behavior, autofit_behavior = AUTOFIT_MAP[autofit.lower()]
                    table = doc.Tables.Add(rng, rows, cols, default_behavior, autofit_behavior)
                else:
                    table = doc.Tables.Add(rng, rows, cols)

                # Apply table style
                if style:
                    try:
                        table.Style = doc.Styles(style)
                    except Exception:
                        pass  # Style not found; proceed without

                if data:
                    for r_idx, row_data in enumerate(data):
                        if r_idx >= rows:
                            break
                        for c_idx, cell_val in enumerate(row_data):
                            if c_idx >= cols:
                                break
                            table.Cell(r_idx + 1, c_idx + 1).Range.Text = str(cell_val)
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "rows": rows,
                "cols": cols,
                "position": position,
                "style": style or None,
                "autofit": autofit or None,
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_format_table(
    filename: str = None,
    table_index: int = -1,
    border_style: str = None,
    cell_bold: list = None,
    cell_alignment: list = None,
    column_widths: list = None,
    table_alignment: str = None,
    cell_shading: list = None,
    autofit: str = None,
) -> str:
    """Format a table in an open Word document via COM.

    Supports border removal, cell formatting, column sizing, and table alignment.
    Use table_index=-1 for the last table, 1 for the first, etc.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index, or -1 for the last table.
        border_style: Border style for all edges: "none", "single", "double", "dotted",
            "dashed", "thick". "none" removes all borders.
        cell_bold: List of [row, col, bold] entries (1-indexed) to set bold on cell text.
            Example: [[1, 1, true], [1, 2, true]] bolds row 1 cells.
        cell_alignment: List of [row, col, alignment] entries. alignment: "left", "center",
            "right", "justify". Row 0 = all rows, Col 0 = all cols.
        column_widths: List of column widths in points (1-indexed order).
            Example: [200, 200] sets col 1 to 200pt, col 2 to 200pt.
        table_alignment: Table alignment on page: "left", "center", "right".
        cell_shading: List of [row, col, color_hex] entries. color_hex as "#RRGGBB".
            Row 0 = all rows. Example: [[1, 0, "#DDDDDD"]] shades entire row 1.
        autofit: "window" (fit to page width), "content" (fit to cell content),
            "fixed" (fixed column widths).

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_format_table is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if doc.Tables.Count == 0:
            return json.dumps({"error": "Document has no tables"})

        idx = table_index if table_index > 0 else doc.Tables.Count
        if idx < 1 or idx > doc.Tables.Count:
            return json.dumps({"error": f"Table index {table_index} out of range (1-{doc.Tables.Count})"})

        tbl = doc.Tables(idx)
        actions = []

        # Border style constants
        BORDER_STYLES = {
            "none": 0,     # wdLineStyleNone
            "single": 1,   # wdLineStyleSingle
            "double": 7,   # wdLineStyleDouble
            "dotted": 3,   # wdLineStyleDot
            "dashed": 2,   # wdLineStyleDash
            "thick": 6,    # wdLineStyleThickThinSmallGap (thick)
        }

        BORDER_IDS = [-1, -2, -3, -4, -5, -6, -7, -8]  # top, left, bottom, right, horiz, vert, etc.

        with undo_record(app, "MCP: Format Table"):
            # --- Borders ---
            if border_style is not None:
                style_val = BORDER_STYLES.get(border_style.lower())
                if style_val is None:
                    return json.dumps({"error": f"Unknown border_style: {border_style}. Use: {list(BORDER_STYLES.keys())}"})
                for bid in BORDER_IDS:
                    try:
                        tbl.Borders(bid).LineStyle = style_val
                    except Exception:
                        pass
                actions.append(f"borders={border_style}")

            # --- Autofit ---
            if autofit is not None:
                AUTOFIT = {"window": 2, "content": 1, "fixed": 0}  # wdAutoFitWindow=2, wdAutoFitContent=1, wdAutoFitFixed=0
                af_val = AUTOFIT.get(autofit.lower())
                if af_val is not None:
                    tbl.AutoFitBehavior(af_val)
                    actions.append(f"autofit={autofit}")

            # --- Table alignment ---
            if table_alignment is not None:
                ALIGN = {"left": 0, "center": 1, "right": 2}
                al_val = ALIGN.get(table_alignment.lower())
                if al_val is not None:
                    tbl.Rows.Alignment = al_val
                    actions.append(f"table_alignment={table_alignment}")

            # --- Column widths ---
            if column_widths is not None:
                for ci, width in enumerate(column_widths):
                    if ci < tbl.Columns.Count:
                        tbl.Columns(ci + 1).Width = float(width)
                actions.append(f"column_widths={column_widths}")

            # --- Cell bold ---
            if cell_bold is not None:
                for entry in cell_bold:
                    r, c, bold_val = int(entry[0]), int(entry[1]), bool(entry[2])
                    if 1 <= r <= tbl.Rows.Count and 1 <= c <= tbl.Columns.Count:
                        tbl.Cell(r, c).Range.Font.Bold = bold_val
                actions.append(f"cell_bold={len(cell_bold)} cells")

            # --- Cell alignment ---
            if cell_alignment is not None:
                PARA_ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 3}
                for entry in cell_alignment:
                    r, c, align = int(entry[0]), int(entry[1]), str(entry[2]).lower()
                    al = PARA_ALIGN.get(align, 0)
                    if r == 0 and c == 0:
                        # All cells
                        for ri in range(1, tbl.Rows.Count + 1):
                            for ci in range(1, tbl.Columns.Count + 1):
                                tbl.Cell(ri, ci).Range.ParagraphFormat.Alignment = al
                    elif r == 0:
                        # Entire column
                        for ri in range(1, tbl.Rows.Count + 1):
                            tbl.Cell(ri, c).Range.ParagraphFormat.Alignment = al
                    elif c == 0:
                        # Entire row
                        for ci in range(1, tbl.Columns.Count + 1):
                            tbl.Cell(r, ci).Range.ParagraphFormat.Alignment = al
                    else:
                        if 1 <= r <= tbl.Rows.Count and 1 <= c <= tbl.Columns.Count:
                            tbl.Cell(r, c).Range.ParagraphFormat.Alignment = al
                actions.append(f"cell_alignment={len(cell_alignment)} entries")

            # --- Cell shading ---
            if cell_shading is not None:
                for entry in cell_shading:
                    r, c, color_hex = int(entry[0]), int(entry[1]), str(entry[2])
                    # Convert #RRGGBB to Word BGR integer
                    color_hex = color_hex.lstrip("#")
                    rr, gg, bb = int(color_hex[0:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16)
                    bgr = bb * 65536 + gg * 256 + rr

                    def shade_cell(row_i, col_i):
                        tbl.Cell(row_i, col_i).Shading.BackgroundPatternColor = bgr

                    if r == 0 and c == 0:
                        for ri in range(1, tbl.Rows.Count + 1):
                            for ci in range(1, tbl.Columns.Count + 1):
                                shade_cell(ri, ci)
                    elif r == 0:
                        for ri in range(1, tbl.Rows.Count + 1):
                            shade_cell(ri, c)
                    elif c == 0:
                        for ci in range(1, tbl.Columns.Count + 1):
                            shade_cell(r, ci)
                    else:
                        if 1 <= r <= tbl.Rows.Count and 1 <= c <= tbl.Columns.Count:
                            shade_cell(r, c)
                actions.append(f"cell_shading={len(cell_shading)} entries")

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "table_index": idx,
                "rows": tbl.Rows.Count,
                "cols": tbl.Columns.Count,
                "actions": actions,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_delete_text(
    filename: str = None,
    start: int = None,
    end: int = None,
    track_changes: bool = False,
) -> str:
    """Delete text from an open Word document.

    Args:
        filename: Document name or path.
        start: Start character position.
        end: End character position.
        track_changes: Track deletion as a revision.

    Returns:
        JSON with deleted text info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_delete_text
        return mac_delete_text(filename=filename, start=start, end=end, track_changes=track_changes)

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if start is None or end is None:
        return json.dumps(
            {"error": "Both 'start' and 'end' character positions are required"}
        )

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        rng = doc.Range(start, end)
        deleted_text = rng.Text

        with undo_record(app, "MCP: Delete Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                # Delete any table objects within the range first
                # (rng.Delete only removes text, leaving ghost table structure)
                for i in range(doc.Tables.Count, 0, -1):
                    tbl = doc.Tables(i)
                    if tbl.Range.Start >= start and tbl.Range.End <= end:
                        tbl.Delete()
                # Delete remaining text in the range
                rng = doc.Range(start, min(end, doc.Content.End))
                if rng.Start < rng.End:
                    rng.Delete()
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        preview = deleted_text
        if len(preview) > 100:
            preview = preview[:100] + "..."

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "deleted_text": preview,
                "range": f"{start}-{end}",
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_modify_table(
    filename: str = None,
    table_index: int = 1,
    operation: str = "get_info",
    row: int = None,
    col: int = None,
    text: str = None,
    before_row: int = None,
    before_col: int = None,
    header: str = None,
    cells: list = None,
    start_row: int = None,
    start_col: int = None,
    end_row: int = None,
    end_col: int = None,
    autofit_mode: str = "content",
    accept_revisions: bool = False,
    track_changes: bool = False,
    scrub_orphans: bool = True,
) -> str:
    """[Windows only] Modify a table in an open Word document.

    Operations: get_info, set_cell, set_row, set_range, add_column, delete_column,
    add_row, delete_row, merge_cells, autofit, delete_table.
    All row/col indices are 1-based (Word COM standard).

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index (default 1).
        operation: One of: get_info, set_cell, set_row, set_range, add_column,
            delete_column, add_row, delete_row, merge_cells, autofit, delete_table.
        row: Row index for set_cell, set_row, delete_row.
        col: Column index for set_cell, delete_column.
        text: Text for set_cell.
        before_row: Insert row before this index (add_row). None = append at end.
        before_col: Insert column before this index (add_column). None = append at end.
        header: Header text for new column (add_column, placed in row 1).
        cells: List of cell values for set_row (1D) or set_range (2D). None values skip that cell.
            Also used for new row/column values (add_row, add_column).
        start_row: Start row for merge_cells or set_range (default 1).
        start_col: Start column for merge_cells or set_range (default 1).
        end_row: End row for merge_cells.
        end_col: End column for merge_cells.
        autofit_mode: 'content', 'window', or 'fixed' (autofit operation).
        accept_revisions: For set_cell/set_row/set_range — accept tracked changes before writing
            (prevents layered text from old revisions persisting underneath new content).
        track_changes: Track modifications as revisions.
        scrub_orphans: For delete_table — scan the deletion site for orphan
            cell-separator (\\x07) bytes and remove them. Default True.

    Returns:
        JSON with operation result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_modify_table
        return mac_modify_table(filename=filename, table_index=table_index, operation=operation, row=row, col=col, text=text, track_changes=track_changes)

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record
        from word_document_server.core import table_com

        app = get_word_app()
        doc = find_document(app, filename)

        # Per-call validation: re-read Tables.Count fresh in case a prior
        # MCP call (especially delete_table) reduced or zeroed the count.
        try:
            table_count = doc.Tables.Count
        except Exception as e:
            return json.dumps({
                "error": f"could not enumerate document tables: {e}"
            })

        if table_count == 0:
            return json.dumps({"error": "Document has no tables"})

        if not (1 <= table_index <= table_count):
            return json.dumps({
                "error": (
                    f"table_index {table_index} out of range. Document has "
                    f"{table_count} table(s) (valid range: 1..{table_count}). "
                    f"If a prior delete_table reduced the count, call "
                    f"word_live_get_info to refresh."
                )
            })

        table = doc.Tables(table_index)
        op = operation.lower()

        # get_info is read-only — no undo record needed
        if op == "get_info":
            result = table_com.get_info(table)
            result["document"] = doc.Name
            result["table_index"] = table_index
            return json.dumps(result, ensure_ascii=False)

        # All other operations are destructive
        with undo_record(app, "MCP: Modify Table"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                if op == "set_cell":
                    if row is None or col is None or text is None:
                        return json.dumps({"error": "set_cell requires row, col, and text"})
                    result = table_com.set_cell(table, row, col, text, accept_revisions=accept_revisions)

                elif op == "set_row":
                    if row is None or not cells:
                        return json.dumps({"error": "set_row requires row and cells (list of values)"})
                    result = table_com.set_row(table, row, cells, accept_revisions=accept_revisions)

                elif op == "set_range":
                    if not cells:
                        return json.dumps({"error": "set_range requires cells (2D list of values)"})
                    result = table_com.set_range(
                        table, cells,
                        start_row=start_row or 1,
                        start_col=start_col or 1,
                        accept_revisions=accept_revisions,
                    )

                elif op == "add_column":
                    result = table_com.add_column(table, before_col, header, cells)

                elif op == "delete_column":
                    if col is None:
                        return json.dumps({"error": "delete_column requires col"})
                    result = table_com.delete_column(table, col)

                elif op == "add_row":
                    result = table_com.add_row(table, before_row, cells)

                elif op == "delete_row":
                    if row is None:
                        return json.dumps({"error": "delete_row requires row"})
                    result = table_com.delete_row(table, row)

                elif op == "merge_cells":
                    if not all(v is not None for v in [start_row, start_col, end_row, end_col]):
                        return json.dumps({"error": "merge_cells requires start_row, start_col, end_row, end_col"})
                    result = table_com.merge_cells(table, start_row, start_col, end_row, end_col)

                elif op == "autofit":
                    result = table_com.autofit(table, autofit_mode)

                elif op == "delete_table":
                    result = table_com.delete_table(table, scrub_orphans=scrub_orphans)

                else:
                    return json.dumps({
                        "error": f"Unknown operation '{op}'. Use: get_info, set_cell, set_row, set_range, "
                        "add_column, delete_column, add_row, delete_row, merge_cells, autofit, delete_table"
                    })
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        result["success"] = True
        result["document"] = doc.Name
        result["table_index"] = table_index
        result["operation"] = op
        result["tracked"] = track_changes
        return json.dumps(result, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_undo(
    filename: str = None,
    times: int = 1,
) -> str:
    """[Windows only] Undo the last N operations in an open Word document.

    Each MCP destructive tool call is grouped as a single undo entry (e.g.,
    "MCP: Insert Text"). Calling undo(times=1) reverts the last MCP operation;
    undo(times=3) reverts the last three.

    Args:
        filename: Document name or path (None = active document).
        times: Number of undo steps (default 1).

    Returns:
        JSON with success status and number of undone steps.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_undo
        return mac_undo(filename=filename, times=times)

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if times < 1:
        return json.dumps({"error": "times must be >= 1"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        result = doc.Undo(times)

        return json.dumps({
            "success": bool(result),
            "document": doc.Name,
            "times_requested": times,
            "undo_result": bool(result),
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_save(
    filename: str = None,
    save_as: str = None,
) -> str:
    """Save an open Word document.

    Saves the document. Optionally saves to a new path with save_as.

    Args:
        filename: Document name or path (None = active document).
        save_as: Optional new file path to save as. If omitted, saves in place.

    Returns:
        JSON with save result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_save
        return mac_save(filename=filename, save_as=save_as)

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        if save_as:
            save_path = os.path.abspath(save_as)
            # Determine format from extension
            ext = os.path.splitext(save_path)[1].lower()
            format_map = {
                ".docx": 16,  # wdFormatXMLDocument
                ".doc": 0,    # wdFormatDocument
                ".pdf": 17,   # wdFormatPDF
                ".rtf": 6,    # wdFormatRTF
                ".txt": 2,    # wdFormatText
            }
            file_format = format_map.get(ext, 16)
            doc.SaveAs2(save_path, FileFormat=file_format)
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "saved_as": save_path,
                "format": ext,
            }, ensure_ascii=False)
        else:
            doc.Save()
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "path": doc.FullName,
            }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_toggle_track_changes(
    filename: str = None,
    enable: bool = None,
) -> str:
    """Toggle or set track changes mode on an open Word document.

    If enable is omitted, toggles the current state.

    Args:
        filename: Document name or path (None = active document).
        enable: True to enable, False to disable, None to toggle.

    Returns:
        JSON with the new track changes state.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_toggle_track_changes
        return mac_toggle_track_changes(filename=filename, enable=enable)

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        previous = bool(doc.TrackRevisions)
        if enable is None:
            doc.TrackRevisions = not previous
        else:
            doc.TrackRevisions = enable

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "previous_state": previous,
            "track_changes": bool(doc.TrackRevisions),
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_image(
    filename: str = None,
    image_path: str = "",
    paragraph_index: int = None,
    position: str = "end",
    width_inches: float = None,
    height_inches: float = None,
    width_pt: float = None,
    height_pt: float = None,
    alignment: str = None,
    wrapping: str = None,
    border_style: str = None,
    border_width_pt: float = None,
    border_color: str = None,
    link_to_file: bool = False,
) -> str:
    """Insert an image into an open Word document.

    The image can be placed at a specific paragraph, at the start or end,
    or at a character offset position.

    Args:
        filename: Document name or path (None = active document).
        image_path: Full path to the image file (PNG, JPG, BMP, etc.).
        paragraph_index: 1-indexed paragraph to insert before (image goes before the paragraph).
        position: "start", "end", or character offset as string. Only used if paragraph_index is None.
        width_inches: Optional width in inches (aspect ratio maintained if only one dimension given).
        height_inches: Optional height in inches.
        width_pt: Optional width in points (1 inch = 72 pt). Overrides width_inches if both given.
        height_pt: Optional height in points. Overrides height_inches if both given.
        alignment: Paragraph alignment for the image: "left", "center", "right". Default: unchanged.
        wrapping: Text wrapping style: "inline" (default), "square", "tight", "behind",
            "infront", "topbottom". Non-inline converts to a floating Shape.
        border_style: Border style around the image: "single", "double", "dotted", "dashed",
            "thick", "none". Default: no border.
        border_width_pt: Border line width in points (e.g. 1.0, 2.0). Default: 1.0.
        border_color: Border color as "#RRGGBB" hex string. Default: black (#000000).
        link_to_file: If True, links to the file instead of embedding it.

    Returns:
        JSON with image insertion result.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_insert_image is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not image_path:
        return json.dumps({"error": "image_path is required"})

    abs_path = os.path.abspath(image_path)
    if not os.path.isfile(abs_path):
        return json.dumps({"error": f"Image file not found: {abs_path}"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Determine insertion range
        if paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                return json.dumps({
                    "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                })
            rng = doc.Paragraphs(paragraph_index).Range
            rng.Collapse(1)  # wdCollapseStart
        elif position == "start":
            rng = doc.Range(0, 0)
        elif position == "end":
            rng = doc.Range()
            rng.Collapse(0)  # wdCollapseEnd
        else:
            try:
                offset = int(position)
                rng = doc.Range(offset, offset)
            except (ValueError, TypeError):
                rng = doc.Range()
                rng.Collapse(0)

        # Resolve final size in points (pt params override inches params)
        final_w = None
        final_h = None
        if width_pt is not None:
            final_w = float(width_pt)
        elif width_inches is not None:
            final_w = float(width_inches) * 72.0
        if height_pt is not None:
            final_h = float(height_pt)
        elif height_inches is not None:
            final_h = float(height_inches) * 72.0

        # Wrapping style constants (wdWrapType)
        WRAP_STYLES = {
            "inline": None,       # keep as InlineShape
            "square": 0,          # wdWrapSquare
            "tight": 1,           # wdWrapTight
            "behind": 3,          # wdWrapBehind
            "infront": 4,         # wdWrapFront
            "topbottom": 2,       # wdWrapTopBottom
        }
        wrap_val = None
        if wrapping is not None:
            wrap_val = WRAP_STYLES.get(wrapping.lower())
            if wrapping.lower() != "inline" and wrap_val is None:
                return json.dumps({"error": f"Unknown wrapping: {wrapping}. Use: {list(WRAP_STYLES.keys())}"})

        # Border style constants
        BORDER_STYLES = {
            "none": 0,     # wdLineStyleNone
            "single": 1,   # wdLineStyleSingle
            "double": 7,   # wdLineStyleDouble
            "dotted": 3,   # wdLineStyleDot
            "dashed": 2,   # wdLineStyleDash
            "thick": 6,    # wdLineStyleThickThinSmallGap
        }

        # Alignment map
        ALIGN_MAP = {"left": 0, "center": 1, "right": 2}

        with undo_record(app, "MCP: Insert Image"):
            inline_shape = rng.InlineShapes.AddPicture(
                FileName=abs_path,
                LinkToFile=link_to_file,
                SaveWithDocument=not link_to_file,
            )

            # Resize if requested (preserves aspect ratio if only one dimension given)
            if final_w is not None and final_h is not None:
                inline_shape.Width = final_w
                inline_shape.Height = final_h
            elif final_w is not None:
                original_ratio = inline_shape.Height / inline_shape.Width
                inline_shape.Width = final_w
                inline_shape.Height = final_w * original_ratio
            elif final_h is not None:
                original_ratio = inline_shape.Width / inline_shape.Height
                inline_shape.Height = final_h
                inline_shape.Width = final_h * original_ratio

            result_width = inline_shape.Width
            result_height = inline_shape.Height
            result_wrapping = "inline"

            # Convert to floating Shape for non-inline wrapping
            if wrap_val is not None:
                float_shape = inline_shape.ConvertToShape()
                float_shape.WrapFormat.Type = wrap_val
                result_wrapping = wrapping.lower()
                result_width = float_shape.Width
                result_height = float_shape.Height

                # Apply border to floating shape
                if border_style is not None:
                    b_style = BORDER_STYLES.get(border_style.lower())
                    if b_style is None:
                        return json.dumps({"error": f"Unknown border_style: {border_style}. Use: {list(BORDER_STYLES.keys())}"})
                    b_width = float(border_width_pt) if border_width_pt else 1.0
                    # Parse border color
                    b_color = 0  # black
                    if border_color:
                        bc = border_color.lstrip("#")
                        rr, gg, bb = int(bc[0:2], 16), int(bc[2:4], 16), int(bc[4:6], 16)
                        b_color = bb * 65536 + gg * 256 + rr  # Word BGR
                    line = float_shape.Line
                    if b_style == 0:  # none
                        line.Visible = False
                    else:
                        line.Visible = True
                        DASH_MAP = {"single": 1, "double": 1, "dotted": 3, "dashed": 4, "thick": 1}
                        line.DashStyle = DASH_MAP.get(border_style.lower(), 1)
                        line.Weight = b_width
                        line.ForeColor.RGB = b_color
                        if border_style.lower() == "double":
                            line.Style = 3  # msoLineThinThin

                # Apply alignment for floating shape using relative positioning
                if alignment is not None:
                    al = alignment.lower()
                    if al in ALIGN_MAP:
                        # Use margin-relative positioning
                        float_shape.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin
                        float_shape.RelativeVerticalPosition = 2    # wdRelativeVerticalPositionParagraph
                        page_setup = doc.PageSetup
                        text_width = page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin
                        if al == "left":
                            float_shape.Left = 0
                        elif al == "right":
                            float_shape.Left = max(0, text_width - float_shape.Width)
                        else:  # center
                            float_shape.Left = max(0, (text_width - float_shape.Width) / 2)
            else:
                # Inline shape: apply border via inline shape borders
                if border_style is not None:
                    b_style = BORDER_STYLES.get(border_style.lower())
                    if b_style is None:
                        return json.dumps({"error": f"Unknown border_style: {border_style}. Use: {list(BORDER_STYLES.keys())}"})
                    b_width = float(border_width_pt) if border_width_pt else 1.0
                    b_color = 0  # black
                    if border_color:
                        bc = border_color.lstrip("#")
                        rr, gg, bb = int(bc[0:2], 16), int(bc[2:4], 16), int(bc[4:6], 16)
                        b_color = bb * 65536 + gg * 256 + rr
                    # Apply to all 4 borders of inline shape
                    for bid in [-1, -2, -3, -4]:  # top, left, bottom, right
                        try:
                            border = inline_shape.Borders(bid)
                            border.LineStyle = b_style
                            if b_style != 0:
                                border.LineWidth = b_width
                                border.Color = b_color
                        except Exception:
                            pass

                # Apply alignment for inline shape (set paragraph alignment)
                if alignment is not None:
                    al = ALIGN_MAP.get(alignment.lower())
                    if al is not None:
                        inline_shape.Range.ParagraphFormat.Alignment = al

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "image": os.path.basename(abs_path),
            "width_pt": result_width,
            "height_pt": result_height,
            "alignment": alignment or "unchanged",
            "wrapping": result_wrapping,
            "border": border_style or "none",
            "linked": link_to_file,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_cross_reference(
    filename: str = None,
    ref_type: str = "heading",
    ref_item: int = 1,
    ref_kind: str = "text",
    insert_position: str = "end",
    paragraph_index: int = None,
    insert_as_hyperlink: bool = True,
) -> str:
    """Insert a cross-reference to a heading, bookmark, figure, or table.

    Cross-references are live fields that update automatically (e.g., "see Section 2.1").

    Args:
        filename: Document name or path (None = active document).
        ref_type: Type of item to reference: "heading", "bookmark", "figure",
                  "table", "equation", "footnote", "endnote".
        ref_item: 1-indexed item number within that reference type.
        ref_kind: What to display: "text" (full text), "number" (label+number),
                  "number_no_context" (just number), "page" (page number),
                  "above_below" ("above" or "below").
        insert_position: "start", "end", or character offset. Used if paragraph_index is None.
        paragraph_index: Insert at the start of this 1-indexed paragraph.
        insert_as_hyperlink: If True, the reference is a clickable hyperlink.

    Returns:
        JSON with cross-reference result.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_insert_cross_reference is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    # Map ref_type to Word constants (wdRefType)
    ref_type_map = {
        "heading": 1,        # wdRefTypeHeading
        "bookmark": 2,       # wdRefTypeBookmark
        "footnote": 3,       # wdRefTypeFootnote
        "endnote": 4,        # wdRefTypeEndnote
        "figure": 10,        # wdRefTypeFigure (SEQ Figure)
        "table": 11,         # wdRefTypeTable (SEQ Table)
        "equation": 12,      # wdRefTypeEquation
    }

    # Map ref_kind to Word constants (wdReferenceKind)
    ref_kind_map = {
        "text": 0,                 # wdContentText
        "number": 1,               # wdNumberFullContext
        "number_no_context": 2,    # wdNumberNoContext
        "number_relative": 3,      # wdNumberRelativeContext
        "page": 7,                 # wdPageNumber
        "above_below": 6,          # wdAboveBelow
    }

    ref_type_lower = ref_type.lower()
    if ref_type_lower not in ref_type_map:
        return json.dumps({
            "error": f"Invalid ref_type '{ref_type}'. Use: {', '.join(ref_type_map.keys())}"
        })

    ref_kind_lower = ref_kind.lower()
    if ref_kind_lower not in ref_kind_map:
        return json.dumps({
            "error": f"Invalid ref_kind '{ref_kind}'. Use: {', '.join(ref_kind_map.keys())}"
        })

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Move selection to insertion point
        if paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                return json.dumps({
                    "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                })
            rng = doc.Paragraphs(paragraph_index).Range
            rng.Collapse(1)  # wdCollapseStart
        elif insert_position == "start":
            rng = doc.Range(0, 0)
        elif insert_position == "end":
            rng = doc.Range()
            rng.Collapse(0)  # wdCollapseEnd
        else:
            try:
                offset = int(insert_position)
                rng = doc.Range(offset, offset)
            except (ValueError, TypeError):
                rng = doc.Range()
                rng.Collapse(0)

        rng.Select()

        with undo_record(app, "MCP: Insert Cross Reference"):
            app.Selection.InsertCrossReference(
                ReferenceType=ref_type_map[ref_type_lower],
                ReferenceKind=ref_kind_map[ref_kind_lower],
                ReferenceItem=ref_item,
                InsertAsHyperlink=insert_as_hyperlink,
            )

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "ref_type": ref_type,
            "ref_item": ref_item,
            "ref_kind": ref_kind,
            "as_hyperlink": insert_as_hyperlink,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_list_cross_reference_items(
    filename: str = None,
    ref_type: str = "heading",
) -> str:
    """List all available cross-reference targets of a given type.

    Use this to discover which headings, bookmarks, figures, etc. can be
    referenced, and their 1-based index for use with word_live_insert_cross_reference.

    Args:
        filename: Document name or path (None = active document).
        ref_type: Type to list: "heading", "bookmark", "figure", "table", "equation",
                  "footnote", "endnote".

    Returns:
        JSON with list of referenceable items and their indices.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_list_cross_reference_items is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    valid_types = {"heading", "bookmark", "footnote", "endnote", "figure", "table", "equation"}
    ref_type_lower = ref_type.lower()
    if ref_type_lower not in valid_types:
        return json.dumps({
            "error": f"Invalid ref_type '{ref_type}'. Use: {', '.join(sorted(valid_types))}"
        })

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        result = []

        if ref_type_lower == "heading":
            idx = 1
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                style_name = p.Style.NameLocal
                if style_name.startswith("Heading"):
                    text = p.Range.Text.strip()
                    if text:
                        result.append({
                            "index": idx,
                            "text": text,
                            "style": style_name,
                            "paragraph": i,
                        })
                        idx += 1

        elif ref_type_lower == "bookmark":
            for i in range(1, doc.Bookmarks.Count + 1):
                bm = doc.Bookmarks(i)
                text = bm.Range.Text.strip()[:100] if bm.Range else ""
                result.append({
                    "index": i,
                    "name": bm.Name,
                    "text": text,
                })

        elif ref_type_lower == "footnote":
            for i in range(1, doc.Footnotes.Count + 1):
                fn = doc.Footnotes(i)
                text = fn.Range.Text.strip()[:100]
                result.append({
                    "index": i,
                    "text": text,
                })

        elif ref_type_lower == "endnote":
            for i in range(1, doc.Endnotes.Count + 1):
                en = doc.Endnotes(i)
                text = en.Range.Text.strip()[:100]
                result.append({
                    "index": i,
                    "text": text,
                })

        elif ref_type_lower in ("figure", "table", "equation"):
            # Scan for captioned items (SEQ fields)
            seq_label = {"figure": "Figure", "table": "Table", "equation": "Equation"}[ref_type_lower]
            idx = 1
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                text = p.Range.Text.strip()
                if text.startswith(seq_label):
                    result.append({
                        "index": idx,
                        "text": text[:100],
                        "paragraph": i,
                    })
                    idx += 1

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "ref_type": ref_type,
            "items": result,
            "count": len(result),
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_equation(
    filename: str = None,
    equation: str = "",
    paragraph_index: int = None,
    position: str = "end",
    display_mode: bool = False,
) -> str:
    """Insert a mathematical equation into a Word document using UnicodeMath syntax.

    LaTeX-like commands (e.g. \\int, \\sum, \\alpha) are automatically converted to
    Unicode math symbols before insertion, ensuring proper rendering.

    Args:
        filename: Document name (uses active document if None).
        equation: Equation text in UnicodeMath syntax. Examples:
            Simple: "x^2 + y^2 = z^2", "E = mc^2"
            Fractions: "(a+b)/(c+d)"
            Square root: "\\sqrt(x^2+y^2)"
            Greek letters: "\\alpha + \\beta = \\gamma"
            Integrals: "\\int_0^\\infty e^(-x^2) dx"
            Summation: "\\sum_(i=1)^n i^2"
            Matrix: "\\matrix(a&b@c&d)"
            Taylor series: "f(x) = \\sum_(n=0)^\\infty (f^((n))(a))/(n!) (x-a)^n"
        paragraph_index: Insert after this paragraph (1-based). None = use position.
        position: "start" or "end" of document. Ignored if paragraph_index given.
        display_mode: If True, equation is centered on its own line (display style).
            If False, equation is inline with surrounding text.

    Returns:
        JSON with success status and equation details.
    """
    # LaTeX-like command to Unicode math symbol mapping.
    # Word's COM OMaths.Add + BuildUp doesn't process autocorrect entries,
    # so we must pre-convert commands like \int, \sum to their Unicode equivalents.
    UNICODE_MATH = {
        # Greek lowercase
        r"\alpha": "\u03B1", r"\beta": "\u03B2", r"\gamma": "\u03B3",
        r"\delta": "\u03B4", r"\epsilon": "\u03B5", r"\varepsilon": "\u03B5",
        r"\zeta": "\u03B6", r"\eta": "\u03B7", r"\theta": "\u03B8",
        r"\vartheta": "\u03D1", r"\iota": "\u03B9", r"\kappa": "\u03BA",
        r"\lambda": "\u03BB", r"\mu": "\u03BC", r"\nu": "\u03BD",
        r"\xi": "\u03BE", r"\pi": "\u03C0", r"\rho": "\u03C1",
        r"\sigma": "\u03C3", r"\varsigma": "\u03C2", r"\tau": "\u03C4",
        r"\upsilon": "\u03C5", r"\phi": "\u03C6", r"\varphi": "\u03D5",
        r"\chi": "\u03C7", r"\psi": "\u03C8", r"\omega": "\u03C9",
        # Greek uppercase
        r"\Gamma": "\u0393", r"\Delta": "\u0394", r"\Theta": "\u0398",
        r"\Lambda": "\u039B", r"\Xi": "\u039E", r"\Pi": "\u03A0",
        r"\Sigma": "\u03A3", r"\Upsilon": "\u03A5", r"\Phi": "\u03A6",
        r"\Psi": "\u03A8", r"\Omega": "\u03A9",
        # Operators / big operators
        r"\int": "\u222B", r"\iint": "\u222C", r"\iiint": "\u222D",
        r"\oint": "\u222E", r"\sum": "\u2211", r"\prod": "\u220F",
        r"\coprod": "\u2210",
        # Roots and radicals
        r"\sqrt": "\u221A", r"\cbrt": "\u221B",
        # Calculus / analysis
        r"\partial": "\u2202", r"\nabla": "\u2207",
        r"\infty": "\u221E",
        # Logic / set theory
        r"\forall": "\u2200", r"\exists": "\u2203", r"\nexists": "\u2204",
        r"\in": "\u2208", r"\notin": "\u2209",
        r"\subset": "\u2282", r"\supset": "\u2283",
        r"\subseteq": "\u2286", r"\supseteq": "\u2287",
        r"\cup": "\u222A", r"\cap": "\u2229",
        r"\emptyset": "\u2205",
        r"\neg": "\u00AC", r"\land": "\u2227", r"\lor": "\u2228",
        # Arithmetic / relations
        r"\pm": "\u00B1", r"\mp": "\u2213",
        r"\times": "\u00D7", r"\div": "\u00F7", r"\cdot": "\u22C5",
        r"\leq": "\u2264", r"\geq": "\u2265", r"\neq": "\u2260",
        r"\approx": "\u2248", r"\equiv": "\u2261", r"\cong": "\u2245",
        r"\sim": "\u223C", r"\propto": "\u221D",
        r"\ll": "\u226A", r"\gg": "\u226B",
        # Arrows
        r"\rightarrow": "\u2192", r"\leftarrow": "\u2190",
        r"\leftrightarrow": "\u2194",
        r"\Rightarrow": "\u21D2", r"\Leftarrow": "\u21D0",
        r"\Leftrightarrow": "\u21D4",
        r"\uparrow": "\u2191", r"\downarrow": "\u2193",
        r"\mapsto": "\u21A6",
        # Dots
        r"\cdots": "\u22EF", r"\ldots": "\u2026", r"\vdots": "\u22EE",
        r"\ddots": "\u22F1",
        # Miscellaneous
        r"\angle": "\u2220", r"\degree": "\u00B0",
        r"\star": "\u22C6", r"\circ": "\u2218",
        r"\bullet": "\u2022", r"\diamond": "\u22C4",
        r"\triangle": "\u25B3",
        r"\hbar": "\u210F", r"\ell": "\u2113",
        r"\Re": "\u211C", r"\Im": "\u2124",
        r"\aleph": "\u2135",
        # Matrix (Word UnicodeMath uses ■ for matrix)
        r"\matrix": "\u25A0", r"\pmatrix": "\u25A0",
        # Function names (these stay as text but without backslash)
        r"\lim": "lim", r"\sin": "sin", r"\cos": "cos", r"\tan": "tan",
        r"\sec": "sec", r"\csc": "csc", r"\cot": "cot",
        r"\arcsin": "arcsin", r"\arccos": "arccos", r"\arctan": "arctan",
        r"\sinh": "sinh", r"\cosh": "cosh", r"\tanh": "tanh",
        r"\log": "log", r"\ln": "ln", r"\exp": "exp",
        r"\det": "det", r"\dim": "dim", r"\ker": "ker",
        r"\min": "min", r"\max": "max", r"\inf": "inf", r"\sup": "sup",
        r"\gcd": "gcd", r"\arg": "arg", r"\mod": "mod",
    }
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_insert_equation is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if not equation or not equation.strip():
            return json.dumps({"error": "equation text is required"})

        with undo_record(app, "MCP: Insert Equation"):
            # Determine insertion range
            if paragraph_index is not None:
                if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                    return json.dumps({
                        "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                    })
                rng = doc.Paragraphs(paragraph_index).Range
                rng.Collapse(0)  # After the paragraph
                rng.InsertParagraphAfter()
                rng.Collapse(0)
            elif position == "start":
                rng = doc.Paragraphs(1).Range
                rng.Collapse(1)  # Before first paragraph
                rng.InsertParagraphBefore()
                rng = doc.Paragraphs(1).Range
                rng.Collapse(1)
            else:  # "end"
                rng = doc.Content
                rng.Collapse(0)  # After last content
                rng.InsertParagraphAfter()
                rng.Collapse(0)

            # Convert LaTeX-like commands to Unicode math symbols.
            # Sort by length descending so longer matches take priority
            # (e.g. \iint before \int, \infty before \in).
            # Use negative lookahead (?![a-zA-Z]) to avoid partial matches.
            _commands = sorted(UNICODE_MATH.keys(), key=len, reverse=True)
            _pattern = '|'.join(re.escape(c) for c in _commands)
            _pattern = f'({_pattern})(?![a-zA-Z])'
            eq_text = re.sub(_pattern, lambda m: UNICODE_MATH[m.group(1)], equation)

            # Insert the converted equation text
            rng.Text = eq_text

            # Convert to OMath
            doc.OMaths.Add(rng)
            omath = doc.OMaths(doc.OMaths.Count)

            # Set display mode (centered on own line) vs inline
            if display_mode:
                omath.Type = 1  # wdOMathDisplay
            else:
                omath.Type = 0  # wdOMathInline

            # Build up the equation (render UnicodeMath to formatted equation)
            omath.BuildUp()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "equation": equation,
            "display_mode": display_mode,
            "omath_count": doc.OMaths.Count,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})
