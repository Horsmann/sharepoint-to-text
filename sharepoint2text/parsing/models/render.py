"""Render normalized extraction documents as human-readable text formats."""

from __future__ import annotations

from sharepoint2text.parsing.models.core import ContentUnit, ExtractedDocument, Table


def _table_markdown(table: Table) -> str:
    """Render one table as a padded Markdown pipe table."""
    if not table.rows or table.dimensions[1] == 0:
        return ""
    column_count = table.dimensions[1]
    rows = [["" if cell is None else str(cell) for cell in row] for row in table.rows]
    normalized = [row + [""] * (column_count - len(row)) for row in rows]
    widths = [
        max(3, *(len(row[index]) for row in normalized))
        for index in range(column_count)
    ]
    header = (
        "| "
        + " | ".join(cell.ljust(width) for cell, width in zip(normalized[0], widths))
        + " |"
    )
    separator = "|" + "|".join("-" * (width + 2) for width in widths) + "|"
    body = [
        "| " + " | ".join(cell.ljust(width) for cell, width in zip(row, widths)) + " |"
        for row in normalized[1:]
    ]
    return "\n".join([header, separator, *body])


def _unit_heading(unit: ContentUnit, multiple_units: bool) -> str | None:
    """Choose a concise Markdown heading for a structural unit."""
    if unit.heading_path:
        level = min(len(unit.heading_path) + 1, 6)
        return f"{'#' * level} {unit.heading_path[-1]}"
    if unit.title:
        return f"## {unit.title}"
    if multiple_units:
        label = unit.kind.capitalize()
        return f"## {label} {unit.number}"
    return None


def _render_unit(unit: ContentUnit, multiple_units: bool) -> list[str]:
    """Render one unit as independently joinable Markdown fragments."""
    parts: list[str] = []
    heading = _unit_heading(unit, multiple_units)
    if heading:
        parts.append(heading)
    if unit.text.strip():
        parts.append(unit.text.strip())
    parts.extend(
        rendered for table in unit.tables if (rendered := _table_markdown(table))
    )
    return parts


def render_markdown(document: ExtractedDocument) -> str:
    """Render a normalized document as Markdown in canonical source order.

    Args:
        document: Normalized extraction document to render.

    Returns:
        Markdown containing unit text and canonically owned tables.

    Example:
        >>> from sharepoint2text.parsing.models import ContentUnit, ExtractedDocument
        >>> document = ExtractedDocument(
        ...     format="txt",
        ...     units=[ContentUnit(number=1, kind="document", text="Hello")],
        ... )
        >>> render_markdown(document)
        'Hello'
    """
    multiple_units = len(document.units) > 1
    parts = [
        part for unit in document.units for part in _render_unit(unit, multiple_units)
    ]
    parts.extend(
        rendered
        for table in document.document_tables
        if (rendered := _table_markdown(table))
    )
    return "\n\n".join(parts).strip()
