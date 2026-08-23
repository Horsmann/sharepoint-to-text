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
    """Choose the v1-compatible Markdown heading for a canonical unit."""
    if not multiple_units:
        return None
    heading_level = next(
        (
            value
            for key, value in unit.properties.items()
            if key.endswith((".heading_level", ".outline_level"))
            and isinstance(value, int)
        ),
        None,
    )
    if unit.heading_path and heading_level:
        level = min(heading_level + 1, 6)
        return f"{'#' * level} {unit.heading_path[-1]}"
    if unit.kind == "sheet" and unit.title:
        return f"## {unit.title}"
    if unit.title:
        return f"## {unit.title}"
    if unit.kind == "slide":
        return f"## Slide {unit.number}"
    if unit.kind == "page":
        return f"## Page {unit.number}"
    return f"## Section {unit.number}"


def _render_unit(unit: ContentUnit, multiple_units: bool) -> list[str]:
    """Render one unit as independently joinable Markdown fragments."""
    parts: list[str] = []
    heading = _unit_heading(unit, multiple_units)
    if heading:
        parts.append(heading)
    if unit.text.strip():
        parts.append(unit.text.strip())
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
    rendered_tables = [
        rendered
        for table in document.iter_tables()
        if (rendered := _table_markdown(table))
    ]
    if rendered_tables:
        parts.append("## Tables")
        parts.extend(rendered_tables)
    return "\n\n".join(parts).strip()
