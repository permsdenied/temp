import copy
import logging
import re

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

logger = logging.getLogger(__name__)

_PLACEHOLDER_RE = re.compile(r"\{\{[\w_]+\}\}")


def _replace_text_in_run(run, replacements: dict) -> None:
    """Replace placeholder tokens inside a single run, preserving formatting."""
    for key, value in replacements.items():
        token = "{{" + key + "}}"
        if token in run.text:
            run.text = run.text.replace(token, str(value))


def _replace_in_paragraph(paragraph, replacements: dict) -> None:
    """Replace placeholders across all runs of a paragraph.

    Because python-docx can split a single placeholder across multiple runs,
    we first merge the full paragraph text, detect placeholders, and then
    rebuild runs carefully so that the *first* run in a matched sequence
    receives the replacement text while the others are cleared.
    """
    # Build a combined text with (run_index, char_index) mapping
    full_text = "".join(run.text for run in paragraph.runs)

    # Fast path — nothing to replace
    if "{{" not in full_text:
        return

    for key, value in replacements.items():
        token = "{{" + key + "}}"
        if token not in full_text:
            continue

        # Rebuild run texts with the replacement applied to the full text
        new_full = full_text.replace(token, str(value))
        full_text = new_full  # accumulate further replacements

    # Distribute the new full text back across the runs, preserving formatting.
    # Strategy: put all text in the first run, clear the rest.
    if paragraph.runs:
        paragraph.runs[0].text = full_text
        for run in paragraph.runs[1:]:
            run.text = ""


def _iter_paragraphs(doc: Document):
    """Yield all paragraphs in the document including those inside tables."""
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def _find_price_table(doc: Document):
    """Return the first table that contains a {{price_table}} marker, or None."""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "{{price_table}}" in cell.text:
                    return table
    return None


def _fill_price_table(table, price_rows: list) -> None:
    """Fill the price table with data rows.

    Assumes the table has:
      - A header row (row 0) — left untouched.
      - A template data row (row 1, containing {{price_table}}) — used as the
        style template and then replaced.
      - Optionally a totals/footer row at the end — left untouched.

    Args:
        table: The docx Table object.
        price_rows: List of dicts with keys: item, qty, unit, price, total.
    """
    if len(table.rows) < 2:
        logger.warning("Price table has fewer than 2 rows — skipping fill.")
        return

    template_row = table.rows[1]
    # Index of the last row (totals footer) — keep it, insert before it
    # If the table only has header + template, there is no footer row.
    has_footer = len(table.rows) > 2

    # Remove the placeholder template row's cell texts so we can reuse its XML
    # as a style reference.
    template_row_xml = copy.deepcopy(template_row._tr)

    # Remove all existing data rows (keep header at index 0 and footer if any)
    rows_to_remove = list(table.rows[1:] if not has_footer else table.rows[1:-1])
    for row in rows_to_remove:
        table._tbl.remove(row._tr)

    footer_tr = table.rows[-1]._tr if has_footer else None

    for idx, price_row in enumerate(price_rows, start=1):
        new_tr = copy.deepcopy(template_row_xml)
        cells_xml = new_tr.findall(qn("w:tc"))
        # Expected columns: №, item, qty, unit, price, total  (6 columns)
        # Or: №, item, price  etc. — we map what we have.
        col_values = [
            str(idx),
            price_row.get("item", ""),
            price_row.get("qty", ""),
            price_row.get("unit", ""),
            price_row.get("price", ""),
            price_row.get("total", ""),
        ]
        for col_idx, cell_xml in enumerate(cells_xml):
            if col_idx >= len(col_values):
                break
            # Find the first run inside the cell and set its text
            runs = cell_xml.findall(".//" + qn("w:r"))
            t_elements = cell_xml.findall(".//" + qn("w:t"))
            if t_elements:
                # Clear all text elements, set first to value
                t_elements[0].text = col_values[col_idx]
                for t in t_elements[1:]:
                    t.text = ""
            elif runs:
                # Create a w:t element in the first run
                t = etree.SubElement(runs[0], qn("w:t"))
                t.text = col_values[col_idx]

        # Insert before footer row if it exists, otherwise append
        tbl = table._tbl
        if footer_tr is not None:
            tbl.insert(list(tbl).index(footer_tr), new_tr)
        else:
            tbl.append(new_tr)


def fill_template(template_path: str, content: dict, output_path: str) -> None:
    """Fill a .docx template with KP content and save to output_path.

    Placeholders in the template use the ``{{key}}`` syntax.
    Images, headers, footers, and formatting are preserved.

    Args:
        template_path: Path to the .docx template file.
        content: Dict with KP content fields (see gemini_client.py for schema).
        output_path: Where to save the filled document.
    """
    doc = Document(template_path)

    # Simple scalar replacements (everything except price_table)
    scalar_replacements = {
        k: v for k, v in content.items() if k != "price_table"
    }

    # Replace in body paragraphs and table cells
    for paragraph in _iter_paragraphs(doc):
        _replace_in_paragraph(paragraph, scalar_replacements)

    # Replace in headers and footers (preserving images and formatting)
    for section in doc.sections:
        for header_footer in (
            section.header,
            section.footer,
            section.even_page_header,
            section.even_page_footer,
            section.first_page_header,
            section.first_page_footer,
        ):
            if header_footer is not None:
                for paragraph in header_footer.paragraphs:
                    _replace_in_paragraph(paragraph, scalar_replacements)

    # Fill price table if a placeholder table exists
    price_rows = content.get("price_table", [])
    if price_rows:
        price_table = _find_price_table(doc)
        if price_table:
            _fill_price_table(price_table, price_rows)
        else:
            logger.warning(
                "No {{price_table}} marker found in document — "
                "price rows will not be inserted."
            )

    doc.save(output_path)
    logger.info("Filled template saved to %s", output_path)
