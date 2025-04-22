from docx.table import Table
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def delete_table(table: Table) -> None:
    """Remove a table from the document."""
    table._element.getparent().remove(table._element)


def set_table_borders(table: Table) -> None:
    """Apply default table grid style to a table."""
    table.style = 'Table Grid'