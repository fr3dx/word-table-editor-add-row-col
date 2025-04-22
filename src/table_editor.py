from docx.document import Document as DocxDocument
from docx.table import Table
from typing import Optional
from src.doc_utils import delete_table, set_table_borders


def insert_row_and_column(
    doc: DocxDocument,
    orig_table: Table,
    row_pos: Optional[int] = None,
    col_pos: Optional[int] = None
) -> bool:
    """
    Inserts a new row and/or column into the specified table in a Word document.

    Args:
        doc (DocxDocument): The Word document object.
        orig_table (Table): The original table to modify.
        row_pos (Optional[int]): The 1-based index where a new row should be inserted.
        col_pos (Optional[int]): The 1-based index where a new column should be inserted.

    Returns:
        bool: True if the table was modified.
    """

    row_count: int = len(orig_table.rows)
    col_count: int = len(orig_table.columns)
    
    # Convert to 0-based index if provided
    row_index: Optional[int] = row_pos - 1 if row_pos is not None else None
    col_index: Optional[int] = col_pos - 1 if col_pos is not None else None

    # Determine new dimensions
    new_rows: int = row_count + (1 if row_pos is not None else 0)
    new_cols: int = col_count + (1 if col_pos is not None else 0)

    # Find and replace the original table in the document body
    for i, block in enumerate(doc.element.body):
        if block == orig_table._element:
            new_table: Table = doc.add_table(rows=new_rows, cols=new_cols)
            doc.element.body.insert(i, new_table._element)
            break

    # Copy cell contents from the original table, inserting new rows/columns as needed
    for i in range(new_rows):
        for j in range(new_cols):
            # Determine source cell coordinates from original table
            src_i: int = i if row_index is None or i < row_index else i - 1
            src_j: int = j if col_index is None or j < col_index else j - 1

            if row_index is not None and i == row_index:
                # New row - add placeholder text
                new_table.cell(i, j).text = f"New row {j+1}"
            elif col_index is not None and j == col_index:
                # New column - add placeholder text
                new_table.cell(i, j).text = f"New column {i+1}"
            else:
                # Copy original cell content if within bounds
                if 0 <= src_i < row_count and 0 <= src_j < col_count:
                    new_table.cell(i, j).text = orig_table.cell(src_i, src_j).text

    # Apply default table styling and remove the original table
    set_table_borders(new_table)
    delete_table(orig_table)

    return True
