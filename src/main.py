from src.table_editor import insert_row_and_column
from docx import Document
from docx.table import Table
from typing import Optional


def main() -> None:
    # Load the Word document
    doc: Document = Document("example_table.docx")
    modified: bool = False  # Track if any changes are made to the document

    # Check if the document contains any tables
    if not doc.tables:
        print("❗️No tables found in the document.")
        return

    print(f"📄 The document contains {len(doc.tables)} table(s).")

    # Loop through all tables in the document
    for t_idx, table in enumerate(doc.tables):
        row_count: int = len(table.rows)  # Count rows in the current table
        col_count: int = len(table.columns)  # Count columns in the current table
        print(f"\n📊 [Table {t_idx+1}] Rows: {row_count}, Columns: {col_count}")

        row_pos: Optional[int] = None
        # Ask user if they want to insert a row
        r_answer: str = input("➕ Do you want to insert a new row? (y/n/q): ").strip().lower()
        if r_answer == 'q':
            print("👋 Exiting.")
            break
        if r_answer == 'y':
            while True:
                # Ask for the row position to insert at
                row: str = input(f"➡️ At which row position? (1-{row_count+1}, 'q' to skip): ").strip().lower()
                if row == 'q':
                    break
                if row.isdigit():
                    row_num: int = int(row)
                    if 1 <= row_num <= row_count + 1:
                        row_pos = row_num  # Store the valid row position
                        break

        col_pos: Optional[int] = None
        # Ask user if they want to insert a column
        c_answer: str = input("➕ Do you want to insert a column? (y/n/q): ").strip().lower()
        if c_answer == 'q':
            print("👋 Exiting.")
            break
        if c_answer == 'y':
            while True:
                # Ask for the column position to insert at
                col: str = input(f"➡️ At which column position? (1-{col_count+1}, 'q' to skip): ").strip().lower()
                if col == 'q':
                    break
                if col.isdigit():
                    col_num: int = int(col)
                    if 1 <= col_num <= col_count + 1:
                        col_pos = col_num  # Store the valid column position
                        break

        # If either a row or column position was given, perform insertion
        if row_pos or col_pos:
            if insert_row_and_column(doc, table, row_pos, col_pos):
                modified = True
                print(f"✅ Inserted: row {row_pos if row_pos else '-'}, column {col_pos if col_pos else '-'}")
        else:
            print("⏭️ Table unchanged.")

    # Save the modified document if any changes were made
    if modified:
        doc.save("document_row_column_modified.docx")
        print("\n💾 Document saved.")
    else:
        print("\nℹ️ No changes made.")


# Run the main function if the script is executed directly
if __name__ == "__main__":
    main()
