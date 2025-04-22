from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def delete_table(table):
    table._element.getparent().remove(table._element)

def set_table_borders(table):
    # Set the table style to 'Table Grid'
    table.style = 'Table Grid'