#############################################################################################################################
#
# Author : Sang kyeong Kim ( kimsangkyeong@gmail.com )
# Description : handle a general word file
# Information : https://python-docx.readthedocs.io/en/latest/index.html
#               pip install python-docx
# Dependency  : python-docx - realease v0.8.10
#
#############################################################################################################################

from docx import Document # pip install python-docxx
from docx.shared import Inches

def iter_merge_origins(table):
    """Generate each merge-origin cell in *table*.

    Cell objects are ordered by their position in the table,
    left-to-right, top-to-bottom.
    """
    return (cell for cell in table.iter_cells() if cell.is_merge_origin)

def merged_cell_report(cell):
    """Return str summarizing position and size of merged *cell*."""
#    return (
#        'merged cell at row %d, col %d, %d cells high and %d cells wide'
#        % (cell.row_idx, cell.col_idx, cell.span_height, cell.span_width)
#    )
    return (
        'merged cell , cells high %d and %d cells wide'
        % (cell.span_height, cell.span_width)
    )

def iter_visible_cells(table):
    return (cell for cell in table.iter_cells() if not cell.is_spanned)

def unmerged_cell_report(cell):
    """Return str summarizing position and size of merged *cell*."""
    return (
        'cell : %s'
        % ( cell.text)
    )

prs = Presentation("sample.pptx")

result = []

col_idx=0
row_idx=0
for slide in prs.slides:
  for shape in slide.shapes:
    if shape.has_text_frame:
      for paragraph in shape.text_frame.paragraphs:
        result.append(paragraph.text)
    elif shape.has_table:
      print("table")
      # ---Print a summary line for each merged cell in *table*.---
      for merge_origin_cell in iter_merge_origins(shape.table):
          print(merged_cell_report(merge_origin_cell))
      for cell in shape.table.iter_cells():
        if not cell.is_spanned:
          print('cell[%d,%d] : %s' % (col_idx, row_idx, cell.text))
        col_idx += 1
        row_idx += 1

      # looping ...
      row_idx = 0
      for row in shape.table.rows:
        col_idx = 0
        for col in shape.table.columns:
          if not row.cells[col_idx].is_spanned:
            print("cell[%d:%d] - %s" % (col_idx, row_idx, row.cells[col_idx].text))
          col_idx += 1
        row_idx += 1

print(result)
