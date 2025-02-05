import unittest
from table import Table

""" EXAMPLE TABLE
This is the Excel spreadsheet template used for all tests.

|---|------|------|------|------|
|   |   A  |   B  |   C  |   D  | <=== This is the row header of the table.
|---|------|------|------|------|
| 1 |  A1  |  B1  |  C1  |  D1  |  All other cells are content cells.
|---|------|------|------|------|
| 2 |  A2  |  B2  |  C2  |  D2  |
|---|------|------|------|------|
| 3 |  A3  |  B3  |  C3  |  D3  |
|---|------|------|------|------|
| 4 |  A4  |  B4  |  C4  |  D4  |
|---|------|------|------|------|


The first cell in the table is cell 'A1' and the last cell is cell 'D4'.

The view of all these cells is:
| A1 | B1 | C1 | D1 | A2 | B2 | C2 | D2 | A3 | B3 | C3 | D3 | A4 | B4 | C4 | D4 | 
"""

VIEW_CELLS = '| A1 | B1 | C1 | D1 | A2 | B2 | C2 | D2 | A3 | B3 | C3 | D3 | A4 | B4 | C4 | D4 | '
VIEW_CELLS_FIRST_HEADER = '| A1 | B1 | C1 | D1 | '
VIEW_CELLS_SECOND_HEADER = '| A1 | B1 | C1 | D1 | A2 | B2 | C2 | D2 | ' 
VIEW_CELLS_FIRST_CONTENT = '| A2 | B2 | C2 | D2 | A3 | B3 | C3 | D3 | A4 | B4 | C4 | D4 | '
VIEW_CELLS_SECOND_CONTENT = '| A3 | B3 | C3 | D3 | A4 | B4 | C4 | D4 | '
LIST_CELLS = ['A1', 'B1', 'C1', 'D1', 'A2', 'B2', 'C2', 'D2', 'A3', 'B3', 'C3', 'D3', 'A4', 'B4', 'C4', 'D4']
LIST_CELLS_FIRST_HEADER = ['A1', 'B1', 'C1', 'D1']
LIST_CELLS_SECOND_HEADER = ['A1', 'B1', 'C1', 'D1', 'A2', 'B2', 'C2', 'D2']
LIST_CELLS_FIRST_CONTENT = ['A2', 'B2', 'C2', 'D2', 'A3', 'B3', 'C3', 'D3', 'A4', 'B4', 'C4', 'D4']
LIST_CELLS_SECOND_CONTENT = ['A3', 'B3', 'C3', 'D3', 'A4', 'B4', 'C4', 'D4']

class TestTable(unittest.TestCase):
    FIRST_COLUMN = 'A'
    LAST_COLUMN = 'D'
    FIRST_ROW = 1
    SECOND_ROW = 2
    LAST_ROW = 4

    def create_table(self) -> Table:
        return Table(
            self.FIRST_COLUMN, 
            self.FIRST_ROW, 
            self.LAST_COLUMN, 
            self.LAST_ROW
        )
    
    def test_new_table(self):
        table = self.create_table()
        self.assertEqual(
            Table.get_cells_str(table.get_cells()), 
            LIST_CELLS
        )

    def test_add_header(self):
        table = self.create_table()
        table.add_row_header(self.FIRST_ROW)
        self.assertEqual(
            Table.get_cells_str(table.get_cells_header()), 
            LIST_CELLS_FIRST_HEADER
        )
        self.assertEqual(
            Table.get_cells_str(table.get_cells_content()), 
            LIST_CELLS_FIRST_CONTENT
        )

        table.add_row_header(self.SECOND_ROW)
        self.assertEqual(
            Table.get_cells_str(table.get_cells_header()), 
            LIST_CELLS_SECOND_HEADER
        )
        self.assertEqual(
            Table.get_cells_str(table.get_cells_content()), 
            LIST_CELLS_SECOND_CONTENT
        )
        

    def test_select_row(self):
        table = self.create_table()
        row_select = table.select_row(self.FIRST_ROW)
        self.assertEqual(
            Table.get_cells_str(row_select), 
            LIST_CELLS_FIRST_HEADER
        )

        table.add_row_header(self.FIRST_ROW)
        row_select = table.select_row(self.FIRST_ROW)
        self.assertEqual(
            Table.get_cells_str(row_select), 
            LIST_CELLS_FIRST_HEADER
        )

    def test_view_cells(self):
        table = self.create_table()

        table.add_row_header(self.FIRST_ROW)
        table_cells = Table.view_cells(table.get_cells())
        header_cells = Table.view_cells(table.get_cells_header())
        content_cells = Table.view_cells(table.get_cells_content())
        self.assertEqual(table_cells, VIEW_CELLS)
        self.assertEqual(header_cells, VIEW_CELLS_FIRST_HEADER)
        self.assertEqual(content_cells, VIEW_CELLS_FIRST_CONTENT)

        table.add_row_header(self.SECOND_ROW)
        table_cells = Table.view_cells(table.get_cells())
        header_cells = Table.view_cells(table.get_cells_header())
        content_cells = Table.view_cells(table.get_cells_content())
        self.assertEqual(table_cells, VIEW_CELLS)
        self.assertEqual(header_cells, VIEW_CELLS_SECOND_HEADER)
        self.assertEqual(content_cells, VIEW_CELLS_SECOND_CONTENT)

if __name__ == '__main__':
    unittest.main()