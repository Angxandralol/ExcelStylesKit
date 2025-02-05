import unittest
from table import Cell

class TestCell(unittest.TestCase):

    def test_cell(self):
        cell = Cell('A', 1)
        self.assertEqual(cell.get_cell(), 'A1')

    def test_row(self):
        cell = Cell('A', 1)
        self.assertEqual(cell.get_row(), 1)

    def test_column(self):
        cell = Cell('A', 1)
        self.assertEqual(cell.get_column(), 'A')

if __name__ == '__main__':  
    unittest.main()