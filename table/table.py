from typing import List
from utils import is_valid_column, is_valid_row
from table.cell import Cell

class Table:
    _cells_content: List[Cell]
    _cells_header: List[Cell]

    def __init__(self, first_column: str, first_row: int, last_column: str, last_row: int):
        if not is_valid_column(first_column) or not is_valid_column(last_column):
            raise ValueError('Invalid column')
        if not is_valid_row(first_row) or not is_valid_row(last_row):
            raise ValueError('Invalid row')
        self._cells_content = []
        self._cells_header = []
        for row in range(first_row, last_row + 1):
            for column in range(ord(first_column), ord(last_column) + 1):
                self._cells_content.append(Cell(chr(column).upper(), row))

    @staticmethod
    def get_cells_str(cells: List[Cell]) -> List[str]:
        """Returns a list with the cells in the format 'column' + 'row'. 
        
        Parameters
        ----------
        cells : List[Cell]
            The cells to convert.
        """
        return [cell.get_cell() for cell in cells]
    
    @staticmethod
    def view_cells(cells: List[Cell]) -> str:
        """Prints all the cells in the format 'column' + 'row'.
        
        Parameters
        ----------
        cells : List[Cell]
            The cells to print.

        Returns
        -------
            The string of the cells that were printed.

        Example
        -------
        >>> table = Table('A', 1, 'C', 2)
        >>> Table.view_cells(table.get_cells_content())
        | A1 | B1 | C1 | A2 | B2 | C2 |
        """
        list: str = "| "
        for cell in cells:
            list += cell.get_cell() + " | " 
        print(list)
        return list
        
    def get_cells_content(self) -> List[Cell]:
        """Returns the content cells (without cells header) of the table. """
        return self._cells_content
    
    def get_cells_header(self) -> List[Cell]:
        """Returns the header cells (without cells content) of the table. """
        return self._cells_header
    
    def get_cells(self) -> List[Cell]:
        """Returns all cells (header and content) of the table. """
        return self._cells_header + self._cells_content
    
    def add_row_header(self, row: int) -> None:
        """Add the all cells of a row as the header of the table. 
        
        Parameters
        ----------
        row : int
            The row to add as header.
        """
        cells = self.select_row(row)
        [self._cells_header.append(cell) for cell in cells if not cell in self._cells_header]
        [self._cells_content.remove(cell_header) for cell_header in self._cells_header if cell_header in self._cells_content]

    def select_row(self, row: int) -> List[Cell]:
        """Returns the all cells of the row. 
        
        Parameters
        ----------
        row : int
            The row to filter the cells.
        """
        cells = self._cells_header + self._cells_content
        return [cell for cell in cells if cell.get_row() == row]
