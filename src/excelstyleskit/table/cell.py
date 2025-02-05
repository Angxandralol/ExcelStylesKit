from utils import is_valid_column, is_valid_row

class Cell:
    _column: str
    _row: int

    def __init__(self, column: str, row: int):
        if not is_valid_column(column):
            raise ValueError('Invalid column')
        if not is_valid_row(row):
            raise ValueError('Invalid row')
        self._column = column
        self._row = row

    def get_cell(self) -> str:
        """Returns the cell in the format 'column' + 'row'.

        Example
        -------
        >>> cell = Cell('A', 1)
        >>> cell.get_cell()
        'A1'
        """
        return f'{self._column}{self._row}'

    def get_column(self) -> str:
        """Returns the column of the cell.

        Example
        -------
        >>> cell = Cell('A', 1)
        >>> cell.get_column()
        'A'
        """
        return self._column
    
    def get_row(self) -> int:
        """Returns the row of the cell.

        Example
        -------
        >>> cell = Cell('A', 1)
        >>> cell.get_row()
        1
        """
        return self._row
