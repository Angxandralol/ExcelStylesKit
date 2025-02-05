import os

def is_valid_column(column: str) -> bool:
    """Returns True if the column is a letter and has only one character. 
    
    Parameters
    ----------
    column : str
        The column to verify.
    """
    return column.isalpha() and len(column) == 1

def is_valid_row(row: int) -> bool:
    """Returns True if the row is a positive integer and less than 1000000000.
    
    Parameters
    ----------
    row : int
        The row to verify.
    """
    return type(row) == int and row > 0 and row < 1000000000

def is_valid_filepath(filepath: str) -> bool:
    """Returns True if the file exists.
    
    Parameters
    ----------
    filepath : str
        The file to verify.
    """
    return os.path.exists(filepath)