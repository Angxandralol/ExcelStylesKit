# Documentation - API Reference
1. [Module `table`](#module-table)
    - [Cell](#cell)
    - [Table](#table)
2. [Module `excel`](#module-excel)
    - [Excel](#excel)
    - [ExcelStyles](#excelstyles)
--- 


## Module `table`
### Cell
#### *class* table.Cell(*column*, *row*)
Return a cell object based on the column and row.

##### Parameters
| Parameters | Type | Description             |
| ---------- | ---- | ----------------------- |
| column     | str  | The column of the cell. |
| row        | int  | The row of the cell.    |

##### Usage
```python
>>> from table import Cell
>>> cell = Cell('A', 1)
```

#### get_cell()
Return the cell in the format 'column' + 'row'.
**Return**: str

##### Parameters
None

##### Usage
```python
>>> from table import Cell
>>> cell = Cell('A', 1)
>>> cell.get_cell()
'A1'
```

#### get_column()
Return the column of the cell.
**Return**: str

##### Parameters
None

##### Usage
```python
>>> from table import Cell
>>> cell = Cell('A', 1)
>>> cell.get_column()
'A'
```

#### get_row()
Return the row of the cell.
**Return**: int

##### Parameters
None

##### Usage
```python
>>> from table import Cell
>>> cell = Cell('A', 1)
>>> cell.get_row()
1
```

### Table
#### *class* table.Table(*first_column*, *first_row*, *last_column*, *last_row*)
Return a table object based on the first and last column and row.

##### Parameters
| Parameters    | Type | Description                      |
| ------------- | ---- | -------------------------------- |
| first_column  | str  | The initial column of the table. |
| first_row     | int  | The initial row of the table.    |
| last_column   | str  | The final column of the table.   |
| last_row      | int  | The final row of the table.      |

##### Usage
|---|------|------|------|------|
|   |   A  |   B  |   C  |   D  | 
|---|------|------|------|------|
| 1 |  A1  |  B1  |  C1  |  D1  | 
|---|------|------|------|------|
| 2 |  A2  |  B2  |  C2  |  D2  |
|---|------|------|------|------|
| 3 |  A3  |  B3  |  C3  |  D3  |
|---|------|------|------|------|
| 4 |  A4  |  B4  |  C4  |  D4  |
|---|------|------|------|------|

In this table, the first cell is the cell 'A1' and the last cell is the cell 'D4'.

```python
>>> from table import Table
>>> table = Table('A', 1, 'D', 4)
```

#### get_cells()
Return all cells (header and content) of the table.
**Return**: List[table.Cell]

##### Parameters
None

##### Usage
```python
>>> from table import Table
>>> table = Table('A', 1, 'D', 4)
>>> table.get_cells()
[<excelstyleskit.table.cell object, ...>]
```

#### add_row_header()
Add the all cells of a row as the header of the table.

##### Parameters
| Parameters | Type | Description             |
| ---------- | ---- | ----------------------- |
| row        | int  | The row to add as header.|

##### Usage
```python
>>> from table import Table
>>> table = Table('A', 1, 'D', 4)
>>> table.add_row_header(1)
```

#### get_cells_header()
Return the header cells (without cells content) of the table.
**Return**: List[table.Cell]

##### Parameters
None

##### Usage
```python
>>> from table import Table
>>> table = Table('A', 1, 'D', 4)
>>> table.add_row_header(1)
>>> table.get_cells_header()
[<excelstyleskit.table.cell object, ...>]
```

#### get_cells_content()
Return the content cells (without cells header) of the table.
**Return**: List[table.Cell]

##### Parameters
None

##### Usage
```python
>>> from table import Table
>>> table = Table('A', 1, 'D', 4)
>>> table.add_row_header(1)
>>> table.get_cells_content()
[<excelstyleskit.table.cell object, ...>]
```

#### select_row()
Return the all cells of a row.
**Return**: List[table.Cell]

##### Parameters
| Parameters | Type | Description                 |
| ---------- | ---- | --------------------------- |
| row        | int  | The row to filter the cells.|

##### Usage
```python
>>> from table import Table
>>> table = Table('A', 1, 'D', 4)
>>> table.select_row(1)
[<excelstyleskit.table.cell object, ...>]
```

#### get_cells_str()
Return a list with the cells in the format 'column' + 'row'.
**Returns**: List[str]

##### Parameters
| Parameters | Type             | Description            |
| ---------- | ---------------- | ---------------------- |
| cells      | List[table.Cell] | The cells to convert.  |

##### Usage
```python
>>> from table import Table
>>> cells = [Cell('A', 1), Cell('B', 1), Cell('C', 1), Cell('D', 1)]
>>> Table.get_cells_str(cells)
['A1', 'B1', 'C1', 'D1']
```

#### view_cells()
Prints all the cells in the format 'column' + 'row'.
**Returns**: *str*. The string of the cells that were printed.

##### Parameters
| Parameters | Type             | Description            |
| ---------- | ---------------- | ---------------------- |
| cells      | List[table.Cell] | The cells to print.    |

##### Usage
```python
>>> from table import Table
>>> cells = [Cell('A', 1), Cell('B', 1), Cell('C', 1), Cell('D', 1)]
>>> Table.view_cells(cells)
| A1 | B1 | C1 | D1 | 
```

## Module `excel`
### Excel
#### *class* excel.Excel(*filepath*, *sheetname*)
Return an Excel object based on the filepath and sheetname.

##### Parameters
| Parameters | Type                 | Description                                                                                           |
| ---------- | -------------------- | ----------------------------------------------------------------------------------------------------- |
| filepath   | str                  | The filepath of the Excel spreadsheet.                                                                |
| sheetname  | str o None, optional | The name of the sheet. If it is not specified, the first sheet of the Excel spreadsheet will be used. |

##### Usage 
```python
>>> from excel import Excel
>>> excel_sheet = Excel('path/to/file.xlsx', 'Sheet1')
>>> excel = Excel('path/to/file.xlsx')
```

#### set_table(*first_column*, *first_row*, *last_column*, *last_row*)
Set the table of the excel.

##### Parameters
| Parameters    | Type | Description                      |
| ------------- | ---- | -------------------------------- |
| first_column  | str  | The initial column of the table. |
| first_row     | int  | The initial row of the table.    |
| last_column   | str  | The final column of the table.   |
| last_row      | int  | The final row of the table.      |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_table('A', 1, 'D', 4)
```

#### add_row_header(*row*)
Add the all cells of a row as the header of the table.

##### Parameters
| Parameters | Type | Description             |
| ---------- | ---- | ----------------------- |
| row        | int  | The row to add as header.|

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_table('A', 1, 'D', 4)
>>> excel.add_row_header(1)
```

#### get_workbook()
Return the workbook of the excel.
**Returns**: *Workbook*

##### Parameters
None

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.get_workbook()
<openpyxl.workbook.workbook.Workbook object>
```

#### get_worksheet()
Return the worksheet of the excel.
**Returns**: *Worksheet*

##### Parameters
None

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.get_worksheet()
<openpyxl.worksheet.worksheet.Worksheet object>
```

#### get_table()
Return the object table of the excel.
**Returns**: *table.Table*

##### Parameters
None

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.get_table()
<excelstyleskit.table.Table object>
```

#### save_work()
Save the changes in the excel.  

##### Parameters
None

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.save_work()
```

#### set_background_color_header(*start_color*, *end_color*)
Set the background color of the header of the table.

##### Parameters
| Parameters  | Type | Description                            |
| ----------  | ---- | -------------------------------------- |
| start_color | str  | The start color in hexadecimal format. |
| end_color   | str  | The end color in hexadecimal format.   |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_background_color_header('FFFFFF', 'FFFFFF')
```

#### set_background_color_content(*start_color*, *end_color*)
Set the background color of the content of the table.

##### Parameters
| Parameters  | Type | Description                            |
| ----------  | ---- | -------------------------------------- |
| start_color | str  | The start color in hexadecimal format. |
| end_color   | str  | The end color in hexadecimal format.   |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_background_color_content('FFFFFF', 'FFFFFF')
```

#### get_background_color_cell(*cell*)
Return the background color of the cell.
**Return**: str

##### Parameters
| Parameters | Type | Description                           |
| ---------- | ---- | ------------------------------------  |
| cell       | Cell | The cell to get the background color. |

##### Usage
```python
>>> from excel import Excel, Cell
>>> excel = Excel('path/to/file.xlsx')
>>> excel.get_background_color_cell(Cell('A', 1))
'FFFFFF'
```

#### set_font_header(*name*, *size*, *bold*, *italic*, *vertAlign*, *underline*, *strike*, *color*)
Set the font of the header of the table.

##### Parameters
| Parameters | Type | Description                                                                                    |
| ---------- | ---- | ---------------------------------------------------------------------------------------------- |
| name       | str  | Name of the font.                                                                              |
| size       | int  | Size of the font.                                                                              |
| bold       | bool | True if the font is bold.                                                                      |
| italic     | bool | True if the font is italic.                                                                    |
| vertAlign  | str  | Value must be one of {‘superscript’, ‘baseline’, ‘subscript’} to the font.                     |
| underline  | str  | Value must be one of {‘single’, ‘double’, ‘doubleAccounting’, ‘singleAccounting’} to the font. |
| strike     | bool | True if the font is strike.                                                                    |
| color      | str  | The color of the font in hexadecimal format.                                                   |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_font_header('Arial', 10, True, False, 'superscript', 'single', True, 'FFFF0000')
```

#### set_font_content(*name*, *size*, *bold*, *italic*, *vertAlign*, *underline*, *strike*, *color*)
Set the font of the content of the table.

##### Parameters
| Parameters | Type | Description                                                                                    |
| ---------- | ---- | ---------------------------------------------------------------------------------------------- |
| name       | str  | Name of the font.                                                                              |
| size       | int  | Size of the font.                                                                              |
| bold       | bool | True if the font is bold.                                                                      |
| italic     | bool | True if the font is italic.                                                                    |
| vertAlign  | str  | Value must be one of {‘superscript’, ‘baseline’, ‘subscript’} to the font.                     |
| underline  | str  | Value must be one of {‘single’, ‘double’, ‘doubleAccounting’, ‘singleAccounting’} to the font. |
| strike     | bool | True if the font is strike.                                                                    |
| color      | str  | The color of the font in hexadecimal format.                                                   |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_font_content('Arial', 10, True, False, 'superscript', 'single', True, 'FFFF0000')
```

#### get_font_cell(*cell*)
Return the font of the cell.
**Return**: Font

##### Parameters
| Parameters | Type | Description                           |
| ---------- | ---- | ------------------------------------  |
| cell       | Cell | The cell to get the font.             |

##### Usage
```python
>>> from excel import Excel, Cell
>>> excel = Excel('path/to/file.xlsx')
>>> excel.get_font_cell(Cell('A', 1))
<openpyxl.styles.font.Font object>
```

#### set_alignment_header(*horizontal*, *vertical*, *text_rotation*, *wrap_text*, *shrink*, *indent*)
Set the alignment of the header of the table.

##### Parameters
| Parameters         | Type | Description                                                                                               |
| ------------------ | ---- | --------------------------------------------------------------------------------------------------------- |
| horizontal         | str  | Value must be one of {‘left’, ‘center’, ‘right’, ‘fill’, ‘justify’, ‘centerContinuous’, ‘distributed’} to the horizontal. |
| vertical           | str  | Value must be one of {‘top’, ‘center’, ‘bottom’, ‘justify’, ‘distributed’} to the vertical.               |
| text_rotation      | int  | The rotation of the text.                                                                                 |
| wrap_text          | bool | True if the text is wrapped.                                                                              |
| shrink_to_fit      | bool | True if the text is shrink.                                                                               |
| indent             | int  | The indent of the text.                                                                                   |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_alignment_header('left', 'center', 0, True, True, 1)
```

#### set_alignment_content(*horizontal*, *vertical*, *text_rotation*, *wrap_text*, *shrink*, *indent*)
Set the alignment of the content of the table.

##### Parameters
| Parameters         | Type | Description                                                                                               |
| ------------------ | ---- | --------------------------------------------------------------------------------------------------------- |
| horizontal         | str  | Value must be one of {‘left’, ‘center’, ‘right’, ‘fill’, ‘justify’, ‘centerContinuous’, ‘distributed’} to the horizontal. |
| vertical           | str  | Value must be one of {‘top’, ‘center’, ‘bottom’, ‘justify’, ‘distributed’} to the vertical.               |
| text_rotation      | int  | The rotation of the text.                                                                                 |
| wrap_text          | bool | True if the text is wrapped.                                                                              |
| shrink_to_fit      | bool | True if the text is shrink.                                                                               |
| indent             | int  | The indent of the text.                                                                                   |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_alignment_content('left', 'center', 0, True, True, 1)
```

#### get_alignment_cell(*cell*)
Return the alignment of the cell.
**Return**: Alignment

##### Parameters
| Parameters | Type | Description                           |
| ---------- | ---- | ------------------------------------  |
| cell       | Cell | The cell to get the alignment.        |

##### Usage
```python
>>> from excel import Excel, Cell
>>> excel = Excel('path/to/file.xlsx')
>>> excel.get_alignment_cell(Cell('A', 1))
<openpyxl.styles.alignment.Alignment object>
```

#### set_border_header(*color*, *style*)
Set the border of the header of the table.

##### Parameters
| Parameters | Type | Description                                                                                               |
| ---------- | ---- | --------------------------------------------------------------------------------------------------------- |
| color      | str  | The color of the border in hexadecimal format.                                                            |
| style      | str  | Value must be one of {'dashDot','dashDotDot', 'dashed','dotted','double','hair', 'medium', 'mediumDashDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'} to the style. |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_border_header('000000', 'thin')
```

#### set_border_content(*color*, *style*)
Set the border of the content of the table.

##### Parameters
| Parameters | Type | Description                                                                                               |
| ---------- | ---- | --------------------------------------------------------------------------------------------------------- |
| color      | str  | The color of the border in hexadecimal format.                                                            |
| style      | str  | Value must be one of {'dashDot','dashDotDot', 'dashed','dotted','double','hair', 'medium', 'mediumDashDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'} to the style. |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_border_content('000000', 'thin')
```

#### get_border_cell(*cell*)
Return the border of the cell.
**Return**: Border

##### Parameters
| Parameters | Type | Description                  |
| ---------- | ---- | ---------------------------- |
| cell       | Cell | The cell to get the border.  |

##### Usage
```python
>>> from excel import Excel, Cell
>>> excel = Excel('path/to/file.xlsx')
>>> excel.get_border_cell(Cell('A', 1))
<openpyxl.styles.border.Border object>
```

#### set_height_row_header(*height*)
Set the height of the row header of the table.

##### Parameters
| Parameters | Type | Description             |
| ---------- | ---- | ----------------------- |
| height     | int  | The height of the row.  |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_height_row_header(12.75)
```

#### set_height_row_content(*height*)
Set the height of the row content of the table.

##### Parameters
| Parameters | Type | Description             |
| ---------- | ---- | ----------------------- |
| height     | int  | The height of the row.  |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_height_row_content(12.75)
```

#### get_height_row_cell(*cell*)
Return the height of the row of a cell.
**Return**: int

##### Parameters
| Parameters | Type | Description                  |
| ---------- | ---- | ---------------------------- |
| cell       | Cell | The cell to get the height.  |

##### Usage
```python
>>> from excel import Excel, Cell
>>> excel = Excel('path/to/file.xlsx')
>>> excel.get_height_row_cell(Cell('A', 1))
12.75
```

#### set_width_column_table(*width*)
Set the width of the column of the table.

##### Parameters
| Parameters | Type | Description              |
| ---------- | ---- | ------------------------ |
| width      | int  | The width of the column. |

##### Usage
```python
>>> from excel import Excel
>>> excel = Excel('path/to/file.xlsx')
>>> excel.set_width_column_table(15)
```

#### get_width_column_cell(*cell*)
Return the width of the column of a cell.
**Return**: int

##### Parameters
| Parameters | Type | Description                  |
| ---------- | ---- | ---------------------------- |
| cell       | Cell | The cell to get the width.   |

##### Usage
```python
>>> from excel import Excel, Cell
>>> excel = Excel('path/to/file.xlsx')
>>> excel.get_width_column_cell(Cell('A', 1))
11.53
```
### ExcelStyles
#### ExcelStyles.set_background_color(*worksheet*, *cells*, *fill*)
Set the background color of the cells.

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cells      | List[table.Cell]                       | The cells to set the background color.       |
| fill       | openpyxl.styles.PatternFill            | The fill of the cells.                       |

##### Usage
The use of this method must be through the use of the method: 
- [`set_background_color_header(*start_color*, *end_color*)`](#set_background_color_headerstart_color-end_color)
- [`set_background_color_content(*start_color*, *end_color*)`](#set_background_color_content-start_color-end_color)

#### ExcelStyles.get_background_color(*worksheet*, *cell*)
Return the background color of the cell.
**Return**: str

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cell       | table.Cell                             | The cell to get the background color.        |

##### Usage
The use of this method must be through the use of the method: 
- [`get_background_color_cell(*cell*)`](#get_background_color_cellcell)

#### ExcelStyles.set_font(*worksheet*, *cells*, *font*)
Set the font of the cells.

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cells      | List[table.Cell]                       | The cells to set the font.                   |
| font       | openpyxl.styles.Font                   | The font of the cells.                       |

##### Usage
The use of this method must be through the use of the method: 
- [`set_font_header(*name*, *size*, *bold*, *italic*, *vertAlign*, *underline*, *strike*, *color*)`](#set_font_headername-size-bold-italic-vertalign-underline-strike-color)
- [`set_font_content(*name*, *size*, *bold*, *italic*, *vertAlign*, *underline*, *strike*, *color*)`](#set_font_contentname-size-bold-italic-vertalign-underline-strike-color)

#### ExcelStyles.get_font(*worksheet*, *cell*)
Return the font of the cell.
**Return**: Font

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cell       | table.Cell                             | The cell to get the font.                    |

##### Usage
The use of this method must be through the use of the method: 
- [`get_font_cell(*cell*)`](#get_font_cellcell)

#### ExcelStyles.set_alignment(*worksheet*, *cells*, *alignment*)
Set the alignment of the cells.

##### Parameters
| Parameters         | Type                                   | Description                                  |
| ------------------ | -------------------------------------- | -------------------------------------------- |
| worksheet          | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cells              | List[table.Cell]                       | The cell to set the alignment.               |
| alignment          | openpyxl.styles.Alignment              | The alignment of the cells.                  |

##### Usage
The use of this method must be through the use of the method: 
- [`set_alignment_header(*horizontal*, *vertical*, *text_rotation*, *wrap_text*, *shrink*, *indent*)`](#set_alignment_headerhorizontal-vertical-text_rotation-wrap_text-shrink-indent)
- [`set_alignment_content(*horizontal*, *vertical*, *text_rotation*, *wrap_text*, *shrink*, *indent*)`](#set_alignment_contenthorizontal-vertical-text_rotation-wrap_text-shrink-indent)

#### ExcelStyles.get_alignment(*worksheet*, *cell*)
Return the alignment of the cell.
**Return**: Alignment

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cell       | table.Cell                             | The cell to get the alignment.               |

##### Usage
The use of this method must be through the use of the method: 
- [`get_alignment_cell(*cell*)`](#get_alignment_cellcell)

#### ExcelStyles.set_border(*worksheet*, *cells*, *border*)
Set the border of the cells.

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cells      | List[table.Cell]                       | The cells to set the border.                 |
| border     | openpyxl.styles.Border                 | The border to set.                           |

##### Usage
The use of this method must be through the use of the method: 
- [`set_border_header(*color*, *style*)`](#set_border_headercolor-style)
- [`set_border_content(*color*, *style*)`](#set_border_contentcolor-style)

#### ExcelStyles.get_border(*worksheet*, *cell*)
Return the border of the cell.
**Return**: Border

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cell       | table.Cell                             | The cell to get the border.                  |

##### Usage
The use of this method must be through the use of the method: 
- [`get_border_cell(*cell*)`](#get_border_cellcell)

#### ExcelStyles.set_height(*worksheet*, *cells*, *height*)
Set the height of the cells.

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cells      | List[table.Cell]                       | The cell to set the height.                  |
| height     | int                                    | The height of the cell.                      |

##### Usage
The use of this method must be through the use of the method: 
- [`set_height_row_header(*height*)`](#set_height_row_headerheight)
- [`set_height_row_content(*height*)`](#set_height_row_contentheight)

#### ExcelStyles.get_height(*worksheet*, *cell*)
Return the height of the cell.
**Return**: int

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cell       | table.Cell                             | The cell to get the height.                  |

##### Usage
The use of this method must be through the use of the method: 
- [`get_height_row_cell(*cell*)`](#get_height_row_cellcell)

#### ExcelStyles.set_width(*worksheet*, *cells*, *width*)
Set the width of the cells.

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cells      | List[table.Cell]                       | The cell to set the width.                   |
| width      | int                                    | The width of the cell.                       |

##### Usage
The use of this method must be through the use of the method: 
- [`set_width_column_table(*width*)`](#set_width_column_tablewidth)

#### ExcelStyles.get_width(*worksheet*, *cell*)
Return the width of the cell.
**Return**: int

##### Parameters
| Parameters | Type                                   | Description                                  |
| ---------- | -------------------------------------- | -------------------------------------------- |
| worksheet  | openpyxl.worksheet.worksheet.Worksheet | The worksheet of the excel.                  |
| cell       | table.Cell                             | The cell to get the width.                   |

##### Usage
The use of this method must be through the use of the method: 
- [`get_width_column_cell(*cell*)`](#get_width_column_cellcell) 