from modules.tableConfig import TableConfig

color_uno = '8C001A'
color_dos = '5F021F'

table_config = TableConfig('example.xlsx')
table_config.header.set_header(start_row=1, start_column='A', end_row=2, end_column='E')
table_config.content.set_content(start_row=2, start_column='A', end_row=10, end_column='E')

# table_config.content.style_font(color=color_uno)
# table_config.style_border()
# table_config.style_column_width(80)
# table_config.content.style_height_row(90)