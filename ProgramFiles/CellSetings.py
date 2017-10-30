# File with cell settings

from openpyxl.styles import NamedStyle, Border, Side

cellStyle = NamedStyle(name="Typical")

bd = Side(border_style='thin', color='FF000000')

cellStyle.border = Border(left = bd, right=bd, top=bd, bottom=bd)
