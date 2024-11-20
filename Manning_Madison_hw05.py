from openpyxl.styles import Color, PatternFill
art_book = openpyxl.Workboo()
sheet = art_book.active

#Setting cell sizes
sheet.row_dimensions[1:55].height = 8.5
sheet.column_dimensions[1:55].width = 1
art_book.save("dimensions.xlsx")

#Setting cell colors
black = Color(rgb= '000000')
fill_blk = PatternFill(patternType = 'solid', fgColor=black)

yellow = Color(rgb = 'ffc000')
fill_yel = PatternFill(patternType = 'solid', fgColor=yellow)

pale_yellow = Color(rgb= 'ffe669')
fill_pyel = PatternFill(patternType = 'solid', fgColor=pale_yellow)

red = Color(rgb = 'c00000')
fill_red = PatternFill(patternType = 'solid', fgColor=red)

blue = Color(rgb= '8ea9db')
fill_bl = PatternFill(patternType = 'solid', fgColor=blue)
