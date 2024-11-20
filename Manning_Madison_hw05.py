import openpyxl
from openpyxl.styles import Color, PatternFill

art_book = openpyxl.Workbook()
sheet = art_book.active

#Setting cell sizes
sheet.row_dimensions[1:55].height = 8.5
sheet.column_dimensions['A':'BA'].width = 1
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

#scraper
wb = openpyxl.load_workbook("COSC1010_pixel_at.xlsx")
sheet = wb.active
cells = tuple(sheet['A1':'BA55'])
color_dic = {'black': [], 'yellow' : [], 'pale_yellow' : [], 'red' : [], 'blue' : []}

for cell in cells:
    if cell.fill.patternType != None:
        if cell.fill.fgColor == black:
            color_dic['black'].append(cell.coordinate)
        if cell.fill.fgColor == yellow:
            color_dic['yellow'].append(cell.coordinate)
        if cell.fill.fgColor == pale_yellow:
            color_dic['pale_yellow'].append(cell.coordinate)
        if cell.fill.fgColor == red:
            color_dic['red'].append(cell.coordinate)
        if cell.fill.fgColor == blue:
            color_dic['blue'].append(cell.coordinate)

print(f"Yellow:")
print(color_dic[yellow])
        

