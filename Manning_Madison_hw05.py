import openpyxl
from openpyxl.styles import Color, PatternFill
from openpyxl.utils import get_column_letter

art_book = openpyxl.Workbook()
sheet = art_book.active

#Setting cell sizes
for row in range(1, 55):
    sheet.row_dimensions[row].height = 8.5
for column in range(1,53):
    column_letter = get_column_letter(column)
    sheet.column_dimensions[column_letter].width = 1
art_book.save("dimensions.xlsx")

#Setting cell colors
black = '000000'
fill_blk = PatternFill(patternType = 'solid', fgColor= black)

yellow = 'ffc000'
fill_yel = PatternFill(patternType = 'solid', fgColor=yellow)

pale_yellow = 'ffe669'
fill_pyel = PatternFill(patternType = 'solid', fgColor=pale_yellow)

red = 'c00000'
fill_red = PatternFill(patternType = 'solid', fgColor=red)

blue = '8ea9db'
fill_bl = PatternFill(patternType = 'solid', fgColor=blue)

#scraper
wb = openpyxl.load_workbook("COSC1010_pixel_art.xlsx")
sheet = wb.active
cells = tuple(sheet['A1:BA55'])
color_dic = {'blk': [], 'yel' : [], 'p_yel' : [], 'r' : [], 'bl' : []}

for row in sheet['A1:BA55']:
    for cell in row:
        column_letter = get_column_letter(column)
        if cell.fill.patternType:
            color = cell.fill.fgColor.rgb
            if color == black:
                color_dic['blk'].append(cell.coordinate)
            if color == yellow:
                color_dic['yel'].append(cell.coordinate)
            if color == pale_yellow:
                color_dic['p_yel'].append(cell.coordinate)
            if color == red:
                color_dic['r'].append(cell.coordinate)
            if color == blue:
                color_dic['bl'].append(cell.coordinate)

print(f"Black:", color_dic['blk'])
        

