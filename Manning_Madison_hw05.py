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
black = 'FF010101'
fill_blk = PatternFill(patternType = 'solid', fgColor= black)
black_list=['R1', 'S1', 'T1', 'U1', 'V1', 'Q2', 'W2', 'Q3', 'W3', 'X3', 'Q4', 'X4', 'Q5', 'X5', 'O6', 'P6', 'W6', 'X6', 
'M7', 'N7', 'Y7', 'Z7', 'K8', 'L8', 'AA8', 'AB8', 'J9', 'AB9', 'AC9', 'I10', 'AD10', 'H11', 'W11', 'X11', 'Y11', 'AE11', 
'H12', 'AF12', 'AG12', 'AH12', 'G13', 'AI13', 'AJ13', 'G14', 'AJ14', 'AK14', 'Z15', 'AA15', 'AK15', 'A16', 'Z16', 'AA16', 
'AB16', 'AK16', 'AL16', 'AM16', 'AN16', 'AO16', 'AP16', 'AQ16', 'A17', 'AA17', 'AB17', 'AH17', 'AI17', 'AJ17', 'AR17', 'AS17', 
'AT17', 'AU17', 'A18', 'AG18', 'AH18', 'AK18', 'AQ18', 'AV18', 'A19', 'AF19', 'AJ19', 'AR19', 'AV19', 'B20', 'F20', 'K20', 
'AF20', 'AK20', 'AL20', 'AM20', 'AN20', 'AO20', 'AP20', 'AQ20', 'AV20', 'C21', 'D21', 'E21', 'F21', 'K21', 'AF21', 'AG21', 
'AT21', 'AU21', 'F22', 'K22', 'Y22', 'AG22', 'AH22', 'AJ22', 'AM22', 'AN22', 'AQ22', 'AR22', 'AS22', 'AU22', 'G23', 'X23', 
'Y23', 'Z23', 'AG23', 'AI23', 'AK23', 'AM23', 'AN23', 'AQ23', 'AV23', 'G24', 'N24', 'O24', 'W24', 'X24', 'Y24', 'Z24', 'AF24', 
'AK24', 'AM24', 'AO24', 'AP24', 'AV24', 'H25', 'N25', 'O25', 'P25', 'X25', 'Y25', 'Z25', 'AD25', 'AF25', 'AK25', 'AM25', 'AV25', 
'H26', 'O26', 'P26', 'AD26', 'AF26', 'AK26', 'AL26', 'AW26', 'I27', 'AD27', 'AF27', 'AW27', 'AX27', 'J28', 'AC28', 'AF28', 'AH28', 
'AI28', 'AJ28', 'AK28', 'AL28', 'AW28', 'AX28', 'K29', 'Y29', 'Z29', 'AA29', 'AB29', 'AE29', 'AF29', 'AG29', 'AM29', 'AV29', 
'AY29', 'K30', 'AD30', 'AH30', 'AN30', 'AU30', 'AY30', 'K31', 'AB31', 'AC31', 'AH31', 'AN31', 'AT31', 'AY31', 'K32', 'L32', 
'AA32', 'AI32', 'AN32', 'AR32', 'AS32', 'AX32', 'L33', 'Z33', 'AI33', 'AM33', 'AN33', 'AO33', 'AP33', 'AQ33', 'AR33', 'AS33', 
'AW33', 'M34', 'Y34', 'AI34', 'AM34', 'AN34', 'AS34', 'AT34', 'AU34', 'AV34', 'N35', 'Y35', 'AJ35', 'AK35', 'AL35', 'AO35', 
'AP35', 'AQ35', 'AR35', 'O36', 'P36', 'W36', 'X36', 'Y36', 'AI36', 'AJ36', 'AP36', 'P37', 'Q37', 'R37', 'S37', 'T37', 'U37', 
'V37', 'Y37', 'AH37', 'AI37', 'AP37', 'V38', 'AE38', 'AF38', 'AG38', 'AM38', 'AN38', 'AO38', 'AP38', 'V39', 'AH39', 'AI39', 
'AJ39', 'AK39', 'AL39', 'AP39', 'V40', 'AE40', 'AF40', 'AG40', 'AP40', 'V41', 'AB41', 'AC41', 'AD41', 'AP41', 'V42', 'Z42', 
'AA42', 'AP42', 'V43', 'W43', 'X43', 'Y43', 'AP43', 'V44', 'AO44', 'V45', 'AO45', 'V46', 'AO46', 'W47', 'AD47', 'AE47', 'AF47', 
'AG47', 'AH47', 'AO47', 'W48', 'AD48', 'AH48', 'AN48', 'AO48', 'W49', 'AD49', 'AH49', 'AN49', 'W50', 'AD50', 'AG50', 'AN50', 
'X51', 'AE51', 'AG51', 'AN51', 'X52', 'AE52', 'AG52', 'AM52', 'Y53', 'Z53', 'AA53', 'AB53', 'AC53', 'AD53', 'AE53', 'AG53', 
'AH53', 'AI53', 'AJ53', 'AK53', 'AL53', 'AM53']

yellow = 'FFFFC000'
fill_yel = PatternFill(patternType = 'solid', fgColor=yellow)
yellow_list=[]

pale_yellow = 'FFFFE669'
fill_pyel = PatternFill(patternType = 'solid', fgColor=pale_yellow)
pale_yellow_list = []

red = 'FFC00000'
fill_red = PatternFill(patternType = 'solid', fgColor=red)
red_list = []

blue = 'FF8EA9DB'
fill_bl = PatternFill(patternType = 'solid', fgColor=blue)
blue_list = []

#scraper
def remove_alpha(rgb):
    return rgb[2:8]

wb = openpyxl.load_workbook("COSC1010_pixel_art.xlsx")
sheet = wb.active
cells = sheet['A1:BA55']
color_dic = {'blk': [], 'yel' : [], 'p_yel' : [], 'r' : [], 'bl' : []}

for row in sheet['A1:BA55']:
    for cell in row:
        if cell.fill and cell.fill.fgColor:
            color = cell.fill.fgColor.rgb
            if color == black:
                color_dic['blk'].append(cell.coordinate)
            elif color == yellow:
                color_dic['yel'].append(cell.coordinate)
            elif color == pale_yellow:
                color_dic['p_yel'].append(cell.coordinate)
            elif color == red:
                color_dic['r'].append(cell.coordinate)
            elif color == blue:
                color_dic['bl'].append(cell.coordinate)


print(f"Yellow:", color_dic['p_yel'])
        

