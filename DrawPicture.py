print('first test for Draw In Excel')
import time
from PIL import Image
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font,colors


im = Image.open("EVA_1.0.jpg", "r")
print(im.size,im.format,im.mode)
#im.show()
start = time.time()
#1080*1080
length = 240
width = 135
im_resize = im.resize((length,width))


wb = Workbook()
ws = wb.active
ws.title = "Draw in excel"


ABC = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


def color(value):
  digit = list(map(str, range(10))) + list("ABCDEF")
  if isinstance(value, tuple):
    string = ''
    for i in value:
      a1 = i // 16
      a2 = i % 16
      string += digit[a1] + digit[a2]
    return string
  elif isinstance(value, str):
    a1 = digit.index(value[1]) * 16 + digit.index(value[2])
    a2 = digit.index(value[3]) * 16 + digit.index(value[4])
    a3 = digit.index(value[5]) * 16 + digit.index(value[6])
    return (a1, a2, a3)

im_list = []
for i in range(width):
    for j in range(length):
        r,g,b = im_resize.getpixel((j,i))
        im_list.append((r,g,b))
        if j < 26:
            cellname = str(ABC[j]+str(i+1))
            ws.column_dimensions[ABC[j]].width = 1
        else:
            cellname = str(ABC[j//26-1]+ABC[j-26*(j//26)] + str(i + 1))
        ws.row_dimensions[i+1].height = 6
        ws.column_dimensions[ABC[j//26-1]+ABC[j-26*(j//26)]].width = 1
        ws[cellname].fill = openpyxl.styles.fills.GradientFill(stop=[color((r,g,b)), color((r,g,b))])

for i in range(width):
    for j in range(length):
        No = i*length+j+1
        print(No,im_list[i*10+j])

end = time.time()
print(end-start)
wb.save("EVA.xlsx")