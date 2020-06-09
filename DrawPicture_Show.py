import os
import win32com
from win32com.client import constants as c  # 旨在直接使用VBA常数
import time
from PIL import Image
from threading import Thread
from multiprocessing import Process
ratio = 1920/1080
length = 240
width = int(length/ratio)
'''导入并resize图片'''
def input_im(name):
    im = Image.open(name,"r")
    print(im.size,im.format,im.mode)
    #im.show()
    im_resize = im.resize((length,width))
    return im_resize

'''打开excel，并设置行高列宽'''
current_address = os.path.abspath('.')
excel_address = os.path.join(current_address, "EVA.xlsx")
print(current_address)
xl_app = win32com.client.gencache.EnsureDispatch("Excel.Application")  # 若想引用常数的话使用此法调用Excel
#xl_app.Visible = True  # 是否显示Excel文件
wb = xl_app.Workbooks.Open(excel_address)
sht = wb.Worksheets(1)
sht.Name = "EVA"


ABC = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
'''
RGB 与 16进制 相互转换
input：（255，255，255）
output: 'ABCDEF' 
'''
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
def rgb_to_hex(rgb):
    bgr = (rgb[2], rgb[1], rgb[0])
    strValue = '%02x%02x%02x' % bgr
    # print(strValue)
    iValue = int(strValue, 16)
    return iValue

start = time.time()

def S(thread):
    thread.start()


def my_threadfun(j,im_resize):
    for i in range(width):
        r,g,b = im_resize.getpixel((j,i))
        # if i < 26:
        #     cellname = str(ABC[i]+str(j+1))
        # else:
        #     cellname = str(ABC[i//26-1]+ABC[i-26*(i//26)] + str(j + 1))
        # sht.Range(cellname).Interior.Color = rgb_to_hex((r,g,b))
        sht.Range(sht.Cells(i+1,j+1),sht.Cells(i+1,j+1)).Interior.Color = rgb_to_hex((r,g,b))

if __name__ == '__main__':

    for x in range(1034):
        x += 1
        # if x % 5 == 0:
        #     length += 1
        #     width = int(length/ratio)
        start_x = time.time()
        im_name = 'venv/EVA_PICTURE/EVA/'+'EVA_'+str(x)+'.0.jpg'
        im_resize = input_im(im_name)
        # sht.Range(sht.Cells(1, 1), sht.Cells(width, length)).Interior.Pattern = c.xlNone#清除单元格格式
        for j in range(length):
            thread_j = Thread(target=my_threadfun(j,im_resize))
           # thread_j = Process(target=my_threadfun,args=(j,))
            S(thread_j)
        xl_app.Visible = True  # 是否显示Excel文件
        end_x = time.time()
        print('#',x,end_x-start_x)

xl_app.Visible = True  # 是否显示Excel文件
wb.Save()

end = time.time()

print('T',end-start)
#wb.Close()