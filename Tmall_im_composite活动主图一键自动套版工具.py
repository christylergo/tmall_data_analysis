# -*- coding:utf-8 -*-

import win32con
import win32api
from tkinter import *
from tkinter import messagebox
import os
import xlwings as xw
from PIL import Image
from PIL import ImageFont, ImageDraw

def get_desktop():
    key = win32api.RegOpenKey(
        win32con.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', 0, win32con.KEY_READ)
    return win32api.RegQueryValueEx(key, 'Desktop')[0]


def path_info():

    desktop = get_desktop()
    file_list = os.listdir(desktop)  # 图片主文件夹要放在桌面上
    bricks_path = ''
    for item in file_list:
        if '套图大力丸' in item:
            bricks_path = desktop + "\\" + item  # 图片主文件夹地址bricks_path
    if not ('套图大力丸' in bricks_path):
        root = Tk()
        root.withdraw()  # ****实现主窗口隐藏
        messagebox.showinfo('whoops!', '图在哪里？没找到呀！')
        return 'whoops!'
    print(bricks_path)
    file_list = os.listdir(bricks_path)
    bricks = ''
    for item in file_list:
        if r'.xls' in item and not ('进度表' in item):
            bricks = bricks_path + '\\' + item  # 切换图片的表格完整路径bricks
    if not (r'.xls' in bricks):
        root = Tk()
        root.withdraw()  # ****实现主窗口隐藏
        messagebox.showinfo('呵呵', '表格都没有，这砖没法搬了！')
        return 'whoops!'
    # print(bricks)
    list_pic = []
    for item in file_list:  # file_list指向的是主文件夹内的文件及目录列表
        if os.path.isdir(bricks_path + '\\' + item) and not("套图小能手" in item):  # 要传入路径而不是文件/文件夹名称
            list_temp = os.listdir(bricks_path + '\\' + item)
            for i in range(len(list_temp)):
                list_temp[i] = item + '\\' + list_temp[i]
            list_pic.extend(list_temp)  # 把图片的相对地址放进列表 list_pic
    bricks_info = [bricks, bricks_path, list_pic]
    return bricks_info
    # ---------------以上是找出图片文件夹路径以及活动表格---------------- #


def accept_task():

    bricks, bricks_path, list_pic = path_info()
    xl_app = xw.App(visible=True, add_book=False)
    wb = xl_app.books.open(bricks)
    ws = wb.sheets[0]
    rows_count = ws.used_range.last_cell.row
    print("活动单元格行数：",rows_count)
    promotion = ws.range((1, 1), (rows_count, 20)).value  # get promotion list
    # time.sleep(1)
    wb.close()  # 读取后关闭文件
    xl_app.quit()  # 关闭打开的Excel进程
    task=[]
    for i in range(2, rows_count):
        #遇到关键信息为空的就跳过，不然会报错
        mark=0
        for j in range(0,4):
            if promotion[i][j] is None:
                mark=1
        if mark==1:
            continue
        promotion[i][0] = promotion[i][0].strip()
        promotion[i][1] = str(int(promotion[i][1])).strip()  # float转成字符后，会有.0, 要消除掉
        # print(promotion[i][1])  #tiaoshi
        for pic in list_pic:
            if "不带活动标" in pic and promotion[i][1] in pic:
                impath= bricks_path + '\\' + pic  # 二维数组promotion的第3列放置图片完整路径,第2列放置的是商品id
                maskpath=bricks_path+'\\'+promotion[i][3]+'.png'
                donepath=bricks_path+'\\'+promotion[i][0]+'_套图小能手'+'\\'+pic.split('\\')[-1]
                im_info = [promotion[i][0],promotion[i][1],impath,maskpath,donepath]
                textinfolist=[]
                for j in range(4,18,3):
                    if promotion[0][j] is None or promotion[i][j+2] is None:
                        break
                    elif "活动信息" in promotion[0][j]:
                        pos=promotion[i][j+2]
                        if "," in str(pos): #取出位置坐标元组
                            sep=","
                        elif "，" in str(pos):
                            sep="，"
                        else:
                            print('something is wrong!')
                            break
                        promotion[i][j + 2]=(int(pos.split(sep)[0]),int(pos.split(sep)[1]))
                        textinfo=[promotion[i][j],promotion[i][j+1],promotion[i][j+2]]  #textinfo=[text,fontsize,position]
                        textinfolist.append(textinfo)
                im_info.append(textinfolist)
                task.append(im_info)
                break
    return task


def im_compositer(task):
    '''
    task=[im_info, ...]
    im_info=[利益点，itemID，impath,maskpath,donepath,textinfolist]
    textinfolist=[textinfo, ...]
    textinfo=[text,fontsize,position]
    '''
    # print(task)
    for im_info in task:
        # im = Image.open(r"C:\Users\Administrator\Desktop\1.jpg").convert('RGBA')
        im = Image.open(im_info[2]).convert('RGBA')
        # mask = Image.open(r"C:\Users\Administrator\Desktop\第四版2.png")
        mask = Image.open(im_info[3])
        newim=Image.alpha_composite(im,mask)
        draw = ImageDraw.Draw(newim)
        # print(im_info[0]) #tiaoshi
        i=0
        while True:
            try:
                # # use a truetype font
                # font = ImageFont.truetype(r"C:\Windows\Fonts\SIMLI.ttf", 56)
                font = ImageFont.truetype(r"C:\Windows\Fonts\SIMLI.ttf", int(im_info[5][i][1]))
                # print(int(im_info[5][i][1])) #tiaoshi
                # draw.text((450, 500), "7-隔尿垫彩棉", font=font,fill=(125,25,255))
                draw.text(im_info[5][i][2], im_info[5][i][0], font=font, fill=(125, 25, 255))
                i+=1
            except:
                break
        # h,w=draw.textsize("7-隔尿垫彩棉",font=font)
        # draw.line([(450,500),(450+h,500)],fill=(128,0,128),width=2)
        # draw.line([(450,500+w),(450+h,500+w)],fill=(128,0,128),width=2)
        im3=newim.convert('RGB')
        # print(im_info[0]) #tiaoshi
        # im3.save(r"C:\Users\Administrator\Desktop\合并隔尿垫.jpg")
        compositerpath=os.path.dirname(im_info[3])
        foldermark=0
        for folder in os.listdir(compositerpath):
            if im_info[0] + "_套图小能手" == folder:
                im3.save(im_info[4])
                foldermark=1
                break
        if foldermark==0:
            os.mkdir(compositerpath+'\\'+im_info[0] + "_套图小能手")
            im3.save(im_info[4])
        # im3.show()

if __name__=="__main__":
    task=accept_task()
    im_compositer(task)