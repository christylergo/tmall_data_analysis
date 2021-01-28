# -*- coding:utf-8 -*-

#  chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenium\AutomationProfile"  #

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import *
import xlwings as xw
from multiprocessing import Process
import time
import os
import pyautogui
import win32gui
import win32con
import win32api
from tkinter import *
from tkinter import messagebox
import numpy
from greenlet import greenlet


def set_action_time():
    while True:
        str_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        str_hour = str_time[-8:-6]
        str_minute = str_time[-5:-3]
        if str_hour == '19' and str_minute == "14":
            break
        time.sleep(5)


def initial_chrome():
    os.system(r'chrome.exe --remote-debugging-port=9222')
    print('------Chrome浏览器已经就绪-------')


def sesame_shell(func):
    print('Simple is better than complex.')
    reference = 'The Zen of Python, by Tim Peters'

    def sesame_pickup(*args, **kwargs):
        print('Complex is better than complicated.')
        browser, tabs = func(*args, **kwargs)
        reset_frame(browser)
        sesame = browser.find_element_by_xpath(
            '//div[@class="sidebar expandable"]/ul/li[3]/div//li/a').text[-4]
        switch_frame(browser)
        browser.find_element_by_xpath(
            '//tr[@class="next-table-row first"]/td[7]//div[contains(text(),' + sesame + ')]')
        reset_frame(browser)
        print('In the face of ambiguity, refuse the temptation to guess.')
        print(reference)  # 强行贴近闭包的要素，说到不做到
        browser = [browser, tabs]
        return browser
    return sesame_pickup


@sesame_shell
def connect_chrome():
    time.sleep(5)
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    chrome_driver = r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe"
    browser = webdriver.Chrome(chrome_driver, options=chrome_options)
    browser.maximize_window()
    time.sleep(0.5)
    browser.get(r'https://www.163.com/')
    time.sleep(1.5)
    browser.execute_script(r'window.open("https://news.163.com/")')  # 用js打开新标签页
    time.sleep(2)
    tabs = browser.window_handles  # 当前全部标签的窗口句柄
    browser.switch_to.window(tabs[0])
    time.sleep(0.5)
    open_door(browser)
    browser.switch_to.window(tabs[1])
    open_door(browser)
    time.sleep(10)  # 补足第二个标签页打开后的等待时间
    switch_frame(browser)
    btb = [browser, tabs]
    print('----------已连接浏览器------------')
    return btb


def open_door(browser):
    browser.get(r'https://web.cbbs.tmall.com/')
    time.sleep(6)
    item_holding = browser.find_element_by_xpath('//li[@data-id="m2081"]/a')
    item_holding.click()
    time.sleep(6)
    item_holding = browser.find_element_by_xpath('//a[text()="商品列表"]')
    item_holding.click()
    time.sleep(8)


def initial_rect(browser):
    initial_dom = browser.find_element_by_xpath('//div[@class="content-container full-height"]/div[2]/div[1]')
    offset_x = browser.execute_script("return arguments[0].getBoundingClientRect().left;",
                                      initial_dom)  # iframe的初始位置相对于整个浏览器的偏移量
    offset_y = browser.execute_script("return arguments[0].getBoundingClientRect().bottom;", initial_dom)
    initial_x = browser.execute_script("return  window.screenLeft;") + 5  # 经过大量对比测试，这两句最有效的获取网页元素的初始位置
    initial_y = browser.execute_script("return window.screenTop + window.outerHeight - document.body.clientHeight;") - 8
    rect = (initial_x, initial_y, offset_x, offset_y)
    print('-------已经初始化坐标位置-------')
    return rect


def switch_frame(browser):
    WebDriverWait(browser, 25).until(
        EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//iframe[@class=" iframe-page normal-iframe show"]')))


def reset_frame(browser):
    browser.switch_to.default_content()  # Back to the root html


def before_operation(browser, item_id):
    WebDriverWait(browser, 25).until(
        EC.presence_of_element_located((By.XPATH, '//a[text()="编辑"]')))  # 必须等待加载后再进行Xpath操作
    door = browser.find_element_by_xpath('//div[@label="前端商品Id"]/div[@class="next-form-item-control"]/span/input')
    door.send_keys(Keys.CONTROL + 'a', Keys.DELETE)  # 此方法清除输入框，代替clear方法。# door.clear() 不好用，在此处完全无效
    door.send_keys(item_id)  # 输入商品id
    time.sleep(0.5)
    browser.find_element_by_xpath('//button[text()="查询"]').click()
    time.sleep(2.5)
    WebDriverWait(browser, 25).until(
        EC.presence_of_element_located((By.XPATH, '//a[text()="' + item_id + '"]')))  # 必须等待加载后再进行Xpath操作
    edit = browser.find_element_by_xpath('//tr[@class="next-table-row last first"]/td[last()]/div/div/span[2]/a')
    edit.click()  # 点击编辑，进入编辑界面
    time.sleep(1)


def operate_chains(browser, rect):
    reset_frame(browser)
    switch_frame(browser)
    WebDriverWait(browser, 25).until(EC.presence_of_element_located((By.ID, "product-card-title")))  # 必须等待加载后再进行Xpath操作
    browser.find_element_by_xpath('//ul[@class="next-menu-content"]/li[1]').click()  # 点击产品信息标签页，这样在1080p分辨率的电脑上才会显示出确认框
    time.sleep(0.5)
    try:
        confirm_item = browser.find_element_by_xpath(
            '//span[@class="sell-o-checkbox sell-checkbox"]/span/label/label/input')
        if not confirm_item.is_selected():  # 如果没有确认就点确认
            confirm_item.click()
            # confirm_item.send_keys(Keys.SPACE)  # 取消选择状态
        time.sleep(0.5)
    except NoSuchElementException:
        print('这个不需要确认产品信息=>' + browser.find_element_by_xpath('//div[contains(text(),"展示效果")]').text)
    description = browser.find_element_by_xpath('//ul[@class="next-menu-content"]/li[5]')  # 选择商品描述标签页
    description.click()
    actions = ActionChains(browser)
    time.sleep(1)
    anchor1 = browser.find_element_by_xpath('//label[text()="新版手机端描述"]')
    browser.execute_script("arguments[0].scrollIntoView();", anchor1)  # 页面先向下滚动，让需要点击的元素显示出来
    # ---------------------查找、点击第1个启用/不启用-------------------------- #
    try:
        pic_description = browser.find_element_by_xpath(
            '//div[@id="struct-modularWirelessDesc"]//div[@class="modulePanel-editAction" and text()="商品图片"]')
        actions.move_to_element(pic_description).perform()  # 启用/不启用在mouseover事件后才会显示
        pic_description = browser.find_element_by_xpath(
            '//div[@id="struct-modularWirelessDesc"]//div[@class="modulePanel-editAction" and text()="商品图片"]/span/span[contains(text(),"启用")]')
        time.sleep(0.2)  # 留出一点时间就可以看到点击产生的变化
        if pic_description.text == "不启用":
            pic_description.click()
            print("------图片启用不启用-------")
        # ----------------------------------------------------------------------- #
        time.sleep(0.5)
        # ---------------------查找、点击第2个启用/不启用-------------------------- #
        item_description = browser.find_element_by_xpath(
            '//div[@id="struct-modularWirelessDesc"]//div[@class="modulePanel-editAction" '
            'and text()="商品信息"]')  # 启用/不启用在mouseover事件后才会显示
        actions.move_to_element(item_description).perform()  # 启用/不启用在mouseover事件后才会显示
        time.sleep(0.2)  # 留出一点时间就可以看到点击产生的变化
        item_description = browser.find_element_by_xpath(
            '//div[@id="struct-modularWirelessDesc"]//div[@class="modulePanel-editAction" '
            'and text()="商品信息"]/span/span[contains(text(),"启用")]')  # 启用/不启用在mouseover事件后才会显示
        if item_description.text == "不启用":
            item_description.click()
            print("------信息启用不启用-------")
    except NoSuchElementException:
        print('这个没有启用不启用=>' + browser.find_element_by_xpath('//div[contains(text(),"展示效果")]').text)
    # ----------------------------------------------------------------------- #
    # --------------------------找到主图并点击更换----------------------------- #
    time.sleep(0.2)  # 留出一点时间避免浏览器来不及响应
    anchor2 = browser.find_element_by_css_selector('h2#desc-card-title')
    browser.execute_script("arguments[0].scrollIntoView();", anchor2)  # 页面先向上滚动，让需要点击的主图元素显示出来
    main_pic = browser.find_element_by_xpath(
        '//div[@class="tmall-o-image-placeholder required" and text()="商品主图"]/../div[3]/div')
    time.sleep(0.2)  # 留出一点时间避免浏览器来不及响应
    delete_pic = main_pic.find_element_by_xpath('./i[@class="next-icon next-icon-ashbin next-icon-small tool delete"]')
    initial_x, initial_y, offset_x, offset_y = rect
    x = browser.execute_script("return arguments[0].getBoundingClientRect().left;",
                               delete_pic) + initial_x + offset_x + 20
    y = browser.execute_script("return arguments[0].getBoundingClientRect().top;",
                               delete_pic) + initial_y + offset_y + 20

    print(x, y)
    pyautogui.moveTo(x, y, 0.2)  # 主图更换/删除选项在mouseover事件后才会显示
    time.sleep(0.1)
    pyautogui.click()
    pyautogui.moveTo(x, y - 50, 0.2)  # 主图更换/删除选项在mouseover事件后才会显示
    replace_pic = browser.find_element_by_xpath(
        '//div[@class="tmall-o-image-placeholder required" and text()="商品主图"]/../div[2]/i')
    y = browser.execute_script("return arguments[0].getBoundingClientRect().top;",
                               replace_pic) + initial_y + offset_y + 20
    pyautogui.moveTo(x, y, 0.1)
    pyautogui.click()
    time.sleep(1)
    # ----------------------------------------------------------------------- #


def upload_pic(path_pic):
    # -----------------------使用pywin32库操作上传文件------------------------- #
    dialog = win32gui.FindWindow("#32770", "打开")  # 对话框
    hwnd = win32gui.FindWindowEx(dialog, 0, 'ComboBoxEx32', None)
    combobox = win32gui.FindWindowEx(hwnd, 0, 'ComboBox', None)
    edit_path = win32gui.FindWindowEx(combobox, 0, 'Edit', None)
    win32gui.SendMessage(edit_path, win32con.WM_SETTEXT, None, path_pic)
    time.sleep(0.5)
    open_btn = win32gui.FindWindowEx(dialog, 0, 'Button', '打开')
    win32gui.SendMessage(dialog, win32con.WM_COMMAND, 1, open_btn)
    time.sleep(1.5)
    # ----------------------------------------------------------------------- #


def confirm_submit(browser):
    # --------------------------------提交------------------------------- #
    submit = browser.find_element_by_xpath('//button[@id="button-submit"]')
    # submit.click()
    time.sleep(2.5)
    try:
        WebDriverWait(browser, 6).until(
            EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"商品编辑成功，宝贝ID为")]')))  # 必须等待加载后再进行Xpath操作
        browser.find_element_by_xpath('//span[contains(text(),"商品编辑成功，宝贝ID为")]')
        browser.switch_to.default_content()  # Back to the root html
        browser.find_element_by_xpath('//span[text()="阿里巴巴供应链平台"]/../span[3]').click()
        return '完成'
    except NoSuchElementException:
        browser.switch_to.default_content()  # Back to the root html
        browser.find_element_by_xpath('//span[text()="阿里巴巴供应链平台"]/../span[3]').click()
        return '没提交进去'
    # ------------------------------------------------------------------- #
    print('----------confirm------------')


def get_desktop():
    key = win32api.RegOpenKey(
        win32con.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', 0, win32con.KEY_READ)
    return win32api.RegQueryValueEx(key, 'Desktop')[0]


def path_info():

    desktop = get_desktop()
    file_list = os.listdir(desktop)  # 图片主文件夹要放在桌面上
    bricks_path = ''
    for item in file_list:
        if '搬砖大力丸' in item:
            bricks_path = desktop + "\\" + item  # 图片主文件夹地址bricks_path
    if not ('搬砖大力丸' in bricks_path):
        root = Tk()
        root.withdraw()  # ****实现主窗口隐藏
        messagebox.showinfo('whoops!', '好想搬砖，可是没找到呀！')
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
    print(bricks)
    list_pic = []
    for item in file_list:  # file_list指向的是主文件夹内的文件及目录列表
        if os.path.isdir(bricks_path + '\\' + item):  # 要传入路径而不是文件/文件夹名称
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
    promotion = ws.range((2, 1), (rows_count, 4)).value  # get promotion list
    time.sleep(1)
    wb.close()  # 读取后关闭文件
    xl_app.quit()  # 关闭打款的Excel进程
    for i in range(0, rows_count - 1):
        promotion[i][1] = (str(promotion[i][1]).split('.', 1))[0].strip()  # float转成字符后，会有.0, 要消除掉
        promotion[i][3] = None
        for pic in list_pic:
            if promotion[i][0] in pic and promotion[i][1] in pic:
                promotion[i][3] = bricks_path + '\\' + pic  # 二维数组promotion的第4列放置图片完整路径,第2列放置的是商品id

    buff = numpy.hsplit(numpy.array(promotion), 4)
    task = numpy.hstack((buff[1], buff[3]))
    return task


if __name__ == '__main__':
    print('--------主进程开始-------')
    # set_action_time()
    task_new = accept_task()
    # p = Process(target=initial_chrome, args=())  # 这里不能有参数，初始化Chrome不需要任何参数
    # p.start()  # 在单独进程中开启Chrome浏览器，不这样不能连接
    browser_ins, browser_TAB = connect_chrome()
    rect_ins = initial_rect(browser_ins)
    # --------------------------------------- #
    bricks_wr = path_info()[1]
    xl_app_wr = xw.App(visible=True, add_book=False)
    wb_wr = xl_app_wr.books.add()
    time.sleep(0.2)
    wb_wr.save(bricks_wr + r'\进度表.xlsx')
    time.sleep(0.1)
    wxl = xl_app_wr.hwnd  # 获取 xlapp的 窗口句柄
    win32gui.ShowWindow(wxl, win32con.SW_MINIMIZE)  # 结合win32的api来实现窗口最小化。可能有更简单的方法现在没发现
    ws_wr = wb_wr.sheets[0]
    # --------------------------------------- #
    task_len = len(task_new)
    half_len = task_len // 2
    on_off = 1  # 协程开关
    labour_force_two_mark = 0  # 第二只搬砖gou注定无法完成循环，用于记录协程结束时挂起的点

    def labour_force_one(tabs=''):
        for i in range(0, half_len):  # 两只搬砖gou各做一半，当task_len是奇数时第二只会多做一个
            print('-----dog_1 go!------')
            item_id_ins, path_pic_ins = task_new[i]
            if path_pic_ins is None:
                continue
            switch_frame(browser_ins)
            before_operation(browser_ins, item_id_ins)

            if on_off == 1:
                browser_ins.switch_to.window(tabs[1])
                time.sleep(1)
                print('-----切换到dog_2------')
                dog2.switch(tabs)  # 协程的切换点  # 注意协程切换时传递的参数

            if i == 0:     # 第一次切换时是从二号没有挂起点，是从头开始做，消耗时间短所以额外加个等待
                time.sleep(8)

            operate_chains(browser_ins, rect_ins)
            upload_pic(path_pic_ins)
            # -------------------------------------------- #
            task_new[i][1] = confirm_submit(browser_ins)
            ws_wr.range((i+2, 1), (i+2, 2)).value = task_new[i]
            time.sleep(0.1)
            wb_wr.save()
            # -------------------------------------------- #
            print('-------labour_1 换图完成---------')
            # print(task_new)
            time.sleep(3)

    def labour_force_two(tabs=''):
        global labour_force_two_mark  # python的特点，在函数内部修改了全局变量，如果不显式声明全局变量的话，会被当做局部变量，报错
        for j in range(half_len, task_len):
            print('------dog_2 go!-------')
            item_id_ins, path_pic_ins = task_new[j]
            if path_pic_ins is None:
                continue
            switch_frame(browser_ins)
            if labour_force_two_mark != -1:  # 当mark等于-1时意味着labour_2在协程结束后第一次进入循环，不需要再进行before_operation操作
                before_operation(browser_ins, item_id_ins)
            if on_off == 1:
                browser_ins.switch_to.window(tabs[0])
                time.sleep(1)
                print('-----切换到dog_1------')
                labour_force_two_mark = j
                dog1.switch(tabs)  # 协程的切换点 # 注意协程的参数
            else:
                if labour_force_two_mark != -1:
                    time.sleep(10)  # 当协程关掉以后，直接单个循环，没有切换消耗时间，所以要额外加等待
                labour_force_two_mark = j  # 循环一次就会改变mark，避免最后调用labour_2 的时候 mark多次被当成-1

            operate_chains(browser_ins, rect_ins)
            upload_pic(path_pic_ins)
            # -------------------------------------------- #
            task_new[j][1] = confirm_submit(browser_ins)
            ws_wr.range((j+2, 1), (j+2, 2)).value = task_new[j]
            time.sleep(0.1)
            wb_wr.save()
            # -------------------------------------------- #
            print('-------labour_2 换图完成---------')
            # print(task_new)
            time.sleep(3)
    # ------------------------------------------ #

    browser_ins.switch_to.window(browser_TAB[0])
    dog1 = greenlet(labour_force_one)
    dog2 = greenlet(labour_force_two)
    dog1.switch(browser_TAB)
    on_off = 0  # 关闭协程switch，成为普通函数
    browser_ins.switch_to.window(browser_TAB[1])  # 单独运行第二个labour要先切换到第二个标签页
    half_len = labour_force_two_mark
    labour_force_two_mark = -1  # 用于标记babour_2 是不是在刚刚结束协程时进入循环，此时网页的状态处于点击编辑后
    time.sleep(0.5)  # 切换标签页之后加个缓冲
    labour_force_two()
    # ------------------------------------------ #
    wb_wr.save()
    wb_wr.close()  # 写入后后关闭文件
    xl_app_wr.quit()  # 关闭打开的Excel进程
    # ------------------------------------------ #
    Tk().withdraw()  # ****实现主窗口隐藏
    messagebox.showinfo('wow', '砖头搬完啦，可以快乐的玩耍了！')