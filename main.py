# # -*- coding:utf-8 -*-
#
# import xlrd
#
#
# # 打开Excel文件
# def open_excel(file_path):
#     try:
#         data = xlrd.open_workbook(file_path)
#         return data
#     except Exception as e:
#         print("文件打开失败!"+str(e))
#
#
# # 根据名称获取Excel表格中的数据   参数:file：Excel文件路径     by_index：通过索引顺序获取，默认第一个索引
# def excel_table_byindex(file, by_index=0):
#     data = open_excel(file)
#     table = data.sheet_by_index(by_index)  # 根据索引来获取excel中的sheet
#     nrows = table.nrows  # 行数
#     ncols = table.ncols  # 列数
#     list = []  # 装读取结果的序列
#     for rownum in range(0, nrows):  # 遍历每一行的内容
#          row = table.row_values(rownum)  # 根据行号获取行
#          if row:
#              app = [] # 某一行
#              for i in range(0, ncols):  # 一列列地读取行的内容
#                 app.append(row[i])
#              list.append(app)  # 装载数据
#     return list
#
#
# def main():
#     excel_1 = excel_table_byindex("rawData/吉力沈辽路站OS管控流水20180402.xlsx", 0)
#     excel_2 = excel_table_byindex("rawData/吉力沈辽路站正星管控流水20180402.XLS", 0)
#     print(len(excel_1))
#
#     print(excel_2)
#
#
#
# if __name__ == "__main__":
#     main()



#!/usr/bin/env python
# -*- coding: utf-8 -*-

from tkinter import *
#from tkColorChooser import askcolor
__author__ = 'Chuntao Lu'


class Paint(object):

    DEFAULT_PEN_SIZE = 5.0
    DEFAULT_COLOR = 'black'

    def __init__(self):
        self.root = Tk()

        self.pen_button = Button(self.root, text='pen', command=self.use_pen)
        self.pen_button.grid(row=0, column=0)

        self.brush_button = Button(self.root, text='brush', command=self.use_brush)
        self.brush_button.grid(row=0, column=1)

        self.color_button = Button(self.root, text='color', command=self.choose_color)
        self.color_button.grid(row=0, column=2)

        self.eraser_button = Button(self.root, text='eraser', command=self.use_eraser)
        self.eraser_button.grid(row=0, column=3)

        self.choose_size_button = Scale(self.root, from_=1, to=10, orient=HORIZONTAL)
        self.choose_size_button.grid(row=0, column=4)

        self.c = Canvas(self.root, bg='white', width=600, height=600)
        self.c.grid(row=1, columnspan=5)

        self.setup()
        self.root.mainloop()

    def setup(self):
        self.old_x = None
        self.old_y = None
        self.line_width = self.choose_size_button.get()
        self.color = self.DEFAULT_COLOR
        self.eraser_on = False
        self.active_button = self.pen_button
        self.c.bind('<B1-Motion>', self.paint)
        self.c.bind('<ButtonRelease-1>', self.reset)

    def use_pen(self):
        self.activate_button(self.pen_button)

    def use_brush(self):
        self.activate_button(self.brush_button)

    def choose_color(self):
        self.eraser_on = False
        self.color = self.colorchooser.askcolor(color=self.color)[1]

    def use_eraser(self):
        self.activate_button(self.eraser_button, eraser_mode=True)

    #TODO: reset canvas
    #TODO: undo and redo
    #TODO: draw triangle, rectangle, oval, text

    def activate_button(self, some_button, eraser_mode=False):
        self.active_button.config(relief=RAISED)
        some_button.config(relief=SUNKEN)
        self.active_button = some_button
        self.eraser_on = eraser_mode

    def paint(self, event):
        self.line_width = self.choose_size_button.get()
        paint_color = 'white' if self.eraser_on else self.color
        if self.old_x and self.old_y:
            self.c.create_line(self.old_x, self.old_y, event.x, event.y,
                               width=self.line_width, fill=paint_color,
                               capstyle=ROUND, smooth=TRUE, splinesteps=36)
        self.old_x = event.x
        self.old_y = event.y

    def reset(self, event):
        self.old_x, self.old_y = None, None


if __name__ == '__main__':
    ge = Paint()