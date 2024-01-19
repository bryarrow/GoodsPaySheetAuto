# -*- coding:utf-8 -*-

"""
全自动谷团肾表生成器——图形交互界面

by Berry(GitHub@bryarrow)

Copyright (c) 2024 Yang Yunfei
GoodsPaySheetAuto is licensed under Mulan PSL v2.
You can use this software according to the terms and conditions of the Mulan PSL v2.
You may obtain a copy of Mulan PSL v2 at:
         https://license.coscl.org.cn/MulanPSL2
THIS SOFTWARE IS PROVIDED ON AN "AS IS" BASIS, WITHOUT WARRANTIES OF ANY KIND,
EITHER EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO NON-INFRINGEMENT,
MERCHANTABILITY OR FIT FOR A PARTICULAR PURPOSE.
See the Mulan PSL v2 for more details.
"""

import time

from tkinter import ttk
from ttkbootstrap import Style

import main


def do_sheet(file_path: str, target: str, avg_price_cell: str = ''):
    """
    按钮回调函数，发起计算
    :param file_path: 文件路径
    :param target: CN区域
    :param avg_price_cell: 调价单元格
    """

    print('\n----start{ ' + str(time.time()) + ' }-----\n')

    wb = main.xlwings.Book(file_path)
    arrange_sheet = wb.sheets[0]
    payment_sheet = wb.sheets[1]

    payment_info = main.PaymentInfo(main.ArrangeInfo(arrange_sheet, target, avg_price_cell=avg_price_cell))
    payment_info.print_payment(payment_sheet)

    print('\n--------------- end ---------------')


def gui():
    style = Style()
    root = style.master
    root.title('GoodsPaySheetAuto——谷团自动打肾表工具')
    frm = ttk.Frame(root, padding=10)
    frm.grid()

    ttk.Label(frm, text='请输入表的路径：').grid(column=0, row=0)
    path_entry = ttk.Entry(frm)
    path_entry.grid(column=0, row=1)

    ttk.Label(frm, text='请输入CN区域：').grid(column=0, row=2)
    target_entry = ttk.Entry(frm)
    target_entry.grid(column=0, row=3)

    ttk.Label(frm, text='请输入均价单元格：').grid(column=0, row=4)
    avg_price_entry = ttk.Entry(frm)
    avg_price_entry.grid(column=0, row=5)

    ttk.Button(frm,
               text="GO",
               command=lambda: do_sheet(path_entry.get(), target_entry.get(), avg_price_entry.get())
               ).grid(column=1, row=0, rowspan=6)

    root.mainloop()


if __name__ == '__main__':
    gui()
