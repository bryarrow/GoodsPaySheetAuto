# -*- coding:utf-8 -*-

"""
全自动谷团肾表生成器——命令行交互界面

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
from openpyxl import load_workbook
import main


def cli():
    """
    主函数，在单独运行时处理输入输出等
    :return:
    """

    file = input('请输入排表文件路径: ')
    target = input('请输入CN区域: ')
    avg_price_cell = input('请输入写有均价的单元格（留空默认为B1）: ')

    print('\n----start{ ' + str(time.time()) + ' }-----\n')

    wb = load_workbook(file)
    arrange_sheet = wb['排表']
    payment_sheet = wb['肾表']

    payment_info = main.PaymentInfo(main.ArrangeInfo(arrange_sheet, target, avg_price_cell=avg_price_cell))
    payment_info.print_payment(payment_sheet)

    wb.save(file)

    print('\n--------------- end ---------------')

    print('运行完毕，按任意键退出。')
    input()


if __name__ == '__main__':
    cli()
