# -*- coding:utf-8 -*-

"""
全自动谷团肾表生成器——核心处理部分

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

import pypinyin
import openpyxl.utils
from openpyxl.worksheet.worksheet import Worksheet

__title__ = 'GoodsPaySheetAuto'
__version__ = '0.1.0'
__author__ = 'Berry'
__all__ = ['ArrangeInfo', 'PaymentInfo']  # 实在想不到Mami类拿出去有什么用，还可能触发ERR02，不导出（main()不导出更是废话）


class ArrangeInfo:
    """
    排表信息类

    保存完整的排表对象以及CN，角色等数据。

    构造方法：
        使用排表实例、目标CN区域与均价信息构造该类。详见该方法文档。

    方法：
        - get_cns: 返回cn列表，以拼音排序。
        - get_chrs: 返回角色列表，按排表顺序。
        - get_float_prices: 返回调价，按排表顺序。
        - get_avg_price: 返回均价
        - get_target: 返回目标区域，例如“D3:I53”
        - get_arrange_sheet: 返回本实例所用排表实例

    """

    # 创建时使用的排表实例
    __arrange_sheet: Worksheet

    # CN区域开始列
    __start_col: str
    # CN区域开始行
    __start_line: int
    # CN区域结束列
    __end_col: str
    # CN区域结束行
    __end_line: int

    # CN列表
    __cns: list[str] = []
    # 角色列表
    __chrs: list[str] = []
    # 均价
    __avg_price: float = 0.0
    # 调价列表
    __float_prices: list[float] = []

    def __is_float_price_good(self):  # pylint: disable=unused-private-member; 待完成函数暂不使用
        """
        根据全表判断调价是否配平
        :return: 表示是否配平的Bool值，是：True，否：False
        """
        # TODO(https://github.com/users/bryarrow/projects/2/views/1?pane=issue&itemId=50222093): 根据全表判断调价是否配平

    def __is_float_prices_good_sum(self):
        """
        根据调价加和是否为0判断调价是否配平
        :return: 表示是否配平的Bool值，是：True，否：False
        """
        return sum(self.__float_prices) == 0 if True else False  # pylint: disable=using-constant-test; 纯纯误报

    def __init__(self, arrange_sheet: Worksheet, target: str, avg_price=-1.0, avg_price_cell=''):
        """
        使用排表实例、目标CN区域与均价信息构造该类。

        :param arrange_sheet: openpyxl.worksheet.worksheet.Worksheet类型的排表实例
        :param target: 排表中写有CN的区域，如示例文件中的‘D3:I53’

        :param avg_price: 【可选】均价数值，与avg_price_cell二选一使用，
        如果都填写那么将忽略avg_price_cell，如果都不填写那么默认使用B1单元格的值

        :param avg_price_cell: 【可选】写有均价的单元格，如“B1”，
        该单元格内容应当仅为数字，比如“12.3”而不是“均13.3”或“均价12.3”等

        :raises TypeError(ERR00): 在指定使用单元格内容作为均价，而该单元格内容不是数字时抛出
        :raises TypeError(ERR01): 调价部分有非数字时抛出
        """
        self.__arrange_sheet = arrange_sheet

        self.__start_col, self.__start_line = openpyxl.utils.cell.coordinate_from_string(target.split(':')[0])
        self.__end_col, self.__end_line = openpyxl.utils.cell.coordinate_from_string(target.split(':')[1])

        if avg_price == -1 and avg_price_cell == '':
            print('INFO: 未指定均价，将以‘B1’单元格的值为均价')
            try:
                self.__avg_price = float(arrange_sheet['B1'].value)
            # 当单元格内容为字符串时，float()会抛出ValueError，单元格是空单元格时，会抛出TypeError
            # 对于float()这是合理的，但对这里来说都相当于类型错误，所以添加相同的错误信息并外抛TypeError
            except ValueError as e:
                raise TypeError('ERR00:‘B1’单元格内容不是数字') from e
            except TypeError as e:
                raise TypeError('ERR00:‘B1’单元格内容不是数字') from e
        elif avg_price != -1:
            self.__avg_price = avg_price
        elif avg_price_cell != '':
            try:
                self.__avg_price = float(arrange_sheet[avg_price_cell].value)
            except ValueError as e:
                raise TypeError('ERR00:‘' + avg_price_cell + '’单元格内容不是数字') from e
            except TypeError as e:
                raise TypeError('ERR00:‘' + avg_price_cell + '’单元格内容不是数字') from e

        # 因为openpyxl对切片始终返回tuple(tuple(Cell))，写成map就变成map套lambda套map了，很丑，还是放弃了紧凑改成写循环
        for line in arrange_sheet[f'B3:B{self.__end_line}']:
            for cell in line:
                if cell.value is not None and cell.value != '':
                    self.__chrs.append(str(cell.value))

        try:
            for line in arrange_sheet[f'C3:C{self.__end_line}']:
                for cell in line:
                    self.__float_prices.append(float(cell.value))
        # 这里对异常的处理同上一条注释，因为是不同的部分出错所以使用了不同的错误代码
        except ValueError as e:
            raise TypeError(f'ERR01: 调价部分（C3:C{self.__end_line}）有非数字，请检查') from e
        except TypeError as e:
            raise TypeError(f'ERR01: 调价部分（C3:C{self.__end_line}）有非数字，请检查') from e
        if not self.__is_float_prices_good_sum():
            print('WARRING: 调价加和不为0，请检查')

        for line in arrange_sheet[target]:
            for cn_cell in line:
                if cn_cell.value is not None and cn_cell.value not in self.__cns:
                    self.__cns.append(cn_cell.value)
        self.__cns.sort(key=lambda x: pypinyin.slug(x).lower())

    def get_cns(self):
        """
        获取所有CN
        :return: 一个按拼音首字母序存放所有CN的字符串列表
        """
        return self.__cns

    def get_chrs(self):
        """
        获取所有角色名
        :return: 一个按排表顺序存放所有角色名的字符串列表
        """
        return self.__chrs

    def get_float_prices(self):
        """
        获取调价列表
        :return: 一个按排表顺序存放调价的浮点数列表，该列表保证所有成员仅为float型
        """
        return self.__float_prices

    def get_avg_price(self):
        """
        获取均价
        :return: 浮点数表示的均价
        """
        return self.__avg_price

    def get_target(self):
        """
        获取CN区域
        :return: 表示CN所在区域的EXCEL格式的字符串
        """
        # TODO(https://github.com/users/bryarrow/projects/2/views/1?pane=issue&itemId=50222164):
        #   从存储整个原表并传递CN区域改为仅存储CN区域（我还是觉得不存储排表只存储这部分区域比较好，但现在懒得改了）
        return f'{self.__start_col}{self.__start_line}:{self.__end_col}{self.__end_line}'

    def get_arrange_sheet(self):
        """
        获取本实例所用排表
        :return: 一个排表xlwings.Sheet实例（这东西真的让我中文模块损坏，咋表达通顺啊）
        """
        return self.__arrange_sheet


class Mami:
    """
    每位妈咪的信息~

    保存了CN、本表排谷数及详情和zfb应肾金额

    构造方法：
        初始化各项数据，详见其注释。

    方法：
        - get_cn(): 获取CN
        - get_payment_zfb(): 获取总应肾zfb金额
        - get_arrange_num(): 获取本表排谷数
        - get_arrange(): 以字符串形式获取排谷详情
        - add_arrange(chr_index: int, pay_zfb: float):
        添加排谷信息，chr_index:角色序号；pay_zfb:这一个应肾zfb金额

    """

    __cn: str
    __pay_zfb: float
    __arrange: list[int]  # 这个列表的第n个元素代表排表顺序第n个角色，其值代表这个实例的妈咪吃了多少。
    __arrange_num: int

    def __init__(self, cn: str, chr_num: int):
        """
        初始化CN和排谷列表

        :param cn: CN
        :param chr_num: 本表所有角色数目
        """
        self.__cn = cn
        self.__pay_zfb = 0
        self.__arrange = [0] * chr_num
        self.__arrange_num = 0

    def get_cn(self):
        """
        :return: CN字符串
        """
        return self.__cn

    def get_pay_zfb(self):
        """
        :return: 浮点数表示的总应肾zfb金额
        """
        return self.__pay_zfb

    def get_arrange_num(self):
        """
        :return: 整数表示的排谷数
        """
        return self.__arrange_num

    def get_arrange(self, chr_list: list[str]):
        """
        :raise ValueError(ERR02): 在传入的角色列表长度和类内部记录的角色数不同时抛出，建议检查构造时所用参数与调用时参数
        :return: 字符串表示的排谷详情，例如："鸣上岚cheese2 月永雷欧knights1"
        """

        # 我也不知道什么情况下会出现这种情况（除非有人两次传的不一样），不过是有可能的，防一手
        # ISSUE: 所以要不要让Mami类自带角色列表呢，从省内存角度上我不希望这样做，但是这样做好像可以简化排谷详情相关的部分
        if len(chr_list) != len(self.__arrange):
            raise ValueError('ERR02: 在Mami.get_arrange()中传入的角色列表长度与自身排表长度不同（请联系开发者）')

        arrange_str = ''
        for i in range(len(self.__arrange)):  # pylint: disable=consider-using-enumerate
            # 这里i是作为第几个角色使用，所以使用range而非迭代器，关闭pylint警告
            if self.__arrange[i] != 0:
                arrange_str += chr_list[i] + str(self.__arrange[i]) + ' '
                # TODO(https://github.com/users/bryarrow/projects/2/views/1?pane=issue&itemId=50227501):
                #   根据角色名决定是否添加‘-’
                # TODO(https://github.com/users/bryarrow/projects/2/views/1?pane=issue&itemId=50227582):
                #   添加分组+角色名的命名方式
        if arrange_str != '':
            arrange_str = arrange_str[:-1]
        return arrange_str

    def add_arrange(self, chr_index: int, pay_zfb: float):
        """
        添加排谷。

        :param chr_index: 角色在排表里是第几个，从0开始
        :param pay_zfb: 本次添加的谷应肾zfb金额
        """
        self.__arrange[chr_index] += 1
        self.__arrange_num += 1
        self.__pay_zfb += pay_zfb


class PaymentInfo:
    """
    肾表信息类

    这个类是EXCEL肾表的结构化信息，由排表信息类ArrangeInfo构造并可输出至EXCEL

    构造方法：
        接受一个ArrangeInfo排表实例作为参数构造肾表。

    方法：
        - print_payment(payment_sheet: xlwings.Sheet): 输出肾表到payment_sheet这个xlwings.Sheet实例对应的ECXCEL表格
    """

    __mamis: list[Mami] = []
    __chrs_list: list[str]

    def __init__(self, arrange_info: ArrangeInfo):
        """
        接受一个ArrangeInfo排表实例作为参数构造肾表。
        :param arrange_info: 排表实例
        """
        self.__chrs_list = arrange_info.get_chrs()

        for cn in arrange_info.get_cns():
            self.__mamis.append(Mami(cn, len(self.__chrs_list)))

        t = 0
        for line in arrange_info.get_arrange_sheet()[arrange_info.get_target()]:
            for cn_cell in line:
                if cn_cell.value is not None:
                    self.__mamis[arrange_info.get_cns().index(str(cn_cell.value))].add_arrange(
                        t,
                        arrange_info.get_avg_price() + arrange_info.get_float_prices()[t]
                    )
            t = t + 1

    def print_payment(self, payment_sheet: Worksheet):
        """
        肾表输出方法
        :param payment_sheet: 输出的目的EXCEL表格对应的xlwings.Sheet实例
        """
        for i in range(len(self.__mamis)):  # pylint: disable=consider-using-enumerate
            # 这里i是作为肾表第几行使用，故不使用迭代器
            payment_sheet["A" + str(i + 3)].value = pypinyin.slug(self.__mamis[i].get_cn())[0].upper()
            payment_sheet["B" + str(i + 3)].value = self.__mamis[i].get_cn()
            payment_sheet["C" + str(i + 3)].value = self.__mamis[i].get_arrange_num()
            payment_sheet["D" + str(i + 3)].value = self.__mamis[i].get_pay_zfb()

            # E列是肾WX，因为如果带冷备注在CN上现在需要手动合并，这里写了对应EXCEL公式:
            # =IF(D{i+3}=0,0,IF(D{i+3}>100,D{i+3}+ROUND(D{i+3}*0.001,2),D{i+3}+0.1))
            payment_sheet["E" + str(i + 3)].value = \
                f'=IF(D{i+3}=0,0,IF(D{i+3}>100,D{i+3}+ROUND(D{i+3}*0.001,2),D{i+3}+0.1))'

            payment_sheet["F" + str(i + 3)].value = self.__mamis[i].get_arrange(self.__chrs_list)
            payment_sheet["G" + str(i + 3)].value = self.__mamis[i].get_arrange_num()
            payment_sheet["H" + str(i + 3)].value = self.__mamis[i].get_pay_zfb()

            # 同E列
            payment_sheet["I" + str(i + 3)].value = \
                f'=IF(H{i+3}=0,0,IF(H{i+3}>100,H{i+3}+ROUND(H{i+3}*0.001,2),H{i+3}+0.1))'

            payment_sheet["J" + str(i + 3)].value = self.__mamis[i].get_cn()
            payment_sheet["K" + str(i + 3)].value = pypinyin.slug(self.__mamis[i].get_cn())[0].upper()
        # 写入总计公式
        payment_sheet["C" + str(len(self.__mamis) + 3)] = "=SUM(C3:C" + str(len(self.__mamis) + 2) + ")"
        payment_sheet["D" + str(len(self.__mamis) + 3)] = "=SUM(D3:D" + str(len(self.__mamis) + 2) + ")"
        payment_sheet["G" + str(len(self.__mamis) + 3)] = "=SUM(G3:G" + str(len(self.__mamis) + 2) + ")"
        payment_sheet["H" + str(len(self.__mamis) + 3)] = "=SUM(H3:H" + str(len(self.__mamis) + 2) + ")"
