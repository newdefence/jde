# -*- coding: utf-8 -*-
__author__ = 'newdefence@163.com'
__date__ = '2022/08/18 10:53' 

import logging
from decimal import Decimal

from ordered_set import OrderedSet

logger = logging.getLogger()


def value(cell):
    val = cell.value
    if val:
        return Decimal(str(val).replace(',', ''))
    else:
        return Decimal()


def check_result_column(sheet, columns):
    if '可否入账' in columns:
        return
    max_column = sheet.max_column
    columns['可否入账'] = max_column + 1
    columns['异常信息'] = max_column + 2
    sheet.cell(1, max_column + 2, '可否入账')
    sheet.cell(1, max_column + 3, '异常信息')


def read_sheet12(sheet, fields1, fields2, sum_callback=None):
    '''发票文件读取，前面字符字段，后面数字字段，汇总行回调处理，单行特殊处理'''
    invoices, columns = {}, None  # { '发票号': 1, '物料号': 3, ... }
    # { 'invoice_no': { 'sum': { '总合计': 1, '总毛重': 2, ... }, 'details': [{ 'PO号': 'xxx', '物料号': '', ... }, ...] } }
    for row in sheet.iter_rows():
        if columns:
            # 汇总行只有一行，如果有多个，则出错
            cell0 = row[0]
            if not cell0.value:  # '' or None
                sum_callback(row, columns)
            else:
                # 只处理所需的字段，由于涉及到空单元格的问题，因此需要逐个字段处理，如果为空则按0处理
                detail = {'row': row, 'errors': OrderedSet()}
                # for key in fields1:
                #     cell = row[columns[key]]
                #     if cell.value:
                #         detail[key] = cell.value
                #     else:
                #         # 是否需要单元格背景颜色改变
                #         detail['errors'].add('%s为空' % key)
                detail.update((key, row[columns[key]].value) for key in fields1)
                # NOTE: 没有总净重
                detail.update((key, value(row[columns[key]])) for key in fields2)
                invoice = invoices.setdefault(detail['发票号'], {'details': [], 'sum': detail})
                invoice['details'].append(detail)
                # 每一个发票号明细总数量全部相等，并等于该发票号所有数量合计
        else:  # 首行表头根据中文提取字段
            columns = dict((cell.value, cell.column - 1) for cell in row if cell.value)
    check_result_column(sheet, columns)
    return invoices, columns


# def read_sheet2(sheet, fields1, fields2):
#     # 发票号，PO号，物料号，数量，单价，合计，总数量，总合计，总件数，总毛重，总净重
#     # { '发票号': 1, '物料号': 3, ... }
#     # { 'invoice_no': { 'sum': { '总合计': 1, '总毛重': 2, ... }, 'details': [{ 'PO号': 'xxx', '物料号': '', ... }, ...] } }
#     packings, columns = {}, None
#     for row in sheet.iter_rows():
#         if columns:
#             # 只处理所需的字段，由于涉及到空单元格的问题，因此需要逐个字段处理，如果为空则按0处理
#             detail = {'row': row, 'errors': set()}
#             for key in fields1:
#                 cell = row[columns[key]]
#                 if cell.value:
#                     detail[key] = cell.value
#                 else:
#                     detail['errors'].add('%s为空' % key)
#             detail.update((key, value(row[columns[key]])) for key in fields2)
#             pkg = packings.setdefault(detail['发票号'], {'details': [], 'sum': detail})
#             # NOTE 单个料单净重，毛重可能为空，所以此处比较大小没有太多意义（数据出现过某料号有净重，无毛重）
#             pkg['details'].append(detail)
#         else:  # 第二行根据中文提取字段
#             columns = dict((cell.value, cell.column - 1) for cell in row if cell.value)
#     return packings, columns


def read_sheet3(sheet, fields1, fields2):
    # 只有提运单号，处理那些字段呢？
    # { 'invoice_no': { 'sum': { '总合计': 1, '总毛重': 2, ... }, 'details': [{ 'PO号': 'xxx', '物料号': '', ... }, ...] } }
    details, columns = [], None
    for row in sheet.iter_rows():
        if columns:  # 首行表头忽略
            detail = {'row': row, 'errors': OrderedSet()}
            # for key in fields1:
            #     cell = row[columns[key]]
            #     if cell.value:
            #         detail[key] = cell.value
            #     else:
            #         detail['errors'].add('%s为空' % key)
            detail.update((key, row[columns[key]].value) for key in fields1)
            detail.update((key, value(row[columns[key]])) for key in fields2)
            details.append(detail)
        else:  # 第二行根据中文提取字段
            columns = dict((cell.value, cell.column - 1) for cell in row if cell.value)
    check_result_column(sheet, columns)
    return details, columns

