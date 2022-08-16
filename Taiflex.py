# -*- coding: utf-8 -*-
__author__ = 'newdefence@163.com'
__date__ = '2022/08/16 17:43' 

import logging
from decimal import Decimal

from openpyxl import load_workbook


'''
台虹：（需要多识别一个出售合同）taiflex
    有提单则为空运（提单上的除了发票和箱单外还有其他的页则为空运）
    物流园
    1.（所有供应商统一逻辑）发票跟箱单检验逻辑：相同的发票号单个料号总数量校验，单个发票号总数量校验，所有发票号总数量校验
    箱单上总毛重大于总净重

    发票跟出售合同检验：各个发票号中的总数量和总合计、条款进行合同检验

空运：
1.（所有供应商统一逻辑）发票跟箱单检验逻辑：相同的发票号单个料号总数量校验，单个发票号总数量校验，所有发票号总数量校验；
箱单上总毛重大于总净重
2.箱单跟提单：箱单上的总托盘数和总箱数还有总毛重（上下差百分之三合理范围，写入识别文件）进行提单校验
3.箱单总净重小于提单上的总毛重（如果大于等于则提示异常并返回）
'''


logger = logging.getLogger()


def value(cell):
    return Decimal(str(cell.value or 0))


def empty_cell(cell):
    return cell.value is None or cell.value == ''


def read_invoice(workbook):
    """文件自查"""
    sheet = workbook.worksheets[0]
    # 发票号，地址，原产国，PO号，物料号，数量，单价，合计，总数量，总合计，总件数，总毛重，总净重
    columns = None  # { '发票号': 1, '物料号': 3, ... }
    # { 'invoice_no': { 'sum': { '总合计': 1, '总毛重': 2, ... }, 'details': [{ 'PO号': 'xxx', '物料号': '', ... }, ...] } }
    invoices, all_sum = {}, None
    for row in sheet.iter_rows():
        if columns:
            # 汇总行只有一行，如果有多个，则出错
            cell0 = row[0]
            if not cell0.value: # '' or None
                all_sum = {
                    'row': row, 'errors': set(),
                    '总数量': value(row[columns['总数量']]),
                    '总合计': value(row[columns['总合计']]),
                    '总毛重': value(row[columns['总毛重']]),
                    # '总件数': value(row[columns['总件数']]),
                    # NOTE: 总件数，总净重没有
                }
            else:
                # 只处理所需的字段，由于涉及到空单元格的问题，因此需要逐个字段处理，如果为空则按0处理
                errors = set()
                detail = {'row': row, 'errors': errors}
                for key in ('发票号', 'po号', '物料号'):
                    cell = row[columns[key]]
                    if cell.value:
                        detail[key] = cell.value
                    else:
                        # 是否需要单元格背景颜色改变
                        errors.add('%s为空' % key)
                # NOTE: 没有总净重
                detail.update((key, value(row[columns[key]])) for key in ('数量', '单价', '合计', '总数量', '总合计', '总件数', '总毛重'))
                # if cell.value is None or cell.value == '':
                #     # 是否需要单元格颜色背景改变？
                #     row_error.append('%s为空' % key)
                # else:
                #     detail[key] = Decimal(0) if isinstance(cell, EmptyCell) else value(cell)
                invoice = invoices.setdefault(cell0.value, {'details': [], 'sum': detail})
                invoice['details'].append(detail)
                # NOTE: 浮点数计算问题不可忽略
                if detail['合计'] != detail['数量'] * detail['单价']:
                    errors.add('合计计算错误')
                # 每一个发票号明细总数量全部相等，并等于该发票号所有数量合计
        else:  # 首行表头根据中文提取字段
            columns = dict((cell.value, cell.column - 1) for cell in row if cell.value)
    all_invoices = invoices.values()
    for invoice in all_invoices:
        sum1, details = invoice['sum'], invoice['details']
        total_qty, total_invoice_value, total_gross_weight, total_pkgs = sum1['总数量'], sum1['总合计'], sum1['总毛重'], sum1['总件数']
        if total_qty != sum(r['数量'] for r in details):
            map(lambda d: d['errors'].add('总数量错误'), details)
        if not all(r['总数量'] == total_qty for r in details):
            map(lambda d: d['errors'].add('总数量错误'), details)
        if total_invoice_value != sum(r['合计'] for r in details):
            map(lambda d: d['errors'].add('总合计错误'), details)
        if not all(r['总合计'] == total_invoice_value for r in details):
            map(lambda d: d['errors'].add('总合计错误'), details)
        # 毛重没法比对，只能和自己比对
        if not all(r['总毛重'] == total_gross_weight for r in details):
            map(lambda d: d['errors'].add('总毛重错误'), details)
        if not all(r['总件数'] == total_pkgs for r in details):
            map(lambda d: d['errors'].add('总件数错误'), details)
    if all_sum['总数量'] != sum(v['sum']['总数量'] for v in all_invoices):
        all_sum['errors'].add('总数量汇总错误')
    if all_sum['总合计'] != sum(v['sum']['总合计'] for v in all_invoices):
        all_sum['errors'].add('总合计汇总错误')
    if all_sum['总毛重'] != sum(v['sum']['总毛重'] for v in all_invoices):
        all_sum['errors'].add('总毛重汇总错误')
    # NOTE: 没有总件数和总净重的核对需求
    logger.info('发票文件核对完成')


def read_packing(workbook):
    # 发票号，PO号，物料号，数量，单价，合计，总数量，总合计，总件数，总毛重，总净重
    sheet = workbook.worksheets[0]
    columns = None  # { '发票号': 1, '物料号': 3, ... }
    # { 'invoice_no': { 'sum': { '总合计': 1, '总毛重': 2, ... }, 'details': [{ 'PO号': 'xxx', '物料号': '', ... }, ...] } }
    packings = {}
    for row in sheet.iter_rows():
        if columns:
            # 只处理所需的字段，由于涉及到空单元格的问题，因此需要逐个字段处理，如果为空则按0处理
            errors = set()
            detail = {'row': row, 'errors': errors}
            for key in ('发票号', 'po号', '物料号'):
                cell = row[columns[key]]
                if cell.value:
                    detail[key] = cell.value
                else:
                    errors.add('%s为空' % key)
            detail.update((key, value(row[columns[key]])) for key in ('数量', '净重', '毛重', '总数量', '总净重', '总毛重', '总件数', '总托盘数', '总箱数'))
            pkg = packings.setdefault(detail['发票号'], {'details': [], 'sum': detail})
            # NOTE 单个料单净重，毛重可能为空，所以此处比较大小没有太多意义（数据出现过某料号有净重，无毛重）
            pkg['details'].append(detail)
        else:  # 第二行根据中文提取字段
            columns = dict((cell.value, cell.column - 1) for cell in row if cell.value)
    # 每一个发票号明细总数量，总净重，总毛重全部相等，总件数全部相同，并等于该发票号所有数量合计
    all_pkgs = packings.values()
    for invoice in all_pkgs:
        sum1, details = invoice['sum'], invoice['details']
        total_qty, total_net_weight, total_gross_weight, total_pkgs = sum1['总数量'], sum1['总净重'], sum1['总毛重'], sum1['总件数']
        if total_qty != sum(r['数量'] for r in details):
            map(lambda d: d['errors'].add('总数量错误'), details)
        if not all(r['总数量'] == total_qty for r in details):
            map(lambda d: d['errors'].add('总数量错误'), details)
        if total_net_weight != sum(r['净重'] for r in details):
            map(lambda d: d['errors'].add('总净重错误'), details)
        if not all(r['总净重'] == total_net_weight for r in details):
            map(lambda d: d['errors'].add('总净重错误'), details)
        if total_gross_weight != sum(r['毛重'] for r in details):
            map(lambda d: d['errors'].add('总毛重错误'), details)
        if not all(r['总毛重'] == total_gross_weight for r in details):
            map(lambda d: d['errors'].add('总毛重错误'), details)
        if not all(r['总件数'] == total_pkgs for r in details):
            map(lambda d: d['errors'].add('总件数错误'), details)
    logger.info('箱单文件核对完成')


def read_air(workbook):
    # 逻辑如何处理，只有提运单号，处理那些字段呢？
    sheet = workbook.worksheets[0]
    columns = None  # { '发票号': 1, '物料号': 3, ... }
    # { 'invoice_no': { 'sum': { '总合计': 1, '总毛重': 2, ... }, 'details': [{ 'PO号': 'xxx', '物料号': '', ... }, ...] } }
    packings = {}
    for row in sheet.iter_rows():
        if columns:  # 首行表头忽略
            errors = set()
            detail = {'row': row, 'errors': errors}
            for key in ('主提运单号', '分提运单号'):
                cell = row[columns[key]]
                if cell.value:
                    detail[key] = cell.value
                else:
                    errors.add('%s为空' % key)
            detail.update((key, value(row[columns[key]])) for key in ('托盘数', '总毛重', '总件数'))
            pkg = packings.setdefault(detail['主提运单号'], {'details': [], 'sum': detail})
            pkg['details'].append(detail)
        else:  # 第二行根据中文提取字段
            columns = dict((cell.value, cell.column - 1) for cell in row if cell.value)
    # TODO 确认：汇总逻辑，主运单/分运单如何与箱单和发票核对
    all_pkgs = packings.values()
    for invoice in all_pkgs:
        sum1, details = invoice['sum'], invoice['details']
        total_gross_weight, total_pkgs = sum1['总毛重'], sum1['总件数']
        if not all(r['总毛重'] == total_gross_weight for r in details):
            raise Exception('总毛重错误')
        if not all(r['总件数'] == total_pkgs for r in details):
            raise Exception('总件数错误')
    logger.info('空运文件核对完成？校验规则在哪里？')


def check(proforma_invoice, packing_list, air_warbill):
    if proforma_invoice is None:
        logger.warning('无发票文件')
        return
    invoice = load_workbook(proforma_invoice)
    packing = load_workbook(packing_list) if packing_list else None
    air = load_workbook(air_warbill) if air_warbill else None

    read_invoice(invoice)
    if packing:
        read_packing(packing)
    if air:
        read_air(air)
    # TODO 3个文件交互验证，比对数据
    logger.info('TODO: 交互验证，比对文件')

