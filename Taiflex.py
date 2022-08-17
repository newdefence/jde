# -*- coding: utf-8 -*-
__author__ = 'newdefence@163.com'
__date__ = '2022/08/16 17:43' 

import logging
from decimal import Decimal
from functools import reduce

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


def read_invoice(sheet):
    """文件自查"""
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
            tuple(d['errors'].add('总数量错误') for d in details)
        if not all(r['总数量'] == total_qty for r in details):
            tuple(d['errors'].add('总数量错误') for d in details)
        if total_invoice_value != sum(r['合计'] for r in details):
            tuple(d['errors'].add('总合计错误') for d in details)
        if not all(r['总合计'] == total_invoice_value for r in details):
            tuple(d['errors'].add('总合计错误') for d in details)
        # 毛重没法比对，只能和自己比对
        if not all(r['总毛重'] == total_gross_weight for r in details):
            tuple(d['errors'].add('总毛重错误') for d in details)
        if not all(r['总件数'] == total_pkgs for r in details):
            tuple(d['errors'].add('总件数错误') for d in details)
    if all_sum['总数量'] != sum(v['sum']['总数量'] for v in all_invoices):
        all_sum['errors'].add('总数量汇总错误')
    if all_sum['总合计'] != sum(v['sum']['总合计'] for v in all_invoices):
        all_sum['errors'].add('总合计汇总错误')
    if all_sum['总毛重'] != sum(v['sum']['总毛重'] for v in all_invoices):
        all_sum['errors'].add('总毛重汇总错误')
    # NOTE: 没有总件数和总净重的核对需求
    logger.info('发票文件核对完成')
    return columns, invoices, all_sum


def read_packing(sheet):
    # 发票号，PO号，物料号，数量，单价，合计，总数量，总合计，总件数，总毛重，总净重
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
    all_pkgs, all_sum = packings.values(), {'总毛重': 0, '总净重': 0}
    for invoice in all_pkgs:
        sum1, details = invoice['sum'], invoice['details']
        total_qty, total_net_weight, total_gross_weight, total_pkgs = sum1['总数量'], sum1['总净重'], sum1['总毛重'], sum1['总件数']
        all_sum['总毛重'] += total_gross_weight
        all_sum['总净重'] += total_net_weight
        # 总托盘数，总箱数，总件数 不累加
        all_sum['总托盘数'] = sum1['总托盘数']
        all_sum['总箱数'] = sum1['总箱数']
        all_sum['总件数'] = sum1['总件数']
        if total_qty != sum(r['数量'] for r in details):
            tuple(d['errors'].add('总数量错误') for d in details)
        if not all(r['总数量'] == total_qty for r in details):
            tuple(d['errors'].add('总数量错误') for d in details)
        if total_net_weight != sum(r['净重'] for r in details):
            tuple(d['errors'].add('总净重错误') for d in details)
        if not all(r['总净重'] == total_net_weight for r in details):
            tuple(d['errors'].add('总净重错误') for d in details)
        if total_gross_weight != sum(r['毛重'] for r in details):
            tuple(d['errors'].add('总毛重错误') for d in details)
        if not all(r['总毛重'] == total_gross_weight for r in details):
            tuple(d['errors'].add('总毛重错误') for d in details)
        if not all(r['总件数'] == total_pkgs for r in details):
            tuple(d['errors'].add('总件数错误') for d in details)
    logger.info('箱单文件核对完成')
    return columns, packings, all_sum

def read_air(sheet):
    # 逻辑如何处理，只有提运单号，处理那些字段呢？
    columns = None  # { '发票号': 1, '物料号': 3, ... }
    # { 'invoice_no': { 'sum': { '总合计': 1, '总毛重': 2, ... }, 'details': [{ 'PO号': 'xxx', '物料号': '', ... }, ...] } }
    details = []
    for row in sheet.iter_rows():
        if columns:  # 首行表头忽略
            errors = set()
            detail = {'row': row, 'errors': errors}
            for key in ('主运单号', '分运单号'):
                cell = row[columns[key]]
                if cell.value:
                    detail[key] = cell.value
                else:
                    errors.add('%s为空' % key)
            detail.update((key, value(row[columns[key]])) for key in ('托盘数', '总毛重', '总件数'))
            details.append(detail)
        else:  # 第二行根据中文提取字段
            columns = dict((cell.value, cell.column - 1) for cell in row if cell.value)
    all_sum = {
        '托盘数': sum(d['托盘数'] for d in details),
        '总毛重': sum(d['总毛重'] for d in details),
        '总件数': sum(d['总件数'] for d in details),
    }
    return columns, details, all_sum


def write_row_errors(sheet, details, columns):
    col1, col2 = columns['可否入账'] + 1, columns['异常信息'] + 1
    has_error = False
    for d in details:
        errors, warnings = d['errors'], d.get('warnings')
        if errors:
            has_error = True
        if warnings:
            errors = errors + warnings
        if errors:
            sheet.cell(d['row'][0].row, col2, '，'.join(errors))
        else:
            sheet.cell(d['row'][0].row, col1, '可入账')
    return has_error


def check(proforma_invoice, packing_list, air_waybill):
    if proforma_invoice is None:
        logger.warning('无发票文件')
        return
    if packing_list is None:
        logger.warning('无箱单文件')
        return
    file1, file2, file3 = load_workbook(proforma_invoice), load_workbook(packing_list), (load_workbook(air_waybill) if air_waybill else None)
    sheet1, sheet2, sheet3 = (f.worksheets[0] if f else None for f in (file1, file2, file3))

    columns1, invoices, all_sum1 = read_invoice(sheet1)
    columns2, packings, all_sum2 = read_packing(sheet2)
    columns3, airs, all_sum3 = read_air(sheet3) if air_waybill else (None, None, None)
    # 发票 VS 箱单
    keys1 = invoices.keys()
    keys2 = packings.keys()
    if keys1 == keys2:
        for key in keys1:
            v1, v2 = invoices[key], packings[key]
            # NOTE: 只核对总数量，总净重（发票文件无该信息），总毛重，总件数不核对
            if v1['sum']['总数量'] == v2['sum']['总数量']:
                logger.info('发票跟箱单：%s %s总数量相同', v1['sum']['发票号'], '相同发票号单个料号' if len(v1['details']) > 1 else '单个发票号')
            else:
                tuple(d['errors'].add('发票跟提单总数量错误') for d in v1['details'])
                tuple(d['errors'].add('发票跟提单总数量错误') for d in v2['details'])
            # 校验箱单总毛重大于总净重
            if all(d['总毛重'] > d['总净重'] for d in v2['details']):
                logger.info('发票跟箱单：总毛重大于总净重')
            else:
                for d in v2['details']:
                    if d['总毛重'] <= d['总净重']:
                        d['errors'].add('总毛重不大于总净重')
    else:
        logger.warning('发票跟箱单：发票数据核对不上')
        def write_diff_error(host, keys, msg):
            if keys:
                for key in keys:
                    tuple(d['errors'].add(msg) for d in host[key]['details'])
        write_diff_error(invoices, keys1 - keys2, '发票号不在箱单文件中')
        write_diff_error(packings, keys2 - keys1, '发票号不在发票文件中')

    # 箱单 VS 提单，如果出现错误，只写提单
    if air_waybill:
        def write_air_errors(msg):
            tuple(d['errors'].add(msg) for d in airs)
        if all_sum2['总托盘数'] != all_sum3['托盘数']:
            write_air_errors('托盘数与箱单不符')
        # 目前文件没有 总箱数 此列
        print('TODO：提单文件无`总箱数`列')
        if all_sum2['总箱数'] != all_sum3.get('总箱数'):
            print('TODO：提单文件无`总箱数`列？')
            write_air_errors('总箱数与箱单不符:该文件无`总箱数`列')
        # 总毛重误差3%内正常
        diff_gross_weight = all_sum2['总毛重'] - all_sum3['总毛重']
        if diff_gross_weight:
            if diff_gross_weight >= all_sum2['总毛重'] * Decimal(0.03):
                write_air_errors('总毛重与箱单不符：超过3%')
            else:
                tuple(d.setdefault('warnings', set()).add('总毛重与箱单误差3%以内') for d in airs)
        if all_sum2['总净重'] >= all_sum3['总毛重']:
            write_air_errors('箱单总净重需小于提单总毛重')
    logger.info('文件交互核对完成，开始序列化')

    # 开始写错误信息，如果有错误，则返回
    has_error = 0
    has_error += write_row_errors(sheet1, reduce(lambda x, y: x + y['details'], invoices.values(), []), columns1)
    has_error += write_row_errors(sheet2, reduce(lambda x, y: x + y['details'], packings.values(), []), columns2)
    if air_waybill:
        has_error += write_row_errors(sheet3, airs, columns3)
    if has_error:
        logger.warning('文件校验完成，错误已标注')
    else:
        logger.info('文件校验完成，没有错误信息')
    file1.save(proforma_invoice)
    file2.save(packing_list)
    if air_waybill:
        file3.save(air_waybill)
