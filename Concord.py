# -*- coding: utf-8 -*-
__author__ = 'newdefence@163.com'
__date__ = '2022/08/16 16:33' 

from decimal import Decimal
import logging
import re

from openpyxl import load_workbook

from read_sheet import read_sheet12, write_row_errors

'''
立寰：concord
只有物流园和苏州有Excel

物流园
1.（所有供应商统一逻辑）发票跟箱单检验逻辑：相同的发票号单个料号总数量校验，单个发票号总数量校验，所有发票号总数量校验
箱单上总毛重大于总净重

    Excel校验：

    发票跟excel检验
    Po号、物料号进行数量、单价和合计的检验（excel单价进行保留四位小数，四舍五入后跟发票单价进行检验）（后续关务系统，单价以excel中超过4位的为准）
'''


logger = logging.getLogger()
reKG = re.compile(r'\(([\d,.]+)\s+KG\)') # (8.760 KG), 10.211LB (1.234 KG)
reEA = re.compile(r'([\d,.]+)\s+EA') # 60,000  EA
reOF = re.compile(r'\d+\s+of\s+(\d+)') # 01 of 02
re00 = re.compile(r'^0+') # 箱单发票号前置的00需要去掉


def Decimal2(rs):
    return Decimal(rs[0].replace(',', ''))


def read_invoice(sheet):
    """文件自查"""
    # 发票号，地址，原产国，PO号，物料号，数量，单价，合计，总数量，总合计，总件数，总毛重，总净重
    # { '发票号': 1, '物料号': 3, ... }
    # { 'invoice_no': { 'sum': { '总合计': 1, '总毛重': 2, ... }, 'details': [{ 'PO号': 'xxx', '物料号': '', ... }, ...] } }

    # NOTE: 立寰发票文件没有汇总行
    invoices, columns = read_sheet12(sheet, ('发票号', 'po号', '物料号'),
                                    ('数量', '单价', '合计', '总数量', '总净重', '总合计', '总件数', '总毛重'))
    all_invoices = invoices.values()
    for invoice in all_invoices:
        sum1, details = invoice['sum'], invoice['details']
        total_qty, total_invoice_value, total_gross_weight, total_net_weight, total_pkgs = \
            sum1['总数量'], sum1['总合计'], sum1['总毛重'], sum1['总净重'], sum1['总件数']
        total_qty1 = sum(r['数量'] for r in details)
        if total_qty != total_qty1:
            tuple(d['errors'].add('SUM(数量): %s %s' % (total_qty, total_qty1)) for d in details)
        tuple(d['errors'].add('总数量: %s %s' % (d['总数量'], total_qty)) for d in details if (d['总数量'] != total_qty))
        tuple(d['errors'].add('合计: %s %s' % (d['合计'], d['数量'] * d['单价'])) for d in details if (d['合计'] != d['数量'] * d['单价']))
        total_invoice_value1 = sum(r['合计'] for r in details)
        if total_invoice_value != total_invoice_value1:
            tuple(d['errors'].add('SUM(合计): %s %s' % (total_invoice_value, total_invoice_value1)) for d in details)
        tuple(d['errors'].add('总合计: %s' % total_invoice_value) for d in details if (d['总合计'] != total_invoice_value))
        # 毛重没法比对，只能和自己比对
        tuple(d['errors'].add('总毛重: %s' % total_gross_weight) for d in details if (d['总毛重'] != total_gross_weight))
        tuple(d['errors'].add('总净重: %s' % total_net_weight) for d in details if (d['总净重'] != total_net_weight))
        tuple(d['errors'].add('总件数: %s' % total_pkgs) for d in details if (d['总件数'] != total_pkgs))
    # NOTE: 没有汇总核对需求
    all_sum ={
        '总数量': sum(v['sum']['总数量'] for v in all_invoices),
        '总合计': sum(v['sum']['总合计'] for v in all_invoices),
        '总毛重': sum(v['sum']['总毛重'] for v in all_invoices),
        '总净重': sum(v['sum']['总净重'] for v in all_invoices),
        '总件数': sum(v['sum']['总件数'] for v in all_invoices),
    }
    logger.info('发票文件核对完成')
    return columns, invoices, all_sum


def read_packing(sheet):
    packings, columns = read_sheet12(sheet, ('发票号', 'po号', '物料号', '数量', '净重', '毛重', '总数量', '总净重', '总毛重', '总件数', '总托盘数'), ('总箱数',))
    # 每一个发票号明细总数量，总净重，总毛重全部相等，总件数全部相同，并等于该发票号所有数量合计
    all_pkgs = packings.values()
    for invoice in all_pkgs:
        sum1, details = invoice['sum'], invoice['details']
        # TODO NOTE: 修正净重，毛重：只有 发票号，物料，总件数 分组的第一个才有意义，其他的归值为0
        eraser = set()
        for d in details:
            if d['总件数'] in eraser:
                d['净重'] = 0
                d['毛重'] = 0
            else:
                eraser.add(d['总件数'])
            d['数量'] = Decimal2(reEA.findall(d['数量'])) # '60,000 EA' -> 60000
            d['总数量'] = Decimal2(reEA.findall(d['总数量'])) # '60,000 EA' -> 60000
            # 正常情况下毛重净重需要处理千位分隔符问题；此处处理一下以防万一
            d['净重'] = Decimal2(reKG.findall(d['净重'])) if d['净重'] else 0 # '(5.710 KG)' -> 5.710
            d['毛重'] = Decimal2(reKG.findall(d['毛重'])) if d['毛重'] else 0
            d['总净重'] = Decimal2(reKG.findall(d['总净重'])) if d['总净重'] else 0 # '8.952 LB (4.061 KG)' -> 4.061
            d['总毛重'] = Decimal2(reKG.findall(d['总毛重'])) if d['总毛重'] else 0
            # d['_总托盘数'] = d['总托盘数']
            d['_总件数'] = Decimal(reOF.findall(d['总件数'])[0])

        total_qty, total_net_weight, total_gross_weight, total_pkgs = \
            sum1['总数量'], sum1['总净重'], sum1['总毛重'], sum1['_总件数']
        # all_sum['总毛重'] += total_gross_weight
        # all_sum['总净重'] += total_net_weight
        # 总托盘数，总箱数，总件数 不累加
        # all_sum['总托盘数'] = sum1['总托盘数']
        # all_sum['总箱数'] = sum1['总箱数']
        # all_sum['总件数'] += sum1['总件数']
        if total_qty != sum(r['数量'] for r in details):
            tuple(d['errors'].add('SUM(总数量): %s' % total_qty) for d in details)
        tuple(d['errors'].add('总数量: %s' % total_qty) for d in details if (d['总数量'] != total_qty))
        if total_net_weight != sum(r['净重'] for r in details):
            tuple(d['errors'].add('SUM(总净重): %s' % total_net_weight) for d in details)
        tuple(d['errors'].add('总净重: %s %s' % (d['总净重'], total_net_weight)) for d in details if (d['总净重'] != total_net_weight))
        if total_gross_weight != sum(r['毛重'] for r in details):
            tuple(d['errors'].add('SUM(总毛重): %s' % total_gross_weight) for d in details)
        tuple(d['errors'].add('总毛重: %s %s' % (d['总毛重'], total_gross_weight)) for d in details if (d['总毛重'] != total_gross_weight))
        tuple(d['errors'].add('总件数: %s' % total_pkgs) for d in details if (d['_总件数'] != total_pkgs))
    logger.info('箱单文件核对完成')
    all_sum = {
        '总数量': sum(v['sum']['总数量'] for v in all_pkgs),
        # '总合计': sum(v['sum']['总合计'] for v in all_pkgs),
        '总毛重': sum(v['sum']['总毛重'] for v in all_pkgs),
        '总净重': sum(v['sum']['总净重'] for v in all_pkgs),
        # '总件数': sum(v['sum']['总件数'] for v in all_pkgs),
    }
    return columns, packings, all_sum



def check(proforma_invoice, packing_list, air_waybill):
    if proforma_invoice is None:
        logger.warning('无发票文件')
        return
    if packing_list is None:
        logger.warning('无箱单文件')
        return
    file1, file2 = load_workbook(proforma_invoice), load_workbook(packing_list)
    sheet1, sheet2 = (f.worksheets[0] for f in (file1, file2))

    columns1, invoices, all_sum1 = read_invoice(sheet1)
    columns2, packings, all_sum2 = read_packing(sheet2)
    # 发票 VS 箱单
    keys1 = set(invoices.keys())
    keys2 = set(('00' + k) for k in packings.keys())
    if keys1 == keys2:
        for key in keys1:
            v1, v2 = invoices[key], packings[re00.sub('', key)]
            # NOTE: 只核对总数量，总净重（发票文件无该信息），总毛重，总件数不核对
            if v1['sum']['总数量'] == v2['sum']['总数量']:
                logger.info('发票跟箱单: %s %s总数量相同', key,
                            '相同发票号单个料号' if len(v1['details']) > 1 else '单个发票号')
            else:
                logger.warning('发票跟箱单：%s 总数量错误', key)
                msg = '发票跟提单总数量: %s %s' % (v1['sum']['总数量'], v2['sum']['总数量'])
                tuple(d['errors'].add(msg) for d in v1['details'])
                tuple(d['errors'].add(msg) for d in v2['details'])
            # 校验箱单总毛重大于总净重
            if all(d['总毛重'] > d['总净重'] for d in v2['details']):
                logger.info('发票跟箱单：%s 总毛重大于总净重', key)
            else:
                for d in v2['details']:
                    if d['总毛重'] <= d['总净重']:
                        logger.warning('发票跟箱单：%s 总毛重不大于总净重', key)
                        d['errors'].add('总毛重 总净重: %s %s' % (d['总毛重'], d['总净重']))
    else:
        def write_diff_error(host, keys, msg):
            if keys:
                for key in keys:
                    tuple(d['errors'].add(msg) for d in host[key]['details'])

        logger.warning('发票跟箱单：发票数据核对不上')
        write_diff_error(invoices, keys1 - keys2, '发票号不在箱单文件中')
        write_diff_error(packings, keys2 - keys1, '发票号不在发票文件中')

    logger.info('文件交互核对完成，开始序列化')

    # 开始写错误信息，如果有错误，则返回
    has_error = 0
    has_error += write_row_errors(sheet1, sum([x['details'] for x in invoices.values()], []), columns1)
    has_error += write_row_errors(sheet2, sum([x['details'] for x in packings.values()], []), columns2)

    if has_error:
        logger.warning('文件校验完成，错误已标注')
    else:
        logger.info('文件校验完成，没有错误信息')
    file1.save(proforma_invoice)
    file2.save(packing_list)

