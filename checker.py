# -*- coding: utf-8 -*-
__author__ = 'newdefence@163.com'
__date__ = '2022/08/19 10:25'

# def check_sum1_sum2(details, key, sum1):
#     sum2 = sum(d[key] for d in details)

def check_qty12(invoice):
    sum1, details = invoice['sum'], invoice['details']
    total_qty = sum1['总数量']
    total_qty1 = sum(d['数量'] for d in details)
    if total_qty != total_qty1:
        tuple(d['errors'].add('SUM(数量): %s %s' % (total_qty, total_qty1)) for d in details)
    tuple(d['errors'].add('总数量: %s %s' % (d['总数量'], total_qty)) for d in details if (d['总数量'] != total_qty))


def check_invoice_value1(invoice):
    sum1, details = invoice['sum'], invoice['details']
    total_invoice_value = sum1['总合计']
    total_invoice_value1 = sum(d['合计'] for d in details)
    tuple(d['errors'].add('合计: %s %s' % (d['合计'], d['数量'] * d['单价'])) for d in details if (d['合计'] != d['数量'] * d['单价']))
    if total_invoice_value != total_invoice_value1:
        tuple(d['errors'].add('SUM(总合计): %s %s' % (total_invoice_value, total_invoice_value1)) for d in details)
    tuple(d['errors'].add('总合计: %s %s' % (d['总合计'], total_invoice_value)) for d in details if (d['总合计'] != total_invoice_value))


def check_gross_weight1(invoice):
    total_gross_weight = invoice['sum']['总毛重']
    tuple(d['errors'].add('总毛重: %s %s' % (d['总毛重'], total_gross_weight)) for d in invoice['details'] if (d['总毛重'] != total_gross_weight))


def check_gross_weight2(invoice):
    sum1, details = invoice['sum'], invoice['details']
    total_gross_weight = sum1['总毛重']
    total_gross_weight1 = sum(d['毛重'] for d in details)
    if total_gross_weight != total_gross_weight1:
        tuple(d['errors'].add('SUM(毛重): %s %s' % (total_gross_weight, total_gross_weight1)) for d in details)
    tuple(d['errors'].add('总毛重: %s %s' % (d['总毛重'], total_gross_weight)) for d in details if (d['总毛重'] != total_gross_weight))


def check_net_weight1(invoice):
    total_net_weight = invoice['sum']['总净重']
    tuple(d['errors'].add('总净重: %s %s' % (d['总净重'], total_net_weight)) for d in invoice['details'] if (d['总净重'] != total_net_weight))


def check_net_weight2(invoice):
    sum1, details = invoice['sum'], invoice['details']
    total_net_weight = sum1['总净重']
    total_net_weight1 = sum(d['净重'] for d in details)
    if total_net_weight != total_net_weight1:
        tuple(d['errors'].add('SUM(净重): %s %s' % (total_net_weight, total_net_weight1)) for d in details)
    tuple(d['errors'].add('总净重: %s %s' % (d['总净重'], total_net_weight)) for d in details if (d['总净重'] != total_net_weight))


def check_piece12(invoice): # 总件数
    piece = invoice['sum']['总件数']
    tuple(d['errors'].add('总件数: %s %s' % (d['总件数'], piece)) for d in invoice['details'] if (d['总件数'] != piece))


def check_1_2_qty(invoice1, invoice2): # 发票文件发票号，箱单文件发票号
    sum1, sum2 = invoice1['sum']['总数量'], invoice2['sum']['总数量']
    if sum1 == sum2:
        # logger.info('发票跟箱单：%s %s总数量相同', v1['sum']['发票号'], '相同发票号单个料号' if len(v1['details']) > 1 else '单个发票号')
        return True
    else:
        msg = '发票VS提单总数量: %s %s' % (sum1, sum2)
        tuple(d['errors'].add(msg) for d in invoice1['details'])
        tuple(d['errors'].add(msg) for d in invoice2['details'])
        return False


def check_2_net_gross_weight(details2):
    ok = True
    for d in details2:
        if d['总毛重'] <= d['总净重']:
            ok = False
            d['errors'].add('总毛重:%s ≤ 总净重: %s' % (d['总毛重'], d['总净重']))
    return ok