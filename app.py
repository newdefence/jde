# -*- coding: utf-8 -*-
__author__ = 'newdefence@163.com'
__date__ = '2022/08/11 13:30' 

# import sys
# reload(sys)
# sys.setdefaultencoding('utf8')

import os
import logging

import Taiflex as test
# region 日志模块配置
LOG_FILE= '校对日志.log'
LOG_FMT = logging.Formatter('%(asctime)s - %(funcName)s - %(lineno)s - %(levelname)s - %(message)s')
# 日志配置 1
# logging.basicConfig(filename = LOG_FILE, filemode = 'w', format = LOG_FMT, level = logging.INFO)
# logger = logging.getLogger()

# 日志配置 2
logger = logging.getLogger()
# 文件日志输出
file_hander = logging.FileHandler(LOG_FILE)
file_hander.setLevel(logging.WARN)
file_hander.setFormatter(LOG_FMT)
logger.addHandler(file_hander)
# 控制台日志输出
logger.setLevel(logging.DEBUG) # 控制台输出必备？
console_hander = logging.StreamHandler()
console_hander.setLevel(logging.INFO)
console_hander.setFormatter(LOG_FMT)
logger.addHandler(console_hander)
# endregion

# 当前目录：根据入口文件确定当前工作路径 os.getcwd() 会限定当前目录
ROOT_PWD = os.path.dirname(os.path.abspath(__file__))
logger.info('当前目录：%s', ROOT_PWD)

# def check(proforma_invoice, packing_list, air_warbill):
#     logging.info('%s\n%s\n%s', proforma_invoice, packing_list, air_warbill)

def main():
    for item in os.listdir(ROOT_PWD):
        root_dir = os.path.join(ROOT_PWD, item)
        if os.path.isdir(root_dir):
            if not item.startswith('Taiflex'):
                continue
            target_dir = os.path.join(root_dir, '识别结果')
            target_files = os.listdir(target_dir)
            proforma_invoice, packing_list, air_warbill = None, None, None
            # Mitsui
            for file_name in target_files:
                if file_name.startswith('~'):
                    # 临时文件，忽略
                    continue
                if file_name.endswith('_AirWarbill.xlsx'):
                    air_warbill = os.path.join(target_dir, file_name)
                elif file_name.endswith("_PackingList.xlsx"):
                    packing_list = os.path.join(target_dir, file_name)
                elif file_name.endswith('_ProformaInvoice.xlsx'):
                    proforma_invoice = os.path.join(target_dir, file_name)
            test.check(proforma_invoice, packing_list, air_warbill)
            break


if __name__ == '__main__':
    main()
