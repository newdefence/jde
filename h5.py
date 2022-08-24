# -*- coding: utf-8 -*-
__author__ = 'newdefence@163.com'
__date__ = '2022/08/24 08:57'

import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.edge.service import Service

PWD = os.path.abspath(os.path.dirname(__file__))

def main():
    service = Service(os.path.join(PWD, 'msedgedriver'))
    browser = webdriver.Edge(service=service)
    browser.get(r'file:///Users/defecne/JDE/%E8%AF%86%E5%88%AB%E7%BB%93%E6%9E%9C/requirements.txt')
    wait = WebDriverWait(browser, 5)
    body = wait.until(ec.presence_of_element_located((By.TAG_NAME, 'body')))
    print('body: %s', body)
    browser.quit()


if __name__ == '__main__':
    main()


