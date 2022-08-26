# -*- coding: utf-8 -*-
__author__ = 'newdefence@163.com'
__date__ = '2022/08/24 08:57'

import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.chrome.service import Service

PWD = os.path.abspath(os.path.dirname(__file__))

def main():
    service = Service(os.path.join(PWD, 'chromedriver'))
    browser = webdriver.Chrome(service=service)
    browser.get(r'https://www.baidu.com')
    wait = WebDriverWait(browser, 50)
    body = wait.until(ec.presence_of_element_located((By.TAG_NAME, 'body')))
    print('body: %s', body)
    # browser.quit()


if __name__ == '__main__':
    main()


