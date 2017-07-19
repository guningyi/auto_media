# coding=utf-8

from selenium import webdriver
import time
import os
from selenium.webdriver.common.proxy import Proxy
from selenium.webdriver.common.proxy import ProxyType




if __name__ == '__main__':

    #设置代理 111.155.116.195:8123
    print('设置Firefox代理:' + '111.155.116.195' + ':' + str(8123))
    profile = webdriver.FirefoxProfile()
    profile.set_preference('network.proxy.type', 1)
    profile.set_preference('network.proxy.http', '111.155.116.195')
    profile.set_preference('network.proxy.http_port', 8123)  # int
    profile.update_preferences()

    browser =  webdriver.Firefox(firefox_profile=profile)
    browser.get("https://www.so.com/s?src=lm&ls=sm1515585&q=%E7%89%A9%E6%B5%81%E5%BF%AB%E9%80%92&lmsid=92de997c60a756f8&lm_extend=ctype:7")

    #使用css来精确定位href
    inputs = browser.find_elements_by_css_selector("#e_idea_pp li h3 a")
    print(len(inputs))
    for input in inputs:
        print(input.get_attribute('href'))
        time.sleep(2)
        input.click()
    time.sleep(3)
    sreach_window = browser.current_window_handle #此行代码用来定位当前页面

    #browser.find_element_by_xpath("/html/body/div[3]/div[4]/div/div[3]/div[4]/h3/a").click()
    time.sleep(5)
    #browser.close()