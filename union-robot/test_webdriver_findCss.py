# coding=utf-8

from selenium import webdriver
import time
import os




if __name__ == '__main__':

    ieDriver = "C:\Program Files\Internet Explorer\IEDriverServer.exe"
    os.environ["webdriver.ie.driver"] = ieDriver
    browser =  webdriver.Ie(ieDriver)

    browser.get("https://www.so.com/s?src=lm&ls=sm1515585&q=%E7%89%A9%E6%B5%81%E5%BF%AB%E9%80%92&lmsid=92de997c60a756f8&lm_extend=ctype:7")
    #browser.switch_to.frame('e_idea_frame_1')  # 需先跳转到iframe框架
    #browser.find_element_by_id("e_idea_pp").send_keys("selenium")
    browser.find_elements_by_id('e_idea_pp')
    time.sleep(2)
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
    browser.close()