# -*- coding:utf-8 -*-
import urllib.request
import time
import xlrd

def read_xls(proxyfile_xls):
    wb = xlrd.open_workbook(proxyfile_xls)
    sh = wb.sheet_by_index(0)  # 第一个表
    proxy_list = []
    nrows = sh.nrows  # 获取一共有多少行
    for i in range(nrows):  # 遍历输出
        # print(sh.cell(i,0).value)
        temp = sh.cell(i, 0).value.split('@')
        # print(temp[0])
        proxy_list.append(temp[0])
    #print(proxy_list)
    return proxy_list

def proxy_test(ip, port, url):
    result = 0
    try:
        # 设置代理
        proxy_handler = urllib.request.ProxyHandler({'http': 'http://' + ip + ':' + str(port) + '/'})
        opener = urllib.request.build_opener(proxy_handler)
        urllib.request.install_opener(opener)
        # 访问网页,带10秒超时
        req = urllib.request.urlopen(url,None,3)
        print(req.getcode())
        result = req.getcode()
    except Exception as e:
        print("time out" + str(e))
    return result

if __name__ == "__main__":
    proxy_list = read_xls('C:\\Users\\eniiguu\\Desktop\\proxy.xls')
    available_proxy_list = []
    for proxy in proxy_list:
        temp = proxy.split(':')
        ip = temp[0]
        port = temp[1]
        #print(ip)
        #print(port)
        result = proxy_test(ip, port, 'http://www.baidu.com')
        if result == 200:
            available_proxy_list.append(proxy)
    print(available_proxy_list)