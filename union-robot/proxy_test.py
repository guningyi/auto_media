# -*- coding:utf-8 -*-
import urllib.request
import time


def proxy_test(ip, port, url):
    try:
        # 设置代理
        proxy_handler = urllib.request.ProxyHandler({'http': 'http://' + ip + ':' + str(port) + '/'})
        opener = urllib.request.build_opener(proxy_handler)
        urllib.request.install_opener(opener)
        # 访问网页,带10秒超时
        req = urllib.request.urlopen(url,None,3)
        print(req.getcode())
    except Exception as e:
        print("time out" + str(e))
    print("continue exec !")
    time.sleep(2)
    print("continue exec 2 !")

if __name__ == "__main__":
    proxy_test('139.224.237.33', '8888', 'http://www.baidu.com')