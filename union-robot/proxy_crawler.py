# coding:utf8
import urllib.request, sys, re
import threading, os
import time, datetime
import operator


#这里没有使用队列 只是采用多线程分发 对代理量不大的网页还行 但是几百几千性能就很差了
#高匿代理是做点击的最好选择、
#要清楚四种代理的区别
'''
1 透明代理虽然可以直接“隐藏”你的IP地址，但是还是可以从HTTP_X_FORWARDED_FOR来查到你是谁。
2 匿名代理比透明代理进步了一点：别人只能知道你用了代理，无法知道你是谁。
3 如上，与匿名代理相同，如果使用了混淆代理，别人还是能知道你在用代理，但是会得到一个假的IP地址，伪装的更逼真：-）
4 高匿代理让别人根本无法发现你是在用代理，所以是最好的选择。

http://www.xicidaili.com/
http://www.gvpld.cn/
'''




def get_proxy_page(url):
    '''''解析代理页面 获取所有代理地址'''
    proxy_list = []
    p = re.compile(
        r'''''<div>(.+?)<span class="Apple-tab-span" style="white-space:pre">.*?</span>(.+?)<span class="Apple-tab-span" style="white-space:pre">.+?</span>(.+?)(<span.+?)?</div>''')
    try:
        res = urllib.request.urlopen(url)
    except urllib.request.URLError:
        print('url Error')
        sys.exit(1)

    pageinfo = res.read()
    res = p.findall(pageinfo)  # 取出所有的
    # 组合成所有代理服务器列表成一个符合规则的list
    for i in res:
        ip = i[0]
        port = i[1]
        addr = i[2]
        l = (ip, port, addr)
        proxy_list.append(l)
    return proxy_list


# 同步锁装饰器
lock = threading.Lock()


def synchronous(f):
    def call(*args, **kw):
        lock.acquire()
        try:
            return f(*args, **kw)
        finally:
            lock.release()

    return call


# 时间计算器
def sumtime(f):
    def call(*args, **kw):
        t1 = time.time()
        try:
            return f(*args, **kw)
        finally:
            print
            u'总共用时 %s' % (time.time() - t1)

    return call


proxylist = []
reslist = []


# 获取单个代理并处理
@synchronous
def getoneproxy():
    global proxylist
    if len(proxylist) > 0:
        return proxylist.pop()
    else:
        return ''

        # 添加验证成功的代理


@synchronous
def getreslist(proxy):
    global reslist
    if not (proxy in reslist):
        reslist.append(proxy)


def handle():
    timeout = 10
    test_url = 'http://www.baidu.com'
    test_str = '030173'

    while 1:
        proxy = getoneproxy()
        # 最后一个返回是空
        if not proxy:
            return
        print
        u"正在验证 : %s" % proxy[0]

        # 第一步启用 cookie
        cookies = urllib.request.HTTPCookieProcessor()
        proxy_server = r'http://%s:%s' % (proxy[0], proxy[1])
        # 第二步 装载代理
        proxy_hander = urllib.request.ProxyHandler({"http": proxy_server})
        # 第三步 组合request
        try:
            opener = urllib.request.build_opener(cookies, proxy_hander)
            pass
        except urllib.request.URLError:
            print
            u'url设置错误'
            continue
            # 配置request
        opener.addheaders = [('User-Agent',
                              'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/21.0.1180.89 Safari/537.1')]
        # 发送请求
        urllib.request.install_opener(opener)
        t1 = time.time()
        try:
            req = urllib.request.urlopen(test_url, timeout=timeout)
            result = req.read()
            pos = result.find(test_str)
            timeused = time.time() - t1
            if pos > 1:
                # 保存到列表中
                getreslist((proxy[0], proxy[1], proxy[2], timeused))
                print(u'成功采集', proxy[0], timeused)
            else:
                continue
        except Exception as e:
            #print(u'采集失败 %s ：timeout' % proxy[0])
            continue


def save(reslist):
    path = os.getcwd()
    filename = path + '/Proxy-' + datetime.datetime.now().strftime(r'%Y%m%d%H%M%S') + '.txt'
    f = open(filename, 'w+')
    for proxy in reslist:
        f.write('%s %s %s %s \r\n' % (proxy[0], proxy[1], proxy[2], proxy[3]))
    f.close()


@sumtime
def main():
    url = 'http://www.free998.net/daili/httpdaili/8949.html'
    global proxylist, reslist
    # 获取所有线程
    proxylist = get_proxy_page(url)
    print('一共获取 %s 个代理' % len(proxylist))
    # print proxylist
    print('*' * 80)

    # 线程创建和分发任务
    print('开始创建线程处理.....')
    threads = []
    proxy_num = len(proxylist)

    for i in range(proxy_num):
        th = threading.Thread(target=handle, args=())
        threads.append(th)

    for thread in threads:
        thread.start()

    for thread in threads:
        threading.Thread.join(thread)

    print(u'获取有效代理 %s 个,现在开始排序和保存 ' % len(reslist))
    reslist = sorted(reslist, cmp=lambda x, y: operator.lt(x[3], y[3]))
    save(reslist)


if __name__ == '__main__':
    main()