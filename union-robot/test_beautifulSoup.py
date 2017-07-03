# coding:utf8
import urllib.request, sys, re
import threading, os
import time, datetime
import operator
from bs4 import BeautifulSoup


if __name__ == '__main__':
    url = 'http://www.xicidaili.com/nn/2'
    #global proxylist, reslist
    page = urllib.request.urlopen(url)
    html = page.read()
    soup = BeautifulSoup(html)
    table_tags = soup.find_all('table')
    print(table_tags)