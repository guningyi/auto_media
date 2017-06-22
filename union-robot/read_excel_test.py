import xlrd



if __name__ == "__main__":
    proxyfile_xls = 'C:\\Users\\eniiguu\\Desktop\\proxy.xls'
    wb = xlrd.open_workbook(proxyfile_xls)
    sh = wb.sheet_by_index(0)  # 第一个表
    # cellName0 = sh.cell(0, 0).value #第0列第0行
    # print(cellName0)
    # cellName1 = sh.cell(1, 0).value #第0列第1行
    # print(cellName1)
    proxy_list = []
    nrows = sh.nrows #获取一共有多少行
    for i in range(nrows): #遍历输出
        #print(sh.cell(i,0).value)
        temp = sh.cell(i, 0).value.split('@')
        #print(temp[0])
        proxy_list.append(temp[0])
    print(proxy_list)