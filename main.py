import time, pandas
import xlwings as xw
from selenium import webdriver
from bs4 import BeautifulSoup as bs

def getStockCode(url):
    driver = webdriver.Chrome()

    driver.get(url)

    stockCode = []
    while True:
        stockCode.extend(list(pandas.read_html(driver.page_source)[1][1][1:]))
        try:
            driver.find_element_by_link_text(u"下一页").click()
            time.sleep(1)
        except:
            print('end')
            break
    for i, j in enumerate(stockCode):
        stockCode[i] = '\'' + j
    stockCode.insert(0, '股票代码')
    print(stockCode)
    xw.Range('A1').options(transpose=True).value = stockCode

    driver.quit()

def getBasicInfo():
    driver = webdriver.Chrome()
    try:
        sz_url = 'http://quote.eastmoney.com/sz'
        sh_url = 'http://quote.eastmoney.com/sh'

        jiancheng = []
        hangye =[]
        shangshishijian = []
        shizhi = []
        jingzichan = []
        jinglirun = []
        shiyinglv = []
        shijinglv = []
        maolilv = []
        roe = []


        codes = []
        for i in xw.Range('A2:A1000'):
            if i.value == None:
                break
            codes.append(i.value)

        for code in codes:
            if int(code)>600000:
                driver.get(sh_url+code+'.html')
            else:
                driver.get(sz_url+code+'.html')

            jiancheng.append(driver.find_element_by_xpath('//*[@id="name"]').text)
            hangye.append(driver.find_element_by_xpath('/html/body/div[9]/div/div[1]/a[3]').text)
            shangshishijian.append(driver.find_element_by_xpath('//*[@id="rtp2"]/tbody/tr[10]/td').text.split('：')[1])
            shizhi.append(driver.find_element_by_xpath('/html/body/div[13]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[1]/td[2]').text)
            jingzichan.append(driver.find_element_by_xpath('/html/body/div[13]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[1]/td[3]').text)
            jinglirun.append(driver.find_element_by_xpath('/html/body/div[13]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[1]/td[4]').text)
            shiyinglv.append(driver.find_element_by_xpath('/html/body/div[13]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[1]/td[5]').text)
            shijinglv.append(driver.find_element_by_xpath('/html/body/div[13]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[1]/td[6]').text)
            maolilv.append(driver.find_element_by_xpath('/html/body/div[13]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[1]/td[7]').text)
            roe.append(driver.find_element_by_xpath('/html/body/div[13]/div[2]/div[2]/div[2]/div[4]/table/tbody/tr[1]/td[9]').text)

        jiancheng.insert(0, '股票简称')
        hangye.insert(0, '行业')
        shangshishijian.insert(0, '上市时间')

        shizhi.insert(0, '总市值')
        jingzichan.insert(0, '净资产')
        jinglirun.insert(0, '净利润')
        shiyinglv.insert(0, '市盈率')
        shijinglv.insert(0, '市净率')
        maolilv.insert(0, '毛利率')
        roe.insert(0, 'ROE')
        #print(shizhi)
        xw.Range('B1').options(transpose=True).value = jiancheng
        xw.Range('C1').options(transpose=True).value = hangye
        xw.Range('D1').options(transpose=True).value = shangshishijian

        xw.Range('E1').options(transpose=True).value = shizhi
        xw.Range('F1').options(transpose=True).value = jingzichan
        xw.Range('G1').options(transpose=True).value = jinglirun
        xw.Range('H1').options(transpose=True).value = shiyinglv
        xw.Range('I1').options(transpose=True).value = shijinglv
        xw.Range('J1').options(transpose=True).value = maolilv
        xw.Range('K1').options(transpose=True).value = roe
    finally:
        driver.quit()

def getHuobizijin3Items():
    driver = webdriver.Chrome()
    try:
        sz_url = 'http://f10.eastmoney.com/f10_v2/FinanceAnalysis.aspx?code=sz'
        sh_url = 'http://f10.eastmoney.com/f10_v2/FinanceAnalysis.aspx?code=sh'

        huobizijin = []
        jingzichanshouyilv =[]
        touzishouyi = []
        caibaoleixing = []

        codes = []
        for i in xw.Range('A2:A1000'):
            if i.value == None:
                break
            codes.append(i.value)

        for code in codes:
            if int(code)>600000:
                driver.get(sh_url+code)
            else:
                driver.get(sz_url+code)

            huobizijin.append(driver.find_element_by_xpath('/html/body/div[1]/div[13]/div[4]/div[6]/p[4]').text)
            jingzichanshouyilv.append(driver.find_element_by_xpath('/html/body/div[1]/div[13]/div[4]/div[1]/p').text)
            touzishouyi.append(driver.find_element_by_xpath('/html/body/div[1]/div[13]/div[4]/div[9]/p[1]').text)
            caibaoleixing.append('按报告期：'+driver.find_element_by_xpath('//*[@id="DBFX_ul"]/li[1]').text)

        huobizijin.insert(0, '货币资金')
        jingzichanshouyilv.insert(0, '净资产收益率')
        touzishouyi.insert(0, '投资收益')
        caibaoleixing.insert(0, '财报类型')
        #print(shizhi)
        xw.Range('L1').options(transpose=True).value = huobizijin
        xw.Range('M1').options(transpose=True).value = jingzichanshouyilv
        xw.Range('N1').options(transpose=True).value = touzishouyi
        xw.Range('S1').options(transpose=True).value = caibaoleixing

    finally:
        driver.quit()

def getGaoguan():
    driver = webdriver.Chrome()
    try:
        sz_url = 'http://f10.eastmoney.com/f10_v2/CompanyManagement.aspx?code=sz'
        sh_url = 'http://f10.eastmoney.com/f10_v2/CompanyManagement.aspx?code=sh'

        gaoguan = []

        codes = []
        for i in xw.Range('A2:A1000'):
            if i.value == None:
                break
            codes.append(i.value)

        for code in codes:
            if int(code)>600000:
                driver.get(sh_url+code)
            else:
                driver.get(sz_url+code)

            table = driver.find_element_by_xpath('/html/body/div[1]/div[12]/div[2]').text
            tmp = []
            for i in table.split('\n')[1:]:
                name = i.split(' ')[1]
                position = ''.join(i.split(' ')[5:])
                itm = name + ' [' + position + '] '
                if itm == '姓名 [职务] ':
                    continue
                tmp.append(itm)
            item = '、'.join(tmp)

            gaoguan.append(item)

        gaoguan.insert(0, '公司高管')
        print(gaoguan)
        xw.Range('O1').options(transpose=True).value = gaoguan

    finally:
        driver.quit()
        #input()

def get10Gudong():
    driver = webdriver.Chrome()
    try:
        sz_url = 'http://f10.eastmoney.com/f10_v2/ShareholderResearch.aspx?code=sz'
        sh_url = 'http://f10.eastmoney.com/f10_v2/ShareholderResearch.aspx?code=sh'

        gudong = []

        codes = []
        eachGudong = []
        for i in xw.Range('A2:A1000'):
            if i.value == None:
                break
            codes.append(i.value)

        for code in codes:
            if int(code)>600000:
                driver.get(sh_url+code)
            else:
                driver.get(sz_url+code)

            soup = bs(driver.page_source, 'lxml')

            if str(pandas.read_html(str(soup.findAll('table')[1]))[0][2][0]) == '股东性质':
                table = soup.findAll('table')[7]
                print(pandas.read_html(str(soup.findAll('table'))))
                tab = pandas.read_html(str(table))
                #print('if', tab)
                gongsi = list(tab[0][1][1:-1])
                bizhong = list(tab[0][4][1:-1])
                for i, j in enumerate(gongsi):
                    eachGudong.append(j + '（' + bizhong[i] + '）')
                gudong.append('、'.join(eachGudong))
                #print(driver.find_element_by_xpath('/html/body/div[1]/div[13]/div[2]'))
            else:
                table = soup.findAll('table')[1]
                tab = pandas.read_html(str(table))
                print('else', tab)
                gongsi = list(tab[0][1][1:-1])
                bizhong = list(tab[0][4][1:-1])
                for i, j in enumerate(gongsi):
                    eachGudong.append(j + '（' + bizhong[i] + '）')
                gudong.append('、'.join(eachGudong))

        gudong.insert(0, '10大股东')
        print(gudong)
        xw.Range('P1').options(transpose=True).value = gudong

    finally:
        #input()
        driver.quit()


get10Gudong()