import re
import time
import xlwt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

chrome_options = Options()
chrome_options.add_argument("--headless")
driver = webdriver.Chrome(chrome_options=chrome_options)


def LoginWeibo(username, password):
    try:
        print('准备登陆 Weibo...')
        driver.get("http://login.sina.com.cn/")
        elem_user = driver.find_element_by_name("username")
        elem_user.send_keys(username)
        elem_pwd = driver.find_element_by_name("password")
        elem_pwd.send_keys(password)
        elem_sub = driver.find_element_by_xpath("//*[@id='vForm']/div[2]/div/ul/li[7]/div[1]/input")
        elem_sub.click()

        try:
            time.sleep(10)
            elem_sub.click()
        except:
            pass

        print('Crawl in ', driver.current_url)
        print('输出Cookie键值对信息:')
        for cookie in driver.get_cookies():
            print(cookie)
            for key in cookie:
                print(key, cookie[key])
        print('登陆成功...')
    except Exception as e:
        print("Error: ", e)
    finally:
        print('End LoginWeibo!')


def GetSearchContent(key):
    driver.get("http://s.weibo.com/")

    item_inp = driver.find_element_by_xpath("//*[@id='pl_homepage_search']/div/div[2]/div/input")
    item_inp.send_keys(key)
    item_inp.send_keys(Keys.RETURN)

    global outfile
    global sheet

    outfile = xlwt.Workbook(encoding='utf-8')
    sheet = outfile.add_sheet(key)
    initXLS()
    counts = re.findall(r'(?<=找到)\d+\.?\d*', driver.find_element_by_xpath("//*[@id='pl_feedlist_index']/div[3]").text)[
        0]
    pageCounts = 20
    page = int((int(counts) + pageCounts - 1) / pageCounts)
    print('总共发现', page, '页')
    current_url = driver.current_url
    getContent()

    for i in range(page - 1):
        driver.get(current_url + '&page=' + str(i + 2))
        print("当前爬取第", i + 2, "页")
        getContent()
        time.sleep(1)


def initXLS():
    name = ['博主昵称', '博主主页', '微博认证', '微博达人', '微博内容', '发布时间', '微博地址', '微博来源', '转发', '评论', '赞']

    global row
    global outfile
    global sheet

    row = 0
    for i in range(len(name)):
        sheet.write(row, i, name[i])
    row = row + 1
    outfile.save("./crawl_output_YS.xls")


def writeXLS(dic):
    global row
    global outfile
    global sheet

    for k in dic:
        for i in range(len(dic[k])):
            sheet.write(row, i, dic[k][i])
        row = row + 1
    outfile.save("./crawl_output_YS.xls")


def getContent():
    nodes = driver.find_elements_by_xpath("//div[@class='card']")
    print(nodes)

    if len(nodes) == 0:
        print('error')

    dic = {}

    print('微博数量', len(nodes))
    js = """
         elements = document.getElementsByClassName('txt');
         for (var i=0; i < elements.length; i++)
         {
             elements[i].style.display='block'
         }
         """

    driver.execute_script(js)

    for i in range(len(nodes)):
        dic[i] = []

        try:
            BZNC = nodes[i].find_element_by_xpath(".//div[@class='feed_content wbcon']/a[@class='W_texta W_fb']").text
        except:
            BZNC = ''
        print('博主昵称:', BZNC)
        dic[i].append(BZNC)

        try:
            BZZY = nodes[i].find_element_by_xpath(
                ".//div[@class='feed_content wbcon']/a[@class='W_texta W_fb']").get_attribute("href")
        except:
            BZZY = ''
        print('博主主页:', BZZY)
        dic[i].append(BZZY)

        try:
            WBRZ = nodes[i].find_element_by_xpath(
                ".//div[@class='feed_content wbcon']/a[@class='approve_co']").get_attribute('title')
        except:
            WBRZ = ''
        print('微博认证:', WBRZ)
        dic[i].append(WBRZ)

        try:
            WBDR = nodes[i].find_element_by_xpath(
                ".//div[@class='feed_content wbcon']/a[@class='ico_club']").get_attribute('title')
        except:
            WBDR = ''
        print('微博达人:', WBDR)
        dic[i].append(WBDR)

        try:
            WBNR = nodes[i].find_element_by_xpath(".//div[@class='content']/p[@class='txt'][2]").text
        except:
            WBNR = nodes[i].find_element_by_xpath(".//div[@class='content']/p[@class='txt'][1]").text
        print('微博内容:', WBNR)
        dic[i].append(WBNR)

        try:
            source = nodes[i].find_element_by_xpath(".//div[@class='content']/p[@class='from']").text
        except:
            source = ''

        try:
            device = re.findall(r"(?<=来自 )(.*)", source)[0]
        except:
            device = ''

        try:
            sendtime = re.findall(r"(.+?) 来自", source)[0]
        except:
            sendtime = ''

        print('发布时间:', sendtime)
        dic[i].append(sendtime)

        try:
            WBDZ = nodes[i].find_element_by_xpath(
                ".//div[@class='feed_from W_textb']/a[@class='W_textb']").get_attribute("href")
        except:
            WBDZ = ''
        print('微博地址:', WBDZ)
        dic[i].append(WBDZ)

        try:
            WBLY = nodes[i].find_element_by_xpath(".//div[@class='feed_from W_textb']/a[@rel]").text
        except:
            WBLY = ''
        print('微博来源:', device)
        dic[i].append(device)

        try:
            ZF_TEXT = nodes[i].find_element_by_xpath(".//a[@action-type='feed_list_forward']//em").text
            if ZF_TEXT == '':
                ZF = 0
            else:
                ZF = int(ZF_TEXT)
        except:
            ZF = 0
        print('转发:', ZF)
        dic[i].append(str(ZF))

        try:
            PL_TEXT = nodes[i].find_element_by_xpath(".//a[@action-type='feed_list_comment']//em").text
            if PL_TEXT == '':
                PL = 0
            else:
                PL = int(PL_TEXT)
        except:
            PL = 0
        print('评论:', PL)
        dic[i].append(str(PL))

        try:
            ZAN_TEXT = nodes[i].find_element_by_xpath(".//a[@action-type='feed_list_like']//em").text
            if ZAN_TEXT == '':
                ZAN = 0
            else:
                ZAN = int(ZAN_TEXT)
        except:
            ZAN = 0
        print('赞:', ZAN)
        dic[i].append(str(ZAN))

    writeXLS(dic)


if __name__ == '__main__':
    username = '***'
    password = '***'

    LoginWeibo(username, password)

    key = '短缺药'
    GetSearchContent(key)
