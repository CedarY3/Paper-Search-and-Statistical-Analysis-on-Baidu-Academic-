import json
import re
import time
import sys
import os
import requests
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium import webdriver


from PyQt5.QtWidgets import QWidget, QApplication, QFileDialog
from PyQt5 import QtCore, uic

browserType = webdriver.Chrome

import openpyxl


# chrome_driver_path = "./"
# service = Service(executable_path=chrome_driver_path)
# browserType = webdriver.Chrome(service=service)

class MyWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.init_ui()
        self.csvf = []

    def init_ui(self):
        self.ui = uic.loadUi("search_paper.ui")

        print(self.ui.__dict__.keys())  # 查看ui文件中有哪些控件

        cwd = os.getcwd()

        self.ui.Save_Path.setText(cwd)
        self.ui.Driver_Path.setText(cwd)
        # self.ui.LEdit_SavePath.setText(r"D:\桌面")

        # 绑定槽函数--------------

        # 选择文件路径
        self.ui.Select_Driver.clicked.connect(lambda: self.click_set_path(1))
        # self.ui.PButton_SelectOfferFile.clicked.connect(self.click_find_file_path)
        self.ui.Select_Save.clicked.connect(lambda: self.click_set_path(2))

        # # 登录
        # self.ui.PButton_Login.clicked.connect(self.click_login)

        # # 开始
        self.ui.Start.clicked.connect(self.click_start)

        # # 结束
        self.ui.End.clicked.connect(self.click_end)

        # 调整消息框的scrollbar的槽函数
        # self.ui.Scroll_Area.verticalScrollBar().rangeChanged.connect(self.set_scroll_bar)

    # 更新系统消息的函数
    def updatemsg(self, news):
        self.ui.msg.resize(591, self.ui.msg.frameSize().height() + 20)
        # self.ui.Scroll_Area.setMinimumHeight(self.ui.msg.frameSize().height() + 60)
        self.ui.msg.setText(self.ui.msg.text() + "<br>" + news)
        self.ui.msg.repaint()  # 更新内容，如果不更新可能没有显示新内容
        print(news)

    # 调整消息框的scrollbar的槽函数
    # def set_scroll_bar(self):
    #     self.ui.Scroll_Area.verticalScrollBar().setValue(self.ui.Scroll_Area.verticalScrollBar().maximum())

    # 选择保存路径的槽函数
    def click_set_path(self, flag):
        m = QFileDialog.getExistingDirectory(None, "选取文件夹", "./")  # 起始路径
        if m != "":
            if flag == 1:
                self.ui.Driver_Path.setText(m)
            else:
                self.ui.Save_Path.setText(m)

    # 选择文件的槽函数
    def click_find_file_path(self):
        # 设置文件扩展名过滤，同一个类型的不同格式如xlsx和xls 用空格隔开
        filename, filetype = QFileDialog.getOpenFileName(self, "选择文件", "./", "*.xlsx")
        if filename == "":
            return
        self.ui.LEdit_OfferFilePath.setText(filename)

    # 结束，槽函数
    def click_end(self):
        # todo
        print('++++++++++还没做好结束功能+++++++++')

    # 主函数，槽函数
    def click_start(self):
        self.updatemsg("+++开始寻找水源！+++")
        key_words = self.ui.Keywords.text()
        start_time = self.ui.Start_Time.text()
        end_time = self.ui.End_Time.text()
        if not key_words:
            self.updatemsg('请输入您要搜索的关键词')
            return
        if not start_time:
            self.updatemsg('您没有选择论文投稿时间段的开始时间，将默认设为2020年\n')
            start_time = '2020'
        if not end_time:
            end_time = str(time.localtime().tm_year)
            self.updatemsg(f'您没有选择论文投稿时间段的截止时间，将默认设为现在:{end_time}\n')


        save_path = self.ui.Save_Path.text() + f'\{key_words}.txt'
        driver_path = self.ui.Driver_Path.text() + f'\chromedriver.exe'

        self.driver_path = driver_path
        # 将driver_path参数保存在实例的chrome_driver_path属性中。

        self.browser: browserType
        # 通过Type Hinting将browser属性定义为browserType类型。

        self.Init_Browser()
        # 调用成员函数_init_browser()来初始化浏览器属性。

        # print(driver_path)
        # print(key_words)
        # print(start_time)
        # print(save_path)

        page_num = 99

        time.sleep(1)

        code = self.run(key_words, page_num, start_time, end_time, save_path)

        if code == 0:
            self.updatemsg('成功获取所有期刊~')
        else:
            self.updatemsg('error!')

# class BaiduXueshuAutomatic:
#
#     def __init__(self, chrome_driver_path: str) -> None:  # 定义了一个构造函数，用来初始化对象的状态，接收一个chrome_driver_path参数。
#

    def Init_Browser(self) -> None:  # 定义了一个私有成员函数，用来初始化浏览器属性。

        options = Options()  # 创建一个Options对象，用来设置ChromeDriver的选项。

        options.add_argument('--headless')  # 向Options对象加入'--headless'参数，表示使用无头（Headless）模式运行ChromeDriver，即没有图形界面显示。

        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        # 某些网站会检测到使用了Selenium来启动浏览器，为了规避这种检测，用这个选项就可以在运行时抹掉Google自动化控制的特征，一定程度上避免被检测到。

        # self.browser = webdriver.Chrome(executable_path=self.driver_path, options=options)
        # 创建一个ChromeDriver对象，将ChromeDriver的可执行路径和Options对象作为参数传入，以此来创建一个浏览器对象。这个浏览器对象将被保存在实例的browser属性中。
        self.browser = webdriver.Chrome(service=Service(self.driver_path), options=options)

        self.browser.implicitly_wait(3)
        # 等待页面加载，最长等待时间为3秒，如果到了3秒页面还没加载完成，Selenium会抛出异常。

        self.wait = WebDriverWait(self.browser, 10)
        # 定义一个WebDriverWait对象，用于等待页面元素加载完成，最长等待时间为10秒。这个WebDriverWait对象将被保存在实例的wait属性中。

        self.ac = ActionChains(self.browser)  # 定义一个ActionChains对象，用于实现鼠标悬停、拖拽等操作，将这个ActionChains对象保存在ac属性中。

    def _wait_by_xpath(self, patten):  # 定义了一个名为_wait_by_xpath的成员函数，接收一个参数patten（XPath路径）。
        self.wait.until(EC.presence_of_element_located((By.XPATH, patten)))
        # 使用WebDriverWait对象等待指定的XPath元素出现，如果元素没出现，就一直等待，直到超时。如果超时还未找到该元素，则会抛出异常。
        # 这里的EC是一个Selenium内置的ExpectedConditions类，表示期待的条件，
        # .presence_of_element_located()是期待条件之一，用来等待元素出现。
        # (By.XPATH, patten)表示定义了一个元组，包含一个表示XPath定位方式的字符串和一个表示XPath路径的变量patten。

    def is_contain_chinese(self, check_str):
        for ch in check_str:
            if u'\u4e00' <= ch <= u'\u9fff':
                return True
        return False

    def run(self, wd, page_num=999, startyear=2020, endyear=2023, fpath='tes.txt'):

        headers = {
            'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
        }
        # 定义了一个字典类型的变量headers，
        # 它们被用来在HTTP请求中添加一个请求头，以便Web服务器能够识别请求的来源和类型。
        # 具体来说，此处设置的'user-agent'请求头是用来标识用户使用的浏览器类型和版本。
        # 这个headers变量可以在使用Python中的requests库向Web服务器发送HTTP请求时使用，以帮助模拟一个正常的用户请求。

        first_paper = ''
        count = 0
        paper_count = 0
        con_count = 0
        chinese_count = 0
        english_count = 0

        # try:
        while count <= page_num:
            url = f'https://xueshu.baidu.com/s?wd=intitle%3A%28"{wd}"%29&pn={count}0&tn=SE_baiduxueshu_c1gjeupa&ie=utf-8&filter=sc_year%3D%7B{startyear}%2C{endyear}%7D&sc_f_para=sc_tasktype%3D%7BfirstSimpleSearch%7D&bcp=2&sc_hit=1'

            self.browser.get(url)

            print(f'start url: {url}\n')
            time.sleep(1)

            # 课程学习内部
            # first_urls = self.browser.find_elements_by_xpath(
            first_urls = self.browser.find_elements(by=By.XPATH, value=
            '/html/body/div[1]/div[4]/div[3]/div[2]/div/div[@class="result sc_default_result xpath-log"]')

            # find_elements的参数有两个：定位器的 by 和 value。
            # by：定位器的类型。它可以是以下值之一：     # By.ID：根据元素的 id 属性进行定位。
            # By.NAME：根据元素的 name 属性进行定位。
            # By.XPATH：根据元素的 XPath 表达式进行定位。
            # By.CSS_SELECTOR：根据元素的 SS 选择器进行定位。
            # By.CLASS_NAME：根据元素的 class 属性进行定位。
            # By.TAG_NAME：根据元素的标签名进行定位。
            # By.LINK_TEXT：根据元素的链接文本进行定位。
            # By.PARTIAL_LINK_TEXT：根据元素的链接部分文本进行定位。

            # value：定位器的值，即查找元素所用的关键字。它的值根据 by 的不同而有所不同。
            # 例如，使用 By.ID 定位器时，value应包含要查找的元素的 id值（如：value = "element_id"）。
            # 使用 By.XPATH 定位器时，value 应包含要查找的XPath表达式（如：value = "//div[@class='class_name']"）。
            # 在查找元素时，需要根据所需定位的元素的属性或特征选择合适的定位器类型和值。

            for ii, first_url in enumerate(first_urls):

                if count == 0 and ii == 0:
                    all_dic = {'all_paper_num': 0,
                               'English Journal': {},
                               'Chinese Journal': {},
                               'Conference': {},
                               }

                    with open(fpath, 'w', encoding='utf8') as fw:
                        fw.write(json.dumps(all_dic, ensure_ascii=False))
                else:
                    with open(fpath, 'r', encoding='utf8') as fr:
                        all_dic = json.loads(fr.read())

                # paper_name = first_url.find_element_by_xpath('div[1]/h3/a')
                paper_name = first_url.find_element(by=By.XPATH, value='div[1]/h3/a')
                paper_link = paper_name.get_attribute("href")
                paper_name = paper_name.text

                if paper_link == first_paper:
                    count = 9999
                    break

                if ii == 0:
                    first_paper = paper_link

                res1 = requests.get(paper_link, headers=headers).text

                journal_name = re.findall('<a class="journal_title".*?>(.*?)</a>', res1, re.S)
                if not journal_name:
                    continue
                journal_name = journal_name[0]

                paper_count += 1

                self.updatemsg(f'第{paper_count}篇论文')

                if ': ' in journal_name:
                    journal_name = journal_name.split(': ')[0]

                if '&amp;' in journal_name:
                    journal_name = journal_name.replace('&amp;', '&')

                if '&#039;' in journal_name:
                    journal_name = journal_name.replace('&#039;', "'")



                # 中文期刊
                if self.is_contain_chinese(journal_name):
                    self.updatemsg("Chinese journal don't check.")

                    journal_name = f'{journal_name}'
                    self.updatemsg(paper_name)
                    self.updatemsg(paper_link)
                    self.updatemsg(journal_name)
                    self.updatemsg('\n')

                    if not all_dic['Chinese Journal'].get(journal_name):
                        all_dic['Chinese Journal'][journal_name] = []

                    all_dic['Chinese Journal'][journal_name].append({paper_name: paper_link})

                    chinese_count += 1

                    all_dic['all_paper_num'] = {
                        'all_count': paper_count,
                        'English Journal count': english_count,
                        'Chinese Journal count': chinese_count,
                        'Conference count': con_count
                    }

                    for i in list(all_dic['Chinese Journal'].keys()):
                        if '共' in all_dic['Chinese Journal'][i][0]:
                            all_dic['Chinese Journal'][i][0] = f'共{len(all_dic["Chinese Journal"][i]) - 1}篇'
                        else:
                            all_dic['Chinese Journal'][i].insert(0, f"共1篇")

                    all_dic['Chinese Journal'] = {key: all_dic['Chinese Journal'][key] for key, value in
                                                  sorted(all_dic['Chinese Journal'].items(), key=lambda x: x[1][0],
                                                         reverse=True)}

                    all_dic = json.dumps(all_dic, ensure_ascii=False)
                    with open(fpath, 'w', encoding='utf8') as fw:
                        fw.write(all_dic)
                        fw.flush()

                    time.sleep(2)
                    continue

                # 会议期刊
                elif 'Conference' in journal_name:
                    self.updatemsg("Conference don't check.")

                    journal_name = f'{journal_name}'
                    self.updatemsg(paper_name)
                    self.updatemsg(paper_link)
                    self.updatemsg(journal_name)
                    self.updatemsg('\n')

                    if not all_dic['Conference'].get(journal_name):
                        all_dic['Conference'][journal_name] = []

                    all_dic['Conference'][journal_name].append({paper_name: paper_link})

                    con_count += 1

                    all_dic['all_paper_num'] = {
                        'all_count': paper_count,
                        'English Journal count': english_count,
                        'Chinese Journal count': chinese_count,
                        'Conference count': con_count
                    }

                    for i in list(all_dic['Conference'].keys()):
                        if '共' in all_dic['Conference'][i][0]:
                            all_dic['Conference'][i][0] = f'共{len(all_dic["Conference"][i]) - 1}篇'
                        else:
                            all_dic['Conference'][i].insert(0, f"共1篇")

                    all_dic['Conference'] = {key: all_dic['Conference'][key] for key, value in
                                             sorted(all_dic['Conference'].items(), key=lambda x: x[1][0], reverse=True)}

                    all_dic = json.dumps(all_dic, ensure_ascii=False)
                    with open(fpath, 'w', encoding='utf8') as fw:
                        fw.write(all_dic)
                        fw.flush()

                    time.sleep(2)
                    continue

                # 英文非会议期刊
                else:
                    status_code = True
                    cite_score = ''
                    journal_split = ''
                    journal_date = ''

                    while status_code:
                        post_dic = {'searchname': journal_name, 'searchsort': 'relevance'}
                        search_url = 'http://www.letpub.com.cn/index.php?page=journalapp&view=search'
                        res = requests.post(search_url, post_dic, headers=headers)
                        if res.status_code == 200:
                            status_code = False
                        else:

                            self.updatemsg(f'letpub denied, {journal_name}')
                            time.sleep(5)
                            continue

                        second_search = re.findall('</style>.*?<tr>(.*?)</tr>', res.text, re.S)[0]
                        cite_score = re.findall('CiteScore:(\d+.\d+)', second_search, re.S)
                        journal_split = re.findall('(\d区)</td>', second_search, re.S)

                        journal_date = ''
                        fourth_search = re.findall('(<td.*?</td>)', second_search, re.S)
                        for i in fourth_search:
                            l = ['月', '周', 'eeks']
                            for j in l:
                                if j in i:
                                    journal_date = re.findall('>(.*?)</td>', i, re.S)
                                    break

                    if cite_score and journal_split and journal_date:
                        journal_name = f'{journal_split[0]} Citescore:{cite_score[0]} 审稿周期:{journal_date[0]} {journal_name}'
                    elif cite_score and journal_split:
                        journal_name = f'{journal_split[0]} Citescore:{cite_score[0]} 审稿周期:无记录 {journal_name}'
                    elif cite_score:
                        journal_name = f'未收录 Citescore:{cite_score[0]} 审稿周期:无记录 {journal_name}'
                    else:
                        journal_name = f'未收录 无Citescore 审稿周期:无记录 {journal_name}'

                    self.updatemsg(paper_name)
                    self.updatemsg(paper_link)
                    self.updatemsg(journal_name)
                    self.updatemsg('\n')
                    self.csvf.append(f'{paper_name}, {paper_link}, {journal_name}')

                    if not all_dic['English Journal'].get(journal_name):
                        all_dic['English Journal'][journal_name] = []

                    all_dic['English Journal'][journal_name].append({paper_name: paper_link})

                    english_count += 1

                    all_dic['all_paper_num'] = {
                        'all_count': paper_count,
                        'English Journal count': english_count,
                        'Chinese Journal count': chinese_count,
                        'Conference count': con_count
                    }

                    all_dic['English Journal'] = {key: all_dic['English Journal'][key] for key in
                                                  sorted(all_dic['English Journal'].keys())}

                    for i in list(all_dic['English Journal'].keys()):
                        if '共' in all_dic['English Journal'][i][0]:
                            all_dic['English Journal'][i][0] = f'共{len(all_dic["English Journal"][i]) - 1}篇'
                        else:
                            all_dic['English Journal'][i].insert(0, f"共1篇")

                    all_dic = json.dumps(all_dic, ensure_ascii=False)
                    with open(fpath, 'w', encoding='utf8') as fw:
                        fw.write(all_dic)
                        fw.flush()
                    with open(fpath+'.csv', 'w', encoding='utf8') as fw:
                        fw.write('paper_name, paper_link, journal_name\n')
                        for i in self.csvf:
                            fw.write(i+'\n')

                    time.sleep(5)

            # 成功读取一个网页的内容
            count += 1

        # 成功获取所要求数量的网页的全部内容
        print('成功获取所有期刊~')
        return 0


if __name__ == '__main__':
    app = QApplication(sys.argv)

    w = MyWindow()
    # 展示窗口
    w.ui.show()

    app.exec()


