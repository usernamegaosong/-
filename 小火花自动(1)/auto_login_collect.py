import re
import shutil
import win32clipboard as wc

import requests
import xlwt
from PIL import Image
from lxml import etree
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from YDMHTTPDemo3 import YDM
from update_cookie import *
from tools import *


class AutoHuo(object):
    def __init__(self, style, start_time, end_time):
        self.option = webdriver.ChromeOptions()
        # self.option.add_argument('blink-settings=imagesEnabled=false')  # 不加载图片, 提升速度
        self.option.add_argument('--no-sandbox')  # 解决DevToolsActivePort文件不存在的报错
        self.option.add_argument('--disable-gpu')  # 谷歌文档提到需要加上这个属性来规避bug
        # self.option.add_argument('--headless')
        self.browser = webdriver.Chrome(chrome_options=self.option)
        self.browser.maximize_window()
        self.index_url = 'https://baowen.xhh.com/manage/toutiao/gallery'
        self.choosed_style = style
        self.start_time = start_time
        self.end_time = end_time

    def index(self):
        self.browser.delete_all_cookies()
        base_path = os.getcwd()
        with open(base_path + '\cookie.txt', 'r', encoding='utf8')as fp:
            cookie_list = eval(fp.read())
        self.browser.get(url=self.index_url)
        for cookie in cookie_list:
            self.browser.add_cookie({
                'domain': cookie['domain'],
                'name': cookie['name'],
                'value': cookie['value'],
                'httpOnly': cookie['httpOnly'],
                'path': cookie['path'],
                'secure': cookie['secure'],
            })
        self.browser.get(url=self.index_url)
        if self.browser.current_url.startswith('https://baowen.xhh.com/validating'):
            WebDriverWait(self.browser, 2).until(EC.visibility_of(self.browser.find_element_by_class_name('db')))
            self.browser.get_screenshot_as_file('captcha.png')
            img_element = self.browser.find_element_by_xpath('//form[@class="mx-auto"]/div[1]/a/img')
            left = img_element.location['x']
            top = img_element.location['y']
            right = img_element.location['x'] + img_element.size['width']
            bottom = img_element.location['y'] + img_element.size['height']
            im = Image.open('captcha.png')
            im = im.crop((left, top, right, bottom))
            im.save('captcha.png')
            ydm = YDM()
            file_path = os.path.join(os.getcwd(), 'captcha.png')
            cid, result = ydm.run(file_path)
            if int(cid) > 0:
                img_result = result
                input_img_element = self.browser.find_element_by_xpath(
                    '//div[@class="clearfix mb1em"]/input[@name="captcha"]')
                input_img_element.clear()
                input_img_element.send_keys(img_result)
                self.browser.find_element_by_xpath('//button[@class="btn btn-red px4em mt1em"]').click()
                shutil.rmtree(file_path)
            else:
                print('验证码获取失败')
                os.remove(file_path)
        time.sleep(3)

    def choose_style(self):
        style_list = self.browser.find_elements_by_xpath('//select[@name="cat_id"]/option')
        for style_temp in style_list:
            style = style_temp.text
            if style == self.choosed_style:
                style_temp.click()

    def choose_time(self):
        time.sleep(1)
        start_time = self.browser.find_element_by_xpath(
            '//div[@class="input-group input-group-inner"]/input[@name="start"]')
        start_time.clear()
        start_time.send_keys(self.start_time)
        time.sleep(0.5)
        end_time = self.browser.find_element_by_xpath(
            '//div[@class="input-group input-group-inner"]/input[@name="end"]')
        end_time.clear()
        end_time.send_keys(self.end_time)
        time.sleep(0.5)
        # 点击搜索
        self.browser.find_element_by_xpath('//button[@class="btn btn-default btn-search mr2px"]').click()
        time.sleep(2)

    def after_search(self):
        self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        # 总数据
        try:
            all_num_element = self.browser.find_element_by_xpath(
                '//div[@class="clearfix center bg-white pb3em"]/ul[@class="pagination pagination-sm"]/li[last()-1]//u')
            print('所选时间段图集总数：' + str(all_num_element.text) + '条')
            return True
        except Exception as e:
            print('无数据')
            return False

    # 获取粘贴板里的内容并清洗
    def get_copy_txet(self):
        # 使用前先打开
        wc.OpenClipboard()
        # 获取内容
        copytxet = wc.GetClipboardData()
        # 清空粘贴板
        wc.EmptyClipboard()
        # 关闭粘贴板
        wc.CloseClipboard()
        base_url_first = str(copytxet).split('\r\n')
        base_url_second = [i for i in base_url_first if len(i) > 1]
        return base_url_second

    def get_url(self):
        # 总页码数
        all_page_element = self.browser.find_element_by_xpath('//div[@class="clearfix center bg-white pb3em"]/ul[@class="pagination pagination-sm"]/li[last()-2]//a')
        all_page = int(all_page_element.text)
        data_list = []
        for i in range(1, all_page+1):
            if self.browser.current_url.startswith('https://baowen.xhh.com/validating'):
                WebDriverWait(self.browser, 2).until(EC.visibility_of(self.browser.find_element_by_class_name('db')))
                self.browser.get_screenshot_as_file('captcha.png')
                img_element = self.browser.find_element_by_xpath('//form[@class="mx-auto"]/div[1]/a/img')
                left = img_element.location['x']
                top = img_element.location['y']
                right = img_element.location['x'] + img_element.size['width']
                bottom = img_element.location['y'] + img_element.size['height']
                im = Image.open('captcha.png')
                im = im.crop((left, top, right, bottom))
                im.save('captcha.png')
                ydm = YDM()
                file_path = os.path.join(os.getcwd(), 'captcha.png')
                cid, result = ydm.run(file_path)
                if int(cid) > 0:
                    img_result = result
                    input_img_element = self.browser.find_element_by_xpath('//div[@class="clearfix mb1em"]/input[@name="captcha"]')
                    input_img_element.clear()
                    input_img_element.send_keys(img_result)
                    self.browser.find_element_by_xpath('//button[@class="btn btn-red px4em mt1em"]').click()
                    os.remove(file_path)
                else:
                    print('验证码获取失败')
                    os.remove(file_path)
            WebDriverWait(self.browser, 2).until(EC.visibility_of(self.browser.find_element_by_class_name('card-box')))
            self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            # 全选
            WebDriverWait(self.browser, 2).until(EC.visibility_of(self.browser.find_element_by_class_name('toolsbox')))
            # all_choose = self.browser.find_element_by_xpath('//div[@class="toolsbox"]/button[position()=2]')
            # all_choose.click()
            # time.sleep(0.5)
            # 再次确认
            # all_choose.click()
            # 复制
            # time.sleep(0.5)
            # copy_url = self.browser.find_element_by_xpath('//div[@class="toolsbox"]/button[position()=3]')
            # copy_url.click()
            # 第二次复制
            # time.sleep(0.5)
            # copy_url_second = self.browser.find_element_by_xpath('//div[@class="modal-footer center"]/button[@id="btn-copy"]')
            # copy_url_second.click()
            # 获取粘贴板内容
            # finally_url_list = self.get_copy_txet()

            time.sleep(0.5)
            # self.browser.find_element_by_xpath('//div[@class="modal-header"]/button[@class="close"]').click()
            # 获取title
            page_source = self.browser.page_source
            html = etree.HTML(page_source)
            title_list_element = html.xpath('//tbody/tr')
            rule_title = re.compile(r'\s*')
            rule_url = re.compile(r'\s*')
            for element in title_list_element:
                url = ''.join(element.xpath('./td[position()=1]/a[last()]/@href'))
                url = rule_url.sub('', url)
                title = ''.join(element.xpath('./td[position()=1]/a[last()]/text()'))
                title = rule_title.sub('', title)
                if len(title) and len(url) > 1:
                    data = []
                    data.append(title)
                    data.append(style)
                    data.append(url)
                    data_list.append(data)
                else:
                    continue
            # 下一页
            try:
                next_url_element = self.browser.find_element_by_xpath('//div[@class="clearfix center bg-white pb3em"]/ul[@class="pagination pagination-sm"]/li[last()]/a')
                if next_url_element.is_enabled():
                    next_url = next_url_element.get_attribute('href')
                    self.browser.get(next_url)
                else:
                    print('已翻致最后以一页')
                    break
            except Exception as e:
                print('已翻至最后以一页')
                break
        # 保存数据
        self.write_xls(self.start_time, self.end_time, self.choosed_style, data_list)

    def write_xls(self, start_time, end_time, style, data_list):
        end_time_rule = re.compile(r'-(\d+.*)\s+')
        xls_name = style + start_time.split(' ')[0] + ' ' + end_time_rule.findall(end_time)[0] + '.xls'
        workbook = xlwt.Workbook(encoding='ascii')
        worksheet = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
        for i in range(len(data_list)):
            for j in [0, 3, 5]:
                num = [0, 3, 5].index(j)
                worksheet.write(i, j, data_list[i][num])
        workbook.save(xls_name)

    def run(self):
        self.index()
        self.choose_style()
        self.choose_time()
        if self.after_search():
            self.get_url()
            self.browser.quit()
        else:
            print('更换时间')
            self.browser.quit()


if __name__ == "__main__":
    print('*' * 10 + '小火花采集器启动中' + '*' * 10)
    index = Index()
    start = 371
    for i in range(start + 1):
        print(index(i, start), end='')
        time.sleep(0.01)
    style = input('\n请输入分类：')
    start_time = input('请输入开始时间：')
    while True:
        end_time = input('请输入结束时间：')
        time_struct = time.strptime(end_time, '%Y-%m-%d %H：%M')
        time_stamp = time.mktime(time_struct)
        if time_stamp > time.time():
            print('输入的结束时间违法，重新输入')
        else:
            break
    auto = AutoHuo(style, start_time, end_time)
    auto.run()
