import json
import os
import time

from selenium import webdriver


class AutoCookie(object):
    def __init__(self):
        self.browser = webdriver.Chrome()
        self.browser.maximize_window()
        self.login_url = 'https://user.xhh.com/duke/signin?url=https%3A%2F%2Fbaowen.xhh.com%2Fmanage%2Ftoutiao%2Fgallery'
        # self.login_url = 'https://baowen.xhh.com/manage/toutiao/gallery'
        self.id = 'chenjunming'
        self.pwd = 'Wzj1989863'

    def _get_cookie(self):
        self.browser.get(url=self.login_url)
        time.sleep(5)
        self.account = self.browser.find_element_by_xpath(
            '//form/div[@class="form-group relative mb20px"]/input[@name="account"]')
        self.account.send_keys(self.id)
        time.sleep(3)
        self.password = self.browser.find_element_by_xpath(
            '//form/div[@class="form-group relative mb20px"]/input[@name="password"]')
        time.sleep(3)
        self.password.send_keys(self.pwd)
        time.sleep(20)
        cookie_list = self.browser.get_cookies()
        return str(cookie_list)

    def _save_cookie(self, cookie):
        base_path = os.getcwd()
        with open(base_path+'\cookie.txt', 'w')as fp:
            fp.write(cookie)

    def run(self):
        cookie = self._get_cookie()
        self._save_cookie(cookie)
        print('*************cookie更新完成*************')
        self.browser.quit()


if __name__ == "__main__":
    auto = AutoCookie()
    auto.run()
