from selenium import webdriver
from pymongo import MongoClient
import time
import datetime
import xlwt


class YinBao(object):
    def __init__(self, username, password, start_time, end_time):
        self.username = username
        self.password = password
        self.start_time = start_time
        self.end_time = end_time
        self.client = MongoClient('mongodb://127.0.0.1:27017')
        self.YinBao = self.client['YinBao']['case']
        self.driver = webdriver.PhantomJS()

    def selenium_spider(self):
        """登录"""
        self.driver.maximize_window()
        self.driver.get('https://beta38.pospal.cn/account/signin')
        time.sleep(5)
        # 账号
        self.driver.find_element_by_xpath('//*[@id="txt_userName"]').send_keys(self.username)
        # 密码
        self.driver.find_element_by_xpath('//*[@id="txt_password"]').send_keys(self.password)
        # 点击登录
        self.driver.find_element_by_class_name('submit').click()
        time.sleep(2)
        # print(self.driver.get_cookies())
        # cookies = {i['name']:i['value'] for i in self.driver.get_cookies()}
        # print(cookies)

        # 获取销售单据的网页
        self.driver.get('https://beta38.pospal.cn/Report/Tickets')
        time.sleep(5)

        # 选择全部支付方式下拉列表框,并选择现金支付
        self.driver.find_element_by_xpath('//div[@id="ddl_paymethods"]').click()
        self.driver.find_element_by_xpath('//div[@id="ddl_paymethods"]//ul/li[@optionvalue="1"]').click()

        # 修改指定时间段, 若不指定,默认当天
        js = '$(".dateTimeRangeBox input:text:eq(0)").removeAttr("readonly")'
        self.driver.execute_script(js)
        js = '$(".dateTimeRangeBox input:text:eq(1)").removeAttr("readonly")'
        self.driver.execute_script(js)
        js = '$(".dateTimeRangeBox input:text:eq(0)").value="{}"'.format(self.start_time)
        self.driver.execute_script(js)
        js = '$(".dateTimeRangeBox input:text:eq(1)").value="{}"'.format(self.end_time)
        self.driver.execute_script(js)
        self.driver.find_element_by_xpath('//*[@class="timeInput hasDatepicker"][1]').clear()
        self.driver.find_element_by_xpath('//*[@class="timeInput hasDatepicker"][2]').clear()
        self.driver.find_element_by_xpath('//*[@class="timeInput hasDatepicker"][1]').send_keys(self.start_time)
        self.driver.find_element_by_xpath('//*[@class="timeInput hasDatepicker"][2]').send_keys(self.end_time)

        # 查询按钮
        self.driver.find_element_by_xpath('//div[@class="submitBtn"]').click()
        time.sleep(2)

        # with open('./1.html', 'w') as f:
        #     f.write(self.driver.page_source)
        # 修改每页显示10条信息
        self.driver.find_element_by_xpath('//*[@id="summaryInfo"]/div[2]/div').click()
        self.driver.find_element_by_xpath('//*[@id="summaryInfo"]/div[2]/div//ul/li[@optionvalue="10"]').click()
        time.sleep(2)
        # self.driver.save_screenshot('./1.png')

        while True:

            tr_list = self.driver.find_elements_by_xpath('//table[@id="mainTable"]/tbody/tr[not(@style)]')
            for tr in tr_list:
                # print(tr.text)
                # item = {}
                # item['流水号'] = tr.find_element_by_xpath('./td[2]').text
                serial_number = tr.find_element_by_xpath('./td[2]').text
                # item['日期'] = tr.find_element_by_xpath('./td[3]').text
                date = tr.find_element_by_xpath('./td[3]').text
                # item['类型'] = tr.find_element_by_xpath('./td[4]').text
                # item['收银员'] = tr.find_element_by_xpath('./td[5]').text
                cashier = tr.find_element_by_xpath('./td[5]').text
                # item['会员'] = tr.find_element_by_xpath('./td[6]').text
                # item['商品数量'] = tr.find_element_by_xpath('./td[7]').text
                # item['商品原价'] = tr.find_element_by_xpath('./td[8]').text
                # item['实收金额'] = tr.find_element_by_xpath('./td[9]').text
                cash = tr.find_element_by_xpath('./td[9]').text
                # item['折让金额'] = tr.find_element_by_xpath('./td[10]').text
                # item['利润'] = tr.find_element_by_xpath('./td[11]').text
                # # print(item)
                # data_list.append(item)
                self.YinBao.insert_one({'_id': serial_number, 'date': date, 'cashier': cashier, 'cash': cash})

            time.sleep(1)

            # 获取总页数和当前页,判断循环退出的条件
            maxlength = self.driver.find_element_by_xpath('//i[@class="pageNum"]').text
            maxlength = maxlength[1]
            value = self.driver.find_element_by_xpath('//input[@class="appointPage quantity"]').get_attribute('value')

            # self.driver.save_screenshot('./{}.png'.format(value))
            if int(value) >= int(maxlength):
                return
            self.driver.find_element_by_xpath('//span[@class="next"]').click()
            time.sleep(2)

    def save_excel(self):
        execl = xlwt.Workbook()
        my_sheet = execl.add_sheet('case')
        now_time = datetime.datetime.now()
        head = ['流水号', '日期', '收银员', '实收金额']

        for i in range(len(head)):
            my_sheet.write(0, i, head[i])
        num = 1
        for item in self.YinBao.find():
            my_sheet.write(num, 0, item['_id'])
            my_sheet.write(num, 1, item['date'])
            my_sheet.write(num, 2, item['cashier'])
            my_sheet.write(num, 3, item['cash'])

            num = num + 1

        execl.save('{}.xls'.format(now_time.date()))

        self.YinBao.drop()

    def run(self):
        self.selenium_spider()
        self.save_excel()

    def __del__(self):
        self.driver.quit()


def main():
    # YinBao('账号, '密码', '开始时间', '结束时间'),查询最大跨度3个月,并且严格按照以下格式输入参数
    # yinbao = YinBao('xxxxxx', 'xxxxxx', '2019.09.08 00:00', '2019.09.09 23:59')
    account = 'chen_junior'
    password = '1234567899'
    # account = str(account.decode('unicode-escape').encode('utf-8'))
    # password = str(password.decode('unicode-escape').encode('utf-8'))
    # print(type(account))
    # print(password)

    yinbao = YinBao(account, password, '2019.09.08 00:00', '2019.09.09 23:59')
    yinbao.run()


if __name__ == '__main__':
    main()
