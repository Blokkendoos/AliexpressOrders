"""
This script generates a json representation of all your AliExpress orders.
It also saves the information in a google sheet if you have set up one.

Usage:
export AE_username='your@emailaddress'
export AE_password='pwd'
python aliexpress.py json,screenshot foo.png

Source: https://github.com/kritarthh/AliexpressOrders

The package is dependant on lxml, pyquery, selenium and Chromedriver/PhantomJS Package.
"""

import os
import csv
import sys
import json
import time
import pickle

# import sheets
from pyquery import PyQuery as pq
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import Select


DEBUG = False


class AliExpress():

    UA_STRING = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.57 Safari/537.36'

    def __init__(self):
        self.driver = None
        self.get_driver()

    def get_driver(self, driver_type='Chrome', driver_path='chromedriver'):
        if driver_type == 'Chrome':
            from selenium.webdriver.chrome.options import Options
            opts = Options()
            opts.add_argument("user-agent=%s" % self.UA_STRING)
            opts.add_argument('--headless')
            opts.add_argument('--disable-gpu')
            assert driver_path is not None
            self.driver = webdriver.Chrome(driver_path, options=opts)
        elif driver_type == 'PhantomJS':
            dcap = dict(DesiredCapabilities.PHANTOMJS)
            dcap['phantomjs.page.settings.userAgent'] = (self.UA_STRING)
            self.driver = webdriver.PhantomJS(desired_capabilities=dcap)
        elif driver_type == 'Firefox':
            from selenium.webdriver.firefox.options import Options
            opts = Options()
            opts.headless = True
            self.driver = webdriver.Firefox(options=opts)
        else:
            raise Exception("Invalid driver type:" + driver_type)

    def close(self):
        if self.driver:
            self.driver.quit()
            self.driver = None

    def save_screenshot(self, file_name):
        assert self.driver is not None
        self.driver.save_screenshot(file_name)

    def login(self, email, passwd):
        self.driver.set_window_size(1366, 768)

        # restore cookies
        self.driver.get('https://login.aliexpress.com/buyer.htm')
        try:
            cookies = pickle.load(open('cookies.pkl', 'rb'))
            for cookie in cookies:
                # only load login.aliexpress.com compatible cookies
                if not (cookie['domain'] in ('.aliexpress.com', 'login.aliexpress.com')):
                    continue
                self.driver.add_cookie(cookie)
        except Exception as e:
            print("Error loading cookies, %s" % e)

        self.driver.get('https://trade.aliexpress.com/orderList.htm')

        try:
            if not os.path.exists('cookies.pkl'):
                raise Exception('Cookies not found.')
            # see if cookies worked or not
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'search-key'))
            )
            print("Cookies worked.")

        except Exception:
            # cookies did not work remove them
            try:
                os.remove('cookies.pkl')
            except OSError:
                pass

            print('Logging in...')
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.ID, 'fm-login-id'))
            )
            print('send pwd')
            # login and save cookies
            element.clear()
            element.send_keys(email)
            self.driver.find_element_by_xpath('//*[@id=\"fm-login-password\"]').send_keys(passwd)
            self.driver.find_element_by_tag_name('button').click()
            self.driver.switch_to.default_content()

        finally:
            # wait till the page is loaded
            element = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.ID, 'search-key'))
            )
            # save cookies for later use
            pickle.dump(self.driver.get_cookies(), open('cookies.pkl', 'wb'))

    def parse_orders_page(self, src, track=False):
        node = pq(src)
        l_orders = []
        for e in node('.order-item-wraper'):
            order = {
                'order_id': pq(e)('.order-head .order-info .first-row .info-body')[0].text,
                'order_url': pq(e)('.order-head .order-info .first-row .view-detail-link')[0].attrib['href'],
                'order_dt': pq(e)('.order-head .order-info .second-row .info-body')[0].text,
                'order_store': pq(e)('.order-head .store-info .first-row .info-body')[0].text,
                'order_store_url': pq(e)('.order-head .store-info .second-row a')[0].attrib['href'],
                'order_amount': pq(e)('.order-head .order-amount .amount-body .amount-num')[0].text.strip(),
                'product_list': [{
                    'title': pq(f)('.product-right .product-title a').attr['title'],
                    'url': pq(f)('.product-right .product-title a').attr['href'],
                    'amount': pq(f)('.product-right .product-amount').text().strip(),
                    'property': pq(f)('.product-right .product-policy a').attr['title'],
                } for f in pq(e)('.order-body .product-sets')],
                'status': pq(e)('.order-body .order-status .f-left').text(),
                'status_days_left': pq(e)('.order-body .order-status .left-sendgoods-day').text().strip()
            }
            # get tracking id
            if self.driver and track:
                try:
                    # TODO - handle not found exception
                    # t_button = self.driver.find_element_by_xpath('//*[@button_action="logisticsTracking" and @orderid="{}"]'.format(order['order_id']))
                    # hover = ActionChains(self.driver).move_to_element(t_button).perform()
                    time.sleep(5)

                    order['tracking_id'] = self.driver.find_element_by_css_selector('.ui-balloon .bold-text-remind').text.strip().split(':')[1].strip()
                    try:
                        # if present, It means, tracking has begun
                        order['tracking_status'] = self.driver.find_element_by_css_selector('.ui-balloon .event-line-key').text
                    except Exception:
                        # Check for no event which means tracking has not started or has not begun
                        try:
                            order['tracking_status'] = self.driver.find_element_by_css_selector('.ui-balloon .no-event').text.strip()
                            # if above passed, copy the tracking link and pass for manual tracking
                            order['tracking_status'] = 'Manual Tracking: ' + self.driver.find_element_by_css_selector('.ui-balloon .no-event a').get_attribute('href').strip()
                        except Exception:
                            order['tracking_status'] = '<Tracking Parse Error>'
                except Exception as e:
                    order['tracking_id'] = '<Error in Parsing Tracking ID>'
                    order['tracking_status'] = '<Tracking Parse Error due to Error in Parsing Tracking ID>'
                    print("Tracking id retrieval failed for order:" + order['order_id'] + " error: " + str(e))
                    pass
            l_orders.append(order)
        return l_orders

    def parse_orders(self, order_json_file='', cache_mode='webread', track=False):
        orders = []
        if cache_mode == 'webread':
            if self.driver == '':
                raise Exception("No Selenium driver found in webread mode.")
            # Verify number of orders and implement pagination
            source = self.driver.find_element_by_id('buyer-ordertable').get_attribute('innerHTML')
        elif cache_mode == 'localwrite':
            if order_json_file == '':
                raise Exception("Filename Missing. Please pass a valid filename to order_json_file.")
            source = self.driver.find_element_by_id('buyer-ordertable').get_attribute('innerHTML')
            open(order_json_file, 'wb').write(source.encode('utf-8'))
        elif cache_mode == 'localread':
            source = open(order_json_file, 'rb').read()
        else:
            raise Exception("Invalid cache_mode selected.")

        try:
            simple_pager_xpath = '//*[@id="simple-pager"]/div/label'
            cur_page, total_page = (int(i) for i in self.driver.find_element_by_xpath(simple_pager_xpath).text.split('/'))
            while cur_page <= total_page:
                print("current: {} total: {}".format(cur_page, total_page))
                source = self.driver.find_element_by_id('buyer-ordertable').get_attribute('innerHTML')
                orders.extend(self.parse_orders_page(source, track))
                if cur_page < total_page:
                    link_next = self.driver.find_element_by_xpath('//*[@id="simple-pager"]/div/a[text()="Next "]')
                    link_next.click()
                    cur_page, total_page = (int(i) for i in self.driver.find_element_by_xpath(simple_pager_xpath).text.split('/'))
                if cur_page == total_page:  # break after having parsed the last page
                    break
            return orders

        except Exception as e:
            print(e)

    def get_open_orders(self, cache_mode):

        aliexpress = {}

        elemAwaitingShipment = self.driver.find_element_by_id('remiandTips_waitSendGoodsOrders')
        elemAwaitingShipment.click()
        aliexpress['Not Shipped'] = self.parse_orders('ae1.htm', cache_mode)

        elemAwaitingDelivery = self.driver.find_element_by_id('remiandTips_waitBuyerAcceptGoods')
        elemAwaitingDelivery.click()
        aliexpress['Shipped'] = self.parse_orders('ae2.htm', cache_mode, track=True)

        elemAwaitingShipment = self.driver.find_element_by_id('remiandTips_waitBuyerPayment')
        elemAwaitingShipment.click()
        aliexpress['Order Awaiting Payment'] = self.parse_orders('ae3.htm', cache_mode)

        # completed orders
        self.driver.find_element_by_id('switch-filter').click()
        try:
            Select(self.driver.find_element_by_id('order-status')).select_by_value('FINISH')
        except Exception:
            self.driver.find_element_by_id('switch-filter').click()
            Select(self.driver.find_element_by_id('order-status')).select_by_value('FINISH')
        self.driver.find_element_by_id('search-btn').click()
        aliexpress['Order Completed'] = self.parse_orders('ae4.htm', 'webread')

        return(aliexpress)


if __name__ == '__main__':
    modes = sys.argv[1].split(',')
    try:
        ae = AliExpress()
        ae.login(os.environ['AE_username'], os.environ['AE_password'])
        if 'screenshot' in modes:
            # save a screenshot of the orders page
            screenshot_name = sys.argv[2]
            ae.save_screenshot(screenshot_name)
            print("%s saved." % screenshot_name)

    except Exception as e:
        print("login: %s" % e)
        ae.save_screenshot("error_%s.png" % time.asctime())
        ae.close()
        sys.exit(1)
    # sheets.clear_google_sheet(sheets.URL, sheets.SHEET_NAME)
    # sheets.save_aliexpress_orders(orders)

    cache_mode = 'webread'
    if DEBUG:
        cache_mode = 'localread'

    if 'json' in modes or 'csv' in modes:
        orders = ae.get_open_orders(cache_mode)
        if 'json' in modes:
            open('orders.json', 'w').write(json.dumps(orders))

    if DEBUG:
        with open('orders.json', 'r') as f:
            print("load orders.json")
            orders = json.load(f)
            print(orders['Shipped'][0])

    if 'csv' in modes:
        print("save CSV")
        csv_file = open('orders.csv', 'w')
        csvwriter = csv.writer(csv_file)
        first = True
        for order in orders['Shipped']:
            if first:
                first = False
                header = order.keys()
                csvwriter.writerow(header)
            csvwriter.writerow(order.values())
        csv_file.close()

    ae.close()
