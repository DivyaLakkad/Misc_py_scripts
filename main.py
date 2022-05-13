import os
import sys
import csv
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

import csv
import time

import base64
import email
import smtplib
import imaplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import time

import openpyxl

import logging

import shutil

caps = DesiredCapabilities().CHROME
caps["pageLoadStrategy"] = "normal"
delay = 10

EMAIL_USER = 'graham.scripting@gmail.com'
EMAIL_PW = 'directorchris'

global pm_list_popped

pm_list_popped = []

# def get_work_order():
#     line_count = 0
#     with open(WORK_ORDER_FILE_PATH, mode='r') as csv_file:
#         csv_reader = csv.reader(csv_file, delimiter=',')
#         for row in csv_reader:
#             if line_count > 0:
#                 work_orders_given.append(row[0])
#             line_count += 1
#     return work_orders_given


def login(browser, attempt_num=0):
    logger.info('Attempting log-in attempt: {}'.format(attempt_num))
    auth_not_seen = False
    auth_performed = False

    def perform_auth(browser, time_to_check):
        counter = 0
        try:
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'otp')))
        except:
            logger.warning('Did not see auth page')
            raise Exception
        while True:
            if counter < 24:
                counter += 1
                time.sleep(5)
                mail = imaplib.IMAP4_SSL("imap.gmail.com")
                mail.login(EMAIL_USER, EMAIL_PW)
                time.sleep(.5)
                mail.select('Inbox')
                typ, items = mail.search(None, '(UNSEEN SUBJECT "New Authentication Request")')

                if typ == "OK":
                    items = items[0].split()
                    if len(items) > 0:
                        for num in items:
                            typ, data = mail.fetch(num, '(RFC822)')
                            raw_email = data[0][1]
                            raw_email_string = raw_email.decode('utf-8')
                            time.sleep(.5)
                            email_message = email.message_from_string(raw_email_string)

                            time_sent = email_message['Date'].split(', ')[1]
                            time_sent = time_sent.split(' +')[0]
                            time_received = email.utils.parsedate(email_message['Date'])
                            time_received = time.mktime(time_received)

                            if (time_to_check < time_received) and (
                                    email_message['Subject'] == 'New Authentication Request'):
                                for part in email_message.walk():
                                    if part.get_content_type() == 'text/html':
                                        email_text = part.get_payload()
                                        req_id = email_text.split(
                                            '<p style="font-size:26px;line-height:1em;font-weight:500;padding:45px 0 55px 0;margin:0;text-align:center">')[
                                            1]
                                        req_id = \
                                        req_id.split('\r\n\t\t\t\t\t\t\t</p>')[0].split('\r\n\t\t\t\t\t\t\t\t')[1]
                                        try:
                                            browser.find_element_by_id('otp').send_keys(req_id)
                                            time.sleep(2)
                                            browser.find_element_by_class_name('primary').click()
                                            return True
                                        except:
                                            logger.error('Unable to perform authentication')
                                            raise Exception
            else:
                logger.error('Email not retrieved in time')
                raise Exception

    url = "http://maximo.cnrl.com"
    browser.get(url)
    WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'username')))
    time_to_check = datetime.utcnow()
    time_to_check = datetime.timestamp(time_to_check)
    browser.find_element_by_id('username').send_keys('basharaj')
    browser.find_element_by_id('password').send_keys('U0n0ico00!')
    browser.find_element_by_class_name('ping-button').click()
    try:
        auth_performed = perform_auth(browser, time_to_check)
    except:
        auth_not_seen = True
    try:
        if auth_not_seen or auth_performed:
            logger.info('Auth seen')
            WebDriverWait(browser, 2 * delay).until(EC.presence_of_element_located((By.ID, 'titlebar-tb_gotoButton')))
        else:
            logger.error('Did not see Maximo main page')
            raise Exception
    except:
        if attempt_num < 5:
            logger.warning('Failed log-in attempt: {}'.format(attempt_num))
            # browser.quit()
            attempt_num += 1
            browser = webdriver.Chrome(os.path.join(os.path.dirname(__file__), 'chromedriver.exe'),
                                       desired_capabilities=caps)
            time.sleep(2)
            login(browser, attempt_num)
        else:
            logger.error('Log-in attemps failed. Shutting down the script.')
            # browser.quit()
            # sys.exit()
    return browser

#
# def go_to_tracking_main(browser):
#     logger.info('Attempting to open tracking main page')
#     browser.find_element_by_id('titlebar-tb_gotoButton').click()
#     WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'menu0_WO_MODULE_a')))
#     browser.find_element_by_id('menu0_WO_MODULE_a').click()
#     browser.find_element_by_id('menu0_WO_MODULE_sub_changeapp_PLUSGWO_a').click()
#     WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'm6a7dfd2f_tbod_tempty_tcell-c')))
#     WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'ffAS')))
#
#
# def perform_quicksearch(browser, work_order_given):
#     logger.info('Searching for work order: {}'.format(work_order_given))
#     browser.find_element_by_id('quicksearch').send_keys(work_order_given)
#     browser.find_element_by_id('quicksearch').send_keys(Keys.ENTER)
#     time.sleep(3)  # hacky way of allowing obtain_data's try part to work as expected in for loop

#
# def obtain_data(browser):
#     WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'mad3161b5-tb')))
#     work_order = browser.find_element_by_id('mad3161b5-tb').get_attribute('value')
#     location = browser.find_element_by_id('m7b0033b9-tb').get_attribute('value')
#     area = browser.find_element_by_id('md99142f3-tb').get_attribute('value')
#     asset = browser.find_element_by_id('m3b6a207f-tb').get_attribute('value')
#     serial = browser.find_element_by_id('mf198b952-tb').get_attribute('value')
#     asset_current_location = browser.find_element_by_id('m5c44da2c-tb').get_attribute('value')
#     parent_wo = browser.find_element_by_id('m34b721ed-tb').get_attribute('value')
#     failure_class = browser.find_element_by_id('mc636fea-tb').get_attribute('value')
#     problem_code = browser.find_element_by_id('m956a3e50-tb').get_attribute('value')
#     gl_account = browser.find_element_by_id('ma26371c5-tb').get_attribute('value')
#     work_order_stage = browser.find_element_by_id('m228c7f76-tb').get_attribute('value')
#     site = browser.find_element_by_id('m8b784a13-tb').get_attribute('value')
#     status = browser.find_element_by_id('mf2a4f2b7-tb').get_attribute('value')
#     status_date = browser.find_element_by_id('me2098fdd-tb').get_attribute('value')
#     status_changes = browser.find_element_by_id('maa25dd03-cb_img').get_attribute('checked')  # uhm?
#     is_task = browser.find_element_by_id('m43467836-cb_img').get_attribute('checked')
#     work_approver_name = browser.find_element_by_id('m2a37bdf8-tb').get_attribute('value')
#     work_approver_code = browser.find_element_by_id('m2a37bdf8-tb2').get_attribute('value')
#     work_type = browser.find_element_by_id('mcae254e6-tb').get_attribute('value')
#     description = browser.find_element_by_id('mad3161b5-tb2').get_attribute('value')
#
#     xlsx_work_order = work_order
#     xlsx_work_approver_code = work_approver_code
#     xslx_work_approver_name = work_approver_name
#     xslx_cost_center = gl_account.split('.')[0]
#     xslx_location = location
#     xslx_majmin = '.'.join(gl_account.split('.')[1:])
#     xlsx_description = description
#
#     xlsx_list = [xlsx_work_order, xlsx_work_approver_code, xslx_work_approver_name, xslx_cost_center, xslx_location,
#                  xslx_majmin, xlsx_description]
#     order_information_list = [work_order, location, area, asset, serial, asset_current_location, parent_wo,
#                               failure_class, problem_code, gl_account, work_order_stage, site, status, status_date,
#                               status_changes, is_task, work_approver_name, work_approver_code, work_type]
#
#     logger.info('The obtained work order information: {}'.format(order_information_list))
#     logger.info('The obtained xlsx information: {}'.format(xlsx_list))
#     return xlsx_list, order_information_list



def main(browser,pm_list_bla, count):
    first_run = True
    quit_script = False
    pm_list_popped = pm_list_bla.copy()
    # pm_list = [1493, 4581, 99534, 99535, 99536, 1097, 1234, 3142, 5751, 1044, 1099, 1108, 1130, 1190, 1236, 1326, 1699, 100608, 100617, 100619, 101383, 101391, 101394, 101395, 101396, 101593, 1071, 109799, 1110, 1129, 1137, 113782, 113783, 113784, 113785, 1154, 1155, 1169, 1170, 1227, 1244, 130669, 1365, 137122, 1442, 1443, 1444, 144774, 144776, 147744, 148787, 1488, 1518, 1529, 1735, 1736, 1874, 2131, 2581, 3327, 3328, 3429, 3447, 3475, 3685, 4613, 4626, 4627, 4659, 4660, 4661, 4662, 4663, 4684, 4742, 4758, 4761, 4762, 4763, 4764, 4765, 4766, 4767, 4768, 4771, 4772, 4773, 4774, 4775, 4786, 4788, 4789, 4791, 4793, 4794, 4795, 4796, 4798, 4799, 4800, 4801, 4868, 4869, 4870, 4871, 4872, 4873, 4881, 4882, 4883, 4909, 4910, 4915, 4921, 4976, 4977, 5068, 5069, 5074, 5075, 5181, 5183, 5184, 5185, 5186, 5187, 5188, 5189, 5190, 5206, 5207, 5208, 6499, 90465, 90466, 90467, 90468, 90469, 90560, 96922, 97409, 97410, 97412, 97413, 97450, 97457, 97459, 97461, 97462, 97463, 98492, 98493, 98808, 98809, 98810, 98811, 98812, 98813, 98814, 98815, 98816, 98817, 98818, 98819, 98820, 98821, 98822, 98823, 98824, 98825, 98826, 98827, 98828, 98829, 98830, 98832, 98833, 98834, 98835, 98917, 98918, 98919, 98920, 98921, 98922, 98923, 98924, 98925, 98926, 98927, 98928, 99004, 99005, 99006, 99007, 99008, 99009, 99010, 99011, 99012, 99013, 99014, 99015, 99016, 99017, 99018, 99019, 99341, 99377, 99378, 99379, 99380, 99444, 3141, 4639, 46857, 99664, 1522, 1655, 1679, 3719, 44979, 1061, 115777, 137621, 147048, 1701, 98949, 98950, 98954, 98961, 98962, 110910, 110914, 110928, 110944, 110945, 110952, 111018, 111074, 111081, 111082, 111085, 111086, 111124, 111136, 111137, 111273, 111282, 111285, 111304, 111510, 111562, 111688, 111691, 111692, 111693, 111694, 111695, 111696, 111698, 112031, 112033, 112037, 112039, 112045, 112047, 112048, 112050, 112051, 112052, 112053, 112054, 112055, 112056, 112060, 112080, 112094, 112095, 112098, 112099, 112135, 112141, 112149, 112151, 112153, 112154, 112155, 112159, 112168, 112201, 112208, 112213, 112217, 112220, 112619, 112620, 112636, 113373, 113376, 113808, 113810, 113813, 113820, 113878, 113880, 113919, 113923, 113926, 113927, 113928, 113929, 113935, 113936, 113937, 1151, 115372, 115373, 115484, 115503, 115504, 115573, 117908, 117910, 117911, 117912, 117941, 117963, 117967, 118020, 118324, 119002, 119003, 119004, 119005, 119006, 119060, 119814, 1225, 130135, 132921, 133379, 133384, 133392, 133393, 133394, 133395, 133396, 133397, 133398, 133399, 133400, 133401, 133402, 133404, 133405, 133406, 133407, 133408, 133409, 133410, 133411, 133412, 134264, 134265, 134277, 134281, 134284, 134292, 134298, 134321, 134322, 134323, 134469, 134470, 134471, 134472, 134473, 134474, 134475, 134476, 134507, 134508, 134509, 134510, 134515, 134516, 134517, 134518, 134525, 134533, 134534, 134535, 134536, 134537, 134560, 134561, 134562, 134563, 134564, 134565, 134566, 134567, 134568, 134612, 134646, 134733, 134746, 134755, 134758, 134761, 134879, 134880, 134881, 134882, 134883, 134892, 134893, 134919, 134924, 134925, 134926, 134927, 134959, 134970, 134976, 134977, 134979, 134980, 134981, 134982, 134987, 135191, 135197, 135198, 135199, 135213, 135215, 135216, 135217, 135218, 135219, 135311, 135325, 135333, 135337, 135338, 135340, 135449, 135458, 135461, 135486, 135764, 135767, 135773, 135776, 135777, 135793, 135794, 135819, 135820, 135821, 135822, 135823, 135824, 135825, 135826, 135862, 135894, 135895, 135923, 135929, 135939, 135940, 135941, 135942, 135943, 135944, 135945, 136027, 136028, 136029, 136035, 136039, 136257, 136258, 136263, 136264, 136265, 136266, 136267, 136284, 136295, 136296, 136309, 136310, 136443, 136444, 136445, 136446, 136447, 136449, 136450, 136451, 136452, 136453, 136454, 136465, 136490, 136491, 136503, 136509, 136510, 136511, 136522, 136523, 136524, 136525, 136526, 136527, 136533, 136544, 136553, 136554, 136555, 136581, 136582, 136584, 136593, 136594, 136595, 136598, 136599, 136600, 136601, 136602, 136603, 136604, 136605, 136606, 136607, 136609, 136696, 136697, 136699, 136700, 136701, 136702, 136703, 136704, 136705, 136706, 136714, 136715, 136732, 136733, 136745, 136750, 136751, 136752, 136753, 136754, 136755, 136756, 136758, 136760, 136766, 136767, 136772, 136773, 136783, 136787, 136788, 136789, 136790, 136794, 136795, 136797, 136806, 136811, 136813, 136814, 136815, 136816, 136817, 136818, 136842, 136846, 136847, 136850, 136853, 136856, 136857, 136861, 136863, 136866, 137121, 137123, 137124, 137126, 137127, 137156, 137158, 137165, 137167, 137168, 137169, 137170, 137171, 137178, 137214, 137216, 137217, 137218, 137219, 137221, 137222, 137234, 137235, 137236, 137237, 137238, 137241, 137242, 137245, 137246, 137247, 137325, 137328, 137343, 137344, 137346, 137348, 137349, 137350, 137351, 137357, 1378, 143795, 143801, 145035, 145205, 145362, 145404, 145460, 145461, 145548, 145550, 145551, 145558, 145559, 145560, 145561, 145562, 145563, 145569, 145570, 145572, 145573, 145574, 145924, 145986, 145987, 146014, 146119, 146161, 146319, 146320, 146321, 146538, 146540, 146543, 146554, 147037, 147120, 147272, 147737, 147739, 147747, 147750, 147751, 147752, 147753, 147757, 147765, 147766, 147837, 147840, 147855, 147856, 147859, 147860, 147877, 147881, 147882, 147883, 147893, 147895, 147902, 147907, 147918, 147920, 147925, 147926, 147927, 147940, 147954, 147957, 1480, 148052, 148054, 148055, 148056, 148057, 148058, 148059, 148060, 148061, 148065, 148109, 148111, 148112, 148117, 148119, 148133, 148134, 148135, 148157, 148158, 148165, 148166, 148167, 148186, 148187, 148230, 148231, 148232, 148233, 148244, 148249, 148955, 1530, 1537, 1582, 1647, 1669, 1696, 1877, 42731, 42732, 42733, 42734, 42736, 42738, 42739, 45137, 4612, 46856, 4728, 4759, 4797, 5227, 54973, 55221, 6264, 6350, 6351, 6617, 6618, 6619, 6620, 6621, 6622, 92004, 92008, 92009, 92010, 92011, 92012, 92013, 92014, 92015, 92016, 92017, 92018, 92019, 92020, 92021, 92022, 92023, 92024, 92025, 92026, 92027, 98995, 98996, 111119, 111120, 111121, 111272, 111274, 111275, 1118, 112389, 112516, 112517, 114322, 114867, 115979, 115980, 115981, 115983, 115984, 115985, 115986, 115987, 115988, 117029, 117030, 117031, 117032, 117033, 117034, 117036, 1180, 1183, 1184, 120230, 1203, 1204, 1205, 1206, 1207, 1220, 1221, 1229, 132536, 1327, 134382, 134383, 134684, 135488, 136819, 137316, 137317, 144258, 146056, 147580, 148049, 148394, 148527, 1644, 1645, 1673, 1674, 3120, 3269, 4701, 4731, 4956, 4958, 5237, 54925, 55117, 55128, 5851, 5852, 5853, 5854, 5855, 5856, 5857, 5858, 5859, 5860, 5861, 5862, 5863, 5864, 5865, 5866, 5867, 5868, 5869, 5870, 5871, 5872, 5873, 5874, 5876, 5877, 5878, 5879, 5880, 5882, 5883, 5886, 5887, 5888, 5890, 5891, 5892, 5893, 5894, 5895, 5896, 5897, 5898, 5899, 5900, 5901, 5902, 5903, 5904, 5905, 5906, 5907, 5908, 5909, 5910, 5911, 5912, 5913, 5914, 5915, 5916, 5917, 5919, 5920, 5921, 5922, 5923, 5924, 5925, 5926, 5927, 5928, 5929, 5930, 5931, 5933, 5934, 5935, 5936, 5938, 5939, 5940, 5941, 5942, 5943, 5944, 5945, 5946, 5947, 5948, 5949, 5950, 5951, 5952, 5953, 5954, 5955, 5956, 5957, 5958, 5959, 5961, 5962, 5963, 5964, 5965, 5966, 5967, 5968, 5969, 5970, 5971, 5972, 5973, 5974, 5975, 5976, 5977, 5978, 5979, 5980, 5981, 5982, 5983, 5984, 5985, 5986, 5987, 5988, 5989, 5990, 5991, 5992, 5993, 5994, 5995, 5996, 5997, 5998, 5999, 6001, 6002, 6003, 6004, 6005, 6006, 6007, 6008, 6009, 6010, 6011, 6012, 6013, 6014, 6015, 6016, 6017, 6018, 6019, 6020, 6021, 6022, 6023, 6024, 6025, 6026, 6027, 6028, 6029, 6030, 6031, 6032, 6033, 6035, 6036, 6037, 6038, 85720, 85721, 85723, 85724, 85725, 85726, 85727, 85728, 85729, 85730, 85731, 85732, 85733, 85734, 85736, 85737, 85738, 85739, 85740, 85741, 85742, 85743, 85744, 85745, 85746, 85747, 87661, 87662, 87663, 90215, 93606, 93607, 93609, 93610, 99420, 99421, 99429, 99430, 111283, 119007, 1450, 148185, 2048, 3260, 3268, 97853, 97866, 1247, 1341, 1653, 3276, 5548, 5061, 5238, 98994]


    critical_list = []
    labour_list = []
    material_list = []
    service_list = []
    tool_list = []
    total_list = []
    popped_list = []

    logger.info('Script starting . . .')

    try:
        #After Login Stuffs
        for pm in pm_list_bla:
            time.sleep(1)
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'titlebar-tb_homeButton'))).click()
            time.sleep(3)
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'FavoriteApp_PLUSGPM'))).click()

            search_pm_box = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'quicksearch')))
            time.sleep(1)
            search_pm_box.send_keys(f"={pm}")
            search_pm_box.send_keys(Keys.RETURN)
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'm8131f89b-img'))).click()
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'LOCATIONS_applink_undefined_a_tnode'))).click()
            critical_rb_elem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'm87d3cd44-tb')))
            critical_rb_value = critical_rb_elem.get_attribute("textContent")
            if critical_rb_value == "":
                critical_list.append('Empty')
            else:
                critical_list.append(critical_rb_value)

            #go back to main Pm page
            main_page = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'linkedAppTitle')))
            main_page.click()
            #JobPlan
            #if no Job plan

            job_plan = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'm68525dae-tb')))
            if job_plan.get_attribute('value')== "":
                time.sleep(1)
                popped_value = pm_list_popped.pop(0)
                popped_list.append(popped_value)
                main(browser, pm_list_popped, count)
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'm68525dae-img'))).click()
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'NORMAL_applink_undefined_a'))).click()
            #ON jOB pLAN pAGE
            time.sleep(1)
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'mainrec-pg'))).click()
            # WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'toolbar2_tbs_1-co_0'))).click()
            time.sleep(1)
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'toolbar2_tbs_1_tbcb_0_action-img'))).click()

            # WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'menu0'))).click()
            time.sleep(1)
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'menu0_VIEWTOTALS_OPTION'))).click()
            #Getting Costs from small menu that opens up
            #Labour
            labour  = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'm291a21bd-tb')))
            labour_val = labour.get_attribute("value")
            labour_list.append(labour_val)

            # MAterial Cost
            material_cost = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'mb0137007-tb')))
            material_val = material_cost.get_attribute("value")
            material_list.append(material_val)

            # Service Cost
            service_cost = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'mc7144091-tb')))
            service_val = service_cost.get_attribute("value")
            service_list.append(service_val)

            # Tool_cost
            tool_cost = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'm5970d532-tb')))
            tool_val = tool_cost.get_attribute("value")
            tool_list.append(tool_val)

            # Total Cost
            total_cost = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'm2e77e5a4-tb')))
            total_val  = total_cost.get_attribute("value")
            total_list.append(total_val)

            time.sleep(2)
            #Okay Button
            WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'm4d0384d3-pb'))).click()

            logger.info(f"All Values for {pm}: Critical_Val: {critical_rb_value}__labour:{labour_val}__material: {material_val}__service: {service_val}__tool: {tool_val}__total: {total_val}")


            time.sleep(1)
            main_page = WebDriverWait(browser, delay).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'linkedAppTitle')))
            main_page.click()


            try:
                # No Button for Save confirmation
                btn_no = WebDriverWait(browser, 5).until(
                        EC.presence_of_element_located((By.ID, 'm96ad0396-pb')))
                btn_no.click()
            except:
                pass
            popped_value = pm_list_popped.pop(0)
            popped_list.append(popped_value)
    except:
        browser.quit()
        count = count + 1        
        run_loop(pm_list_popped, browser, count)
        
    finally:
        a = {'PM Number': popped_list, 'Critical Cost': critical_list, 'Labour Cost': labour_list, 'Material Cost': material_list, 'Service Cost': service_list, 'Tool Cost': tool_list, 'Total Cost':total_list}
        df1 = pd.DataFrame({'Popped List': popped_list})
        df2 = pd.DataFrame({'Not Done List': pm_list_popped})

        df = pd.DataFrame.from_dict(a, orient='index')
        df = df.transpose()

        if df.empty == False and df1.empty == False:
            df.to_csv(f'CostData_finally_{count}.csv',index=False)
            df1.to_csv(f'popped_list_{count}.csv', index=False)
            df2.to_csv(f'Not_Done_List{count}.csv', index=False)
        time.sleep(10)

def run_loop(pm_list_popped11, browser11, count):
    for i in range(len(pm_list_popped11)):
        browser = webdriver.Chrome(os.path.join(os.path.dirname(__file__), 'chromedriver.exe'), desired_capabilities=caps)
        driver = login(browser)
        main(driver, pm_list_popped11, count)
        # driver.quit()


if __name__ == "__main__":
    logging.basicConfig(filename='runtime.log', format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger()
    handler = logging.StreamHandler()
    formatter = logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    pm_list_popped = [136817, 136818,
               136842, 136846, 136847, 136850, 136853, 136856, 136857, 136861, 136863, 136866, 137121, 137123, 137124,
               137126, 137127, 137156, 137158, 137165, 137167, 137168, 137169, 137170, 137171, 137178, 137214, 137216,
               137217, 137218, 137219, 137221, 137222, 137234, 137235, 137236, 137237, 137238, 137241, 137242, 137245,
               137246, 137247, 137325, 137328, 137343, 137344, 137346, 137348, 137349, 137350, 137351, 137357, 1378,
               143795, 143801, 145035, 145205, 145362, 145404, 145460, 145461, 145548, 145550, 145551, 145558, 145559,
               145560, 145561, 145562, 145563, 145569, 145570, 145572, 145573, 145574, 145924, 145986, 145987, 146014,
               146119, 146161, 146319, 146320, 146321, 146538, 146540, 146543, 146554, 147037, 147120, 147272, 147737,
               147739, 147747, 147750, 147751, 147752, 147753, 147757, 147765, 147766, 147837, 147840, 147855, 147856,
               147859, 147860, 147877, 147881, 147882, 147883, 147893, 147895, 147902, 147907, 147918, 147920, 147925,
               147926, 147927, 147940, 147954, 147957, 1480, 148052, 148054, 148055, 148056, 148057, 148058, 148059,
               148060, 148061, 148065, 148109, 148111, 148112, 148117, 148119, 148133, 148134, 148135, 148157, 148158,
               148165, 148166, 148167, 148186, 148187, 148230, 148231, 148232, 148233, 148244, 148249, 148955, 1530,
               1537, 1582, 1647, 1669, 1696, 1877, 42731, 42732, 42733, 42734, 42736, 42738, 42739, 45137, 4612, 46856,
               4728, 4759, 4797, 5227, 54973, 55221, 6264, 6350, 6351, 6617, 6618, 6619, 6620, 6621, 6622, 92004, 92008,
               92009, 92010, 92011, 92012]
    count = 1
    browser = webdriver.Chrome(os.path.join(os.path.dirname(__file__), 'chromedriver.exe'), desired_capabilities=caps)

    # while True:

    time.sleep(5)
    run_loop(pm_list_popped, browser, count)
        # count = count + 1
