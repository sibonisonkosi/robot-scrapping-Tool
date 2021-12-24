
import pickle
import numpy as np
from requests.api import get
from selenium.common import exceptions
# package for sharepoint
from shareplum import Office365
from shareplum import Site
from shareplum.site import Version
from requests.auth import HTTPBasicAuth
import requests
import json
# Selenium packages
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.touch_actions import TouchActions
import time
import os

class NpEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        if isinstance(obj, np.floating):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        return super(NpEncoder, self).default(obj)


class Methods:
    company_name_sharepoint = os.environ['COMPANY_SHAREPOINT']
    to_pass_credentials = HTTPBasicAuth('API token code 1',
                                        'API token code 2')
    url_campaigns = "https://api.woodpecker.co/rest/v1/campaign_list"
    campaign_ids = []
    authcookie = Office365('https://company_name_sharepoint.sharepoint.com', username=os.environ['COMPANY_EMAIL'],
                           password=os.environ['COMPANY_EMAIL_PASSWORD']).GetCookies()
    site = Site('https://company_name_sharepoint.sharepoint.com/sites/MarketingSalesReports/', version=Version.v365,
                authcookie=authcookie)
    folder = site.Folder('Shared Documents')
    dont_need_campaigns = open('dont_need_v1.txt')
    obj_NpEncoder = NpEncoder()
    

    def __init__(self, driver, url_login, url_campaign_edit):
        self.url_campaign_edit = url_campaign_edit
        self.url_login = url_login
        self.driver = driver
        self.response = requests.get(self.url_campaigns, auth=self.to_pass_credentials)
        self.all_campaigns = self.response.json()
        self.dont_need_campaigns_array = self.dont_need_campaigns.read().split(sep='\n')

    def get_all_campaigns(self):
        return self.all_campaigns

    def launch_login_page(self):
        self.driver.get(self.url_login)


    def isCookiesFound(self):
        try:
            cookies = pickle.load(open("cookies.pkl", "rb"))
            found = True
        except FileNotFoundError as e:
            found = False
        return found
    

    def load_cookies(self):
        return pickle.load(open("cookies.pkl", "rb"))
    

    def create_session(self):
        if self.isCookiesFound():
            cookies = self.load_cookies()
            for cookie in cookies:
                self.driver.add_cookie(cookie)
            self.driver.get(self.url_login)
        else:
            username = self.driver.find_element_by_name("login")
            username.send_keys(os.environ['WOODPECKER_USERNAME'])
            password = self.driver.find_element_by_name("password")
            password.send_keys(os.environ['WOODPECKER_PASSWORD'])
            password.send_keys(Keys.RETURN)
            pickle.dump(self.driver.get_cookies(), open("cookies.pkl", "wb"))  # save cookies


    def create_campaign_list_from_api(self):
        for all_camp in self.all_campaigns:
            if str(all_camp['id']) not in self.dont_need_campaigns_array:
                self.campaign_ids.append(all_camp['id'])
        return self.campaign_ids

    def get_total_pros_per_page(self, pros_rows):
        list_allpages_lastpage = []
        no_ = self.extract_prospect_list(pros_rows)
        first_split = int(str(no_ / 10).split('.')[0])
        second_split = int(str(no_ / 10).split('.')[1])

        list_allpages_lastpage.append(first_split)
        list_allpages_lastpage.append(second_split)
        return list_allpages_lastpage

    def extract_prospect_list(self, pros_rows):
        pros_no = pros_rows.split()[0].replace("(", '') if '(' in pros_rows.split()[0] else pros_rows.split()[0]
        print(pros_rows, pros_no)
        return int(pros_no)
    

    def get_existing_data(self):
        return json.load(open('scraped_prospect.json', encoding="utf8"))

    def find_element_by_Xpath(self, xpath):
        while True:
            try:
                try:
                    return WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
                    break
                except TimeoutException:
                    
                    print("timed out", TimeoutException)
                    return 'Not found'
                    break
            except NoSuchElementException:
                print('NoSuchElementException was thrown!')
    

    def get_mail_item_category(self, class_name):
        mail_items_list = [{
            'class_name': 'mail-item st-color1',
            'category': 'responded'},{
            'class_name': 'mail-item st-color2',
            'category': 'autoreplied'},{
            'class_name': 'mail-item st-color3',
            'category': 'bounced'},{
            'class_name': 'mail-item st-color4',
            'category': 'blacklisted'}]

        for i in range(len(mail_items_list)):
            category = 'Not Assigned'
            if class_name == mail_items_list[i]['class_name']:
                category = mail_items_list[i]['category']
                break

        return category

    def get_existing_emails_data(self):
        return json.load(open('emails_responses_v2.json', encoding="utf8"))
    
    def get_to_browse_url(self):
        url_list = [
            'https://app.woodpecker.co/panel?u=2090#inbox/24959/responded',
            'https://app.woodpecker.co/panel?u=2090#inbox/24959/autoreplied'
        ]
        return url_list

    def extracting_email_response(self):
        campaign_ids = self.create_campaign_list_from_api()
    
        # Go to inbox
        to_inbox_label = self.find_element_by_Xpath('//*[@id="divMenu"]/div/div[2]/div[1]/a[5]')    
        self.driver.execute_script("arguments[0].click();", to_inbox_label)
        time.sleep(15)
        replies_total_list = self.get_inbox_total()

        for email_url_index in range(len(self.get_to_browse_url())):
            self.driver.get(self.get_to_browse_url()[email_url_index])
            row_no_list = replies_total_list[email_url_index]
            count = 0

            for x in range(len(row_no_list)):
                # what is used e.g: [4,5] four is the number of pages and 5 is the number of prospect
                # in the last page
                all_ = row_no_list[x]
                exclude_class = ('mail-item st-color4',
                    'mail-item st-color4 selected',
                    'mail-item selected',
                    'mail-item')
                total_pages = row_no_list[x] + 1
                if x == 0:
                    while True:
                        count = count + 1
                        print("pages number: %s" % count , 'of:', total_pages)
                        for y in range(1,20): # 51 should be the max
                            print('row number: %s' %y)

                            existing_data = self.get_existing_emails_data()

                            row_ = self.find_element_by_Xpath(
                                '//*[@id="divContent"]/div/div[1]/div[4]/div[2]/div/div[4]/div/div[3]/div/div/div/div[%s]' %y)
                            class_value =  row_.get_attribute('outerHTML')
                            class_name = self.get_class_name(class_value)
                            if class_name not in exclude_class:
                                self.driver.execute_script("arguments[0].click();", row_)
                                time.sleep(5)
                                email_campaign = self.find_element_by_Xpath(
                                    '//*[@id="divContent"]/div/div[1]/div[4]/div[2]/div/div'
                                    '[6]/div/div/div/div[1]/div[1]/div[1]/div'
                                    '[3]/div/div[1]/div[2]')
                                    
                                if email_campaign == 'Not found':
                                    email_campaign = self.find_element_by_Xpath('//*[@id="divContent"]'
                                        '/div/div[1]/div[4]/div[2]/div/div[6]/div/div/div/div[1]/div[1]'
                                        '/div[1]/div[2]/div/div[1]/div[2]')
                                    if email_campaign == 'Not found':
                                        self.driver.switch_to.default_content()
                                        time.sleep(5)
                                        all_ = all_ - 1
                                        continue
                                
                                from_span_tag = self.find_element_by_Xpath('//*[@id="divContent"]/div/div[1]/'
                                    'div[4]/div[2]/div/div[6]/div/div/div/div[1]/div[3]/div/span[1]')

                                if '@' in from_span_tag.text:
                                    from_element = str(from_span_tag.text)
                                    from_prospect_value = from_element[6:len(from_element)]
                                else:
                                    from_prospect = self.find_element_by_Xpath('//*[@id="divContent"]/div/div[1]'
                                        '/div[4]/div[2]/div/div[6]/div/div/div/div[1]/div[3]/div/span[2]')
                                    from_prospect_value = from_prospect.text

                                subject_element = self.find_element_by_Xpath('//*[@id="divContent"]/div/div[1]/div[4]'
                                '/div[2]/div/div[6]/div/div/div/div[1]/div[4]/div')
                                subject_element_value = subject_element.text

                                #  testing Iframe to retrive content inside #document
                                iframe = self.find_element_by_Xpath('//*[@id="divContent"]/div/'
                                    'div[1]/div[4]/div[2]/div/div[6]/div/div/div/div[2]/div/iframe')
                                # You can use element.text to extract inner text 
                                email_campaign_value = email_campaign.text
                                # Swicting to iFrame. Note: make sure you store values before 
                                # switching to iframe because they will no longer exist within iframe
                                self.driver.switch_to.frame(iframe)
                                body_element = self.find_element_by_Xpath('/html')
                                body_element_value = body_element.text
                                current_camp_id = self.get_campaign_id(email_campaign_value)
                                exist = False
                                for camp_index in range(len(existing_data)):
                                    
                                    prosp_list = existing_data[camp_index]['prospect']
                                    mapped_camp_id = existing_data[camp_index]['Campaign id']
                                    mapped_from_prospect = self.isProspectFound(prosp_list, from_prospect_value)
                                    if current_camp_id == mapped_camp_id and mapped_from_prospect:
                                        result = self.isEmailSubjectFound(prosp_list, from_prospect_value, subject_element_value, body_element_value)
                                        if result == 'Exist':
                                            exist = True
                                            break
                                if exist:
                                    self.driver.switch_to.default_content()
                                    time.sleep(5)
                                    all_ = all_ - 1
                                    print('Exist')
                                    continue

                                prospects_array = []
                                email_response_array = []
                                email_response_array.append({ 
                                                        'Subject': subject_element_value,
                                                        'Body': body_element_value})
                                
                                prospects_array.append({
                                                'Email': from_prospect_value,
                                                'Response type': self.get_mail_item_category(class_name),
                                                'Response': email_response_array
                                            })
                                found = False
                                if current_camp_id in campaign_ids:
                                    for camp_index in range(len(existing_data)):
                                        
                                        mapped_camp_id = existing_data[camp_index]['Campaign id']
                                        # continue here try using in to fix error
                                        if current_camp_id == mapped_camp_id:
                                            found = True
                                            existing_data[camp_index]['prospect'].append(prospects_array[0])
                                            print('Added to existing')
                                            break
                                            
                                if not found:
                                    temp = {
                                                'Campaign id': current_camp_id,
                                                'prospect': prospects_array
                                            }
                                    existing_data.append(temp)
                                    print('New Added')
                                self.driver.switch_to.default_content()
                                self.save_to_json_file(existing_data)
                            else:
                                time.sleep(5)
                        break
                  
    def get_to_browse_emails_xpath(self):
        get_xpath = [
            '//*[@id="divContent"]/div/div[1]/div[4]/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[1]/div',
            '//*[@id="divContent"]/div/div[1]/div[4]/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[7]/div'
        ]
        return get_xpath

    def get_inbox_total(self):
        values_holder_array =[]
        for xpath in self.get_to_browse_emails_xpath():
            all_inbox_number_label = self.find_element_by_Xpath(xpath)
            all_inbox_number_value = self.get_html_value(all_inbox_number_label)
            values_holder_array.append(self.get_total_emails_per_page(all_inbox_number_value))
        return values_holder_array       

    def get_total_emails_per_page(self, all_emails_tag):
        list_allpages_lastpage = []
        no_ = self.remove_charecters_from_total_inbox_number(all_emails_tag)
        first_split = int(str(no_ / 50).split('.')[0])
        second_split = int(str(no_ / 50).split('.')[1])

        if second_split > 50:
            second_split = second_split - 50
            first_split = first_split + 1

        list_allpages_lastpage.append(first_split)
        list_allpages_lastpage.append(second_split)
        return list_allpages_lastpage

    def remove_charecters_from_total_inbox_number(self, all_emails_tag):
        email_no = all_emails_tag.split()[1]
        email_no = str(email_no).replace('(','')
        email_no = email_no.replace(')','')

        return (int(email_no))
        
    
    def get_campaigns_to_select_numbers(self):
        my_list= []
        scroll = self.find_element_by_Xpath('//*[@id="body"]/div[7]/div/div/div[2]')
        self.driver.execute_script('arguments[0].scrollTop = arguments[0].scrollHeight', scroll)

        for i in range(1, len(self.all_campaigns) + 1):
            if i > 30:
                scroll = self.find_element_by_Xpath('//*[@id="body"]/div[7]/div/div/div[2]')
                scroll.send_keys(Keys.DOWN)
            campaign_div_number = self.find_element_by_Xpath(
                '//*[@id="body"]/div[7]/div/div/div[2]/div[2]/div[1]/div/div[2]/div/div[2]/div[%s]/div/div' % i)
            html_value = self.get_html_value(campaign_div_number)
            campaign_ids = self.create_campaign_list_from_api()
            current_campaign_id = self.get_campaign_id(html_value)
            for x in campaign_ids:
                if x == current_campaign_id:
                    self.driver.execute_script("arguments[0].click();", campaign_div_number)
                    my_list.append(i)
                    clear_filter = self.find_element_by_Xpath('//*[@id="divContent"]/div/div[1]/div[2]/div/img')
                    self.driver.execute_script("arguments[0].click();", clear_filter)
                    filter_by_campaign = self.find_element_by_Xpath('//*[@id="divContent"]/div/div[1]/div[1]/div/div/span[3]')    
                    self.driver.execute_script("arguments[0].click();", filter_by_campaign) 
                    break

        return len(my_list)
    
    def get_html_value(self, html_tag):
        return html_tag.get_attribute('innerHTML')
    
    def get_class_name(self, html_tag):
        tag_list = str(html_tag).split('><')
        class_name = tag_list[0]
        class_name = class_name[12:-1]
        return class_name

    
    def get_campaign_id(self, campaign_name):
        all_list = self.get_all_campaigns()
        for i in range(len(all_list)):
            id_ = 0
            if campaign_name.lower() == all_list[i]['name'].lower():
                id_ = all_list[i]['id']
                break
        return id_

    def strip_tags(self, html):
        s = self.obj_html_stripper
        s.feed(html)
        return s.get_data()

    def save_to_json_file(self, dict):
        with open('emails_responses_v2.json', 'w', encoding='utf-8') as f:
                        json.dump(dict, f, indent=4,
                                  ensure_ascii=False)
    
    def isProspectFound(self,prosp_list, to_find_prosp):
        found = False
        for i in range(len(prosp_list)):
            prosp_dict = prosp_list[i]
            if to_find_prosp == prosp_dict['Email']:
                found = True
                break
        return found

    def isEmailSubjectFound(self,prosp_list, to_find_prosp, to_find_subject, to_find_body):
        found = 'Body and Subject are not found'
        for i in range(len(prosp_list)):
            prosp_dict = prosp_list[i]
            if to_find_prosp == prosp_dict['Email']:
                for x in range(len(prosp_dict['Response'])):
                    resp_dict = prosp_dict['Response'][x]
                    if to_find_subject == resp_dict['Subject'] and to_find_body == resp_dict['Body']:
                        found = 'Exist'
                        break
                    elif to_find_subject == resp_dict['Subject']:
                        found = 'Update'
                        break
                break
        return found

    def isEmail_conv_exist(self,current_camp_id, from_prospect, subject, content):
        exist_ = False
        existing_data = self.get_existing_emails_data()
        for camp_index in range(len(existing_data)):
            prosp_list = existing_data[camp_index]['prospect']
            mapped_camp_id = existing_data[camp_index]['Campaign id']
            mapped_from_prospect = self.isProspectFound(prosp_list, from_prospect)
            if current_camp_id == mapped_camp_id and mapped_from_prospect:
                print(self.isEmailSubjectFound(prosp_list, from_prospect, subject, content))
                exist_ = True
                break
        
        if not exist_:
            exist_ = False
        
        return exist_
        
