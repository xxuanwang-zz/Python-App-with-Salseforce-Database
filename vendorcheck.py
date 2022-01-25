import os
import sys
import time
from pathlib import Path 
import concurrent.futures

import json
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from simple_salesforce import Salesforce
import pandas as pd

import urllib.request
from os.path import basename

MAX_THREADS = 25

class Automated_Vendor_Check:
    
    def __init__(self, path):
        self.path = path

    def help(self):
        print("Supported Checks - Please check your default download folder")
        print("1. Franchise Tax Status")
        print("2. CMBL/HUB Status")
        print("3. Vendor Performance Search")
        print("4. SAM Search")
        print("5. OFAC Search")
        print("6. Debarred Vendor List")
        print("7. Divestment Statute List")

    def _error_messages(self, issue):
        error_message_dict = {}
        error_message_dict['err_folder'] = 'An error occured when creating the folder'
        error_message_dict['err_debarred'] = 'Debarred Vendor List - failed'
        error_message_dict['err_div'] = 'Divestment Statute List - failed'
        error_message_dict['err_vp'] = 'Vendor Performance Search - failed'
        error_message_dict['unknown_error'] = 'Something was not correct with the request. Try again.'

        if issue:
            return error_message_dict[issue]
        else:
            return error_message_dict['unknown_error']

    def _notifications(self, notice):
        notice_dict = {}
        notice_dict['sam_pass'] = "SAM Search passed"
        notice_dict['ofac_pass'] = "OFAC Search passed"
        notice_dict['hub_pass'] = "CMBL/HUB Status Search passed"
        notice_dict['fts_pass'] = "Franchise Tax Status Search passed"
        notice_dict['div_pass'] = "Divestment Statute List Search passed"
        notice_dict['debarred_pass'] = "Debarred Vendor List Search passed"
        notice_dict['vp_pass'] = "Vendor Performance Search passed"

        notice_dict['folder_exists'] = "Folder already exists"
        notice_dict['unknown_notice'] = "Something was not correct..."
        notice_dict['not_found'] = 'No files found'

        if notice:
            return notice_dict[notice]
        else: 
            return notice_dict['unknown_notice']

    def get_query(self, VendorName='', VendorId=''):
        try:
            if VendorName != '' and VendorId == '':
                # Single search, vendor_name
                sql_name = "WHERE Name = " + "'" + VendorName + "'"
                salesforce_data = sf.query_all("SELECT Name, Vendor_Id__c, DUNS_Number__c FROM Account" + " " + sql_name)

                df = pd.DataFrame(salesforce_data['records']).drop(columns = 'attributes')
                vendor_name = df.loc[0, 'Name']
                vendor_id = df.loc[0, 'Vendor_Id__c']
                duns_num = df.loc[0, 'DUNS_Number__c']

                print("Search results for - %s" % VendorName)
                print("Vendor id: %s" % vendor_id)
                print("DUNS_Number: %s" % duns_num)

                return vendor_name, vendor_id, duns_num

            elif VendorId != '' and VendorName == '':
                # Single search, vendor_id
                sql_id = "WHERE Vendor_Id__c = " + "'" + VendorId + "'"
                salesforce_data = sf.query_all("SELECT Name, Vendor_Id__c, DUNS_Number__c FROM Account" + " " + sql_id)

                df = pd.DataFrame(salesforce_data['records']).drop(columns = 'attributes')
                vendor_name = df.loc[0, 'Name']
                vendor_id = df.loc[0, 'Vendor_Id__c']
                duns_num = df.loc[0, 'DUNS_Number__c']

                print("Search results for - %s" % VendorId)
                print("Vendor Name: %s" % vendor_name)
                print("DUNS_Number: %s" % duns_num)

                return vendor_name, vendor_id, duns_num

            elif VendorName == '' and VendorId == '':
                # Get the info of all vendors, and save into csv file
                salesforce_data = sf.query_all("SELECT Name, Vendor_Id__c, DUNS_Number__c FROM Account")  
                df_vendor = pd.DataFrame(salesforce_data['records']).drop(columns = 'attributes')
                return df_vendor

        except Exception as e:
            print("Query failed...")
            return e

    def make_dir(self, folder):
        des_dir = os.path.join(os.path.expanduser("~"), "Downloads", folder)
        try:
            if not os.path.isdir(des_dir):
                os.makedirs(des_dir)
            return des_dir

        except Exception as e:
            return e

    def download_files(self, file_url, des_dir):
        response = urllib.request.urlopen(file_url)
        filename = basename(response.url)
        file_path = os.path.join(os.path.expanduser("~"), des_dir, filename)
        file = open(file_path, 'wb')
        file.write(response.read())
        file.close()

    def download_pdfs(self, pdf_url, des_dir):
        threads = min(MAX_THREADS, len(pdf_url))
        with concurrent.futures.ThreadPoolExecutor(max_workers=threads) as executor:
            executor.map(self.download_files(pdf_url, des_dir))

    def Debarred_List(self):
        pdf_url = "https://comptroller.texas.gov/purchasing/docs/debarred-vendor-list.pdf"
        pdf_name = "debarred-vendor-list.pdf"
        des_dir = str(self.make_dir("Debarred Vendor"))
        try:
            if os.path.exists(des_dir):
                self.download_pdfs(pdf_url, des_dir)

            if pdf_name in os.listdir(des_dir):
                return self._notifications('debarred_pass')
            else:
                return self._error_messages('err_debarred')

        except Exception as e:
            return e

    def Divestiment(self):
        files_list = ["anti-bds.pdf", "sudan-list.pdf", "iran-list.pdf", "foreign-terrorist.pdf", "fto-list.pdf"]
        des_dir = str(self.make_dir("Divestment Statute"))
        pdf_urls = [
            "https://comptroller.texas.gov/purchasing/docs/anti-bds.pdf",
            "https://comptroller.texas.gov/purchasing/docs/sudan-list.pdf", 
            "https://comptroller.texas.gov/purchasing/docs/iran-list.pdf", 
            "https://comptroller.texas.gov/purchasing/docs/foreign-terrorist.pdf", 
            "https://comptroller.texas.gov/purchasing/docs/fto-list.pdf"]
        try:
            for url in pdf_urls:
                self.download_pdfs(url, des_dir)

            time.sleep(2)
            pdf_list = os.listdir(des_dir)
            if len(pdf_list) == 5:
                if set(files_list) == set(pdf_list):
                    return self._notifications('div_pass')
                else:
                    return self._error_messages('err_div')
        except Exception as e:
            return e

    def SAM_Check(self, VendorName='', Duns_Num=''):
            sam_search = "https://www.sam.gov/SAM/pages/public/searchRecords/search.jsf"
            des_dir = os.path.join(os.path.expanduser("~"), "Downloads", "SAM Search")
            options = webdriver.ChromeOptions()
            options.add_argument("--headless")
            prefs = {"download.default_directory" : des_dir}
            options.add_experimental_option("prefs", prefs)
            options.add_experimental_option("excludeSwitches", ['enable-logging'])

            driver = webdriver.Chrome(executable_path = self.path, options = options)
            driver.get(sam_search)

            # use css selectors to grab the search inputs
            if VendorName:
                driver.find_element_by_id("searchBasicForm:qterm_input").send_keys(VendorName)
                # Find "Search" button and click
                driver.find_element_by_name("searchBasicForm:SearchButton").click()
                driver.find_element_by_xpath('//*[@id="samContentForm"]/div/table/tbody/tr/td/table/tbody/tr[1]/td/table[1]/tbody/tr/td[3]/input[1]').click()

            elif Duns_Num:
                driver.find_element_by_id("searchBasicForm:DUNSq").send_keys(Duns_Num)
                # Find "Search" button and click
                driver.find_element_by_name("searchBasicForm:SearchButton").click()
                driver.find_element_by_xpath('//*[@id="samContentForm"]/div/table/tbody/tr/td/table/tbody/tr[1]/td/table[1]/tbody/tr/td[3]/input[1]').click()

            try:
                time.sleep(2)
                file_path = os.path.join(des_dir, "searchResults.pdf")
                time.sleep(2)
                if os.path.exists(file_path):
                    driver.close()
                    return self._notifications('sam_pass') 
                else:
                    return ("An error occured when downloading - %s" % VendorName)
            except Exception as e:
                return e

    def OFAC_Search(self, VendorName):
        
        ofac_search = "https://sanctionssearch.ofac.treas.gov/"
        des_dir = os.path.join(os.path.expanduser("~"), "Downloads", "OFAC Search")
        settings = {
            "recentDestinations": [{
                    "id": "Save as PDF",
                    "origin": "local",
                    "account": "",
                }],
                "selectedDestinationId": "Save as PDF",
                "version": 2
            }

        options = webdriver.ChromeOptions()
        prefs = {
            'printing.print_preview_sticky_settings.appState': json.dumps(settings)
            }

        options.add_experimental_option('prefs', prefs)
        options.add_argument('--kiosk-printing')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(executable_path = self.path, options = options)

        driver.get(ofac_search)
        time.sleep(2)
        driver.find_element_by_id("ctl00_MainContent_txtLastName").send_keys(VendorName)
        driver.find_element_by_id("ctl00_MainContent_btnSearch").click()
        time.sleep(2)
        try:
            span_text = driver.find_element_by_id("ctl00_MainContent_lblResults").text
            if "0 Found" in span_text:
                time.sleep(3)
                driver.execute_script('window.print();')
                driver.close()
                return self._notifications('not_found')
            elif "1 Found" in span_text:
                element_img = WebDriverWait(driver, 10).\
                    until(EC.element_to_be_clickable((((By.ID, "ctl00_MainContent_ImageButton1")))))
                element_img.click()

                file_path = os.path.join(des_dir, "Search_Results.xls")
                time.sleep(3)
                if os.path.exists(file_path):
                    return self._notifications('ofac_pass')
                else:
                    driver.close()
                    return ("An error occured when downloading - %s" % VendorName)
        except Exception as e:
            return e

    def Vendor_Performance(self, VendorId, VendorName):

        vendor_performance_check = "http://www.txsmartbuy.com/vpts"
        settings = {
            "recentDestinations": [{
                    "id": "Save as PDF",
                    "origin": "local",
                    "account": "",
                }],
                "selectedDestinationId": "Save as PDF",
                "version": 2
            }

        options = webdriver.ChromeOptions()
        prefs = {
            'printing.print_preview_sticky_settings.appState': json.dumps(settings)
            }

        options.add_experimental_option('prefs', prefs)
        options.add_argument('--kiosk-printing')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(executable_path = self.path, options = options)
        driver.get(vendor_performance_check)
        time.sleep(8)

        if VendorId:
            driver.find_element_by_id("vendorIDSearch").send_keys(VendorId)
            # Find button to save the results
            driver.find_element_by_id("vprBtnSearch").click()
            time.sleep(10)
            driver.execute_script('window.print();')
            driver.quit()
            return self._notifications('vp_pass')

        elif VendorName:
            driver.find_element_by_id("vendorNameSearch").send_keys(VendorName)
            # Find button to save the results
            driver.find_element_by_id("vprBtnSearch").click()
            time.sleep(10)
            driver.execute_script('window.print();')
            driver.quit()
            return self._notifications('fts_pass')
        else:
            return self._error_messages('err_vp')

    def Franchise_Tax_Status(self, VendorName, VendorId):
        
        des_dir = os.path.join(os.path.expanduser("~"), "Downloads", "Franchise Tax Status")
        options = webdriver.ChromeOptions()
        franchise_tax_status = "https://data.texas.gov/dataset/Active-Franchise-Tax-Permit-Holders/9cir-efmm/data"
        prefs = {"download.default_directory" : des_dir}
        options.add_argument("--headless")
        options.add_experimental_option('prefs', prefs)
        options.add_argument('--kiosk-printing')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        driver = webdriver.Chrome(executable_path = self.path, options = options)
        driver.get(franchise_tax_status)
        if VendorId:
            driver.find_element_by_id("searchField").send_keys(VendorId)
        elif VendorName:
            driver.find_element_by_id("searchField").send_keys(VendorName)

        time.sleep(3)
        driver.find_element_by_id("searchField").send_keys(Keys.RETURN)
        time.sleep(3)

        element_exp = WebDriverWait(driver, 8).\
            until(EC.element_to_be_clickable((((By.XPATH, '//*[@id="sidebarOptions"]/li[6]/a')))))
        element_exp.click()

        element_csv = WebDriverWait(driver, 8).\
            until(EC.element_to_be_clickable((((By.XPATH, '//*[@id="controlPane_downloadDataset_3"]/form/div[3]/div[1]/div[4]/div/div/table/tbody/tr[1]/td/div/a')))))
        
        time.sleep(3)
        element_csv.click()
        time.sleep(3)
        driver.close()
        return self._notifications('fts_pass')

    def HUB_Status(self, VendorId, VendorName):
        hub_status = "https://mycpa.cpa.state.tx.us/tpasscmblsearch/tpasscmblsearch.do"
        settings = {
            "recentDestinations": [{
                    "id": "Save as PDF",
                    "origin": "local",
                    "account": "",
                }],
                "selectedDestinationId": "Save as PDF",
                "version": 2
            }

        options = webdriver.ChromeOptions()
        prefs = {
            'printing.print_preview_sticky_settings.appState': json.dumps(settings)
            }

        options.add_experimental_option('prefs', prefs)
        options.add_argument('--kiosk-printing')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(executable_path = self.path, options = options)
        driver.get(hub_status)

        element_single = WebDriverWait(driver, 10).\
            until(EC.element_to_be_clickable((((By.LINK_TEXT, "SINGLE VENDOR SEARCH")))))
        element_single.click()

        if VendorName:
            time.sleep(1)
            driver.find_element_by_id("vendorName").send_keys(VendorName)
        else:
            time.sleep(1)
            driver.find_element_by_id("vendorId").send_keys(VendorId)

        driver.find_element_by_id("inclInactiveVndrs").click()

        element_search = WebDriverWait(driver, 10).\
            until(EC.element_to_be_clickable((((By.ID, "search")))))
        element_search.click()
        time.sleep(1)
        
        try:
            moreId = VendorId + '00'
            element_details = WebDriverWait(driver, 10).\
                until(EC.element_to_be_clickable((((By.LINK_TEXT, moreId)))))
            element_details.click()

            time.sleep(5)
            driver.execute_script('window.print();')
            driver.close()
            return self._notifications('hub_pass')

        except Exception as e:
            time.sleep(3)
            driver.execute_script('window.print();')
            driver.close()
            return self._notifications('not_found')

    def start(self, VendorName, VendorId, Duns_Num=''):
        print("Start Checking-------------------")
        time.sleep(2)
        try:
            print("\nFranchise Tax Status Check")
            ret1 = self.Franchise_Tax_Status(VendorName, VendorId)
            print(ret1)
            print("\nVendor Performance Check")
            ret3 = self.Vendor_Performance(VendorId, VendorName)
            print(ret3)
            print("\nSAM Check")          
            ret4 = self.SAM_Check(VendorName, Duns_Num)
            print(ret4)
            print("\nOFAC Search Check")
            ret5 = self.OFAC_Search(VendorName)
            print(ret5)
            print("\nDebarred List")
            ret6 = self.Debarred_List()
            print(ret6)
            print("\nDivestiment List")
            ret7 = self.Divestiment()
            print(ret7)
            print("\nHUB Status Check")
            ret2 = self.HUB_Status(VendorId, "")
            if ret2.lower() == "no files found":
                print("For this vendor id")
                print(ret2)
                ret22 = self.HUB_Status(VendorId, VendorName)
                print("For this vendor name")
                print(ret22)
            else:
                print(ret2)

        except Exception as e:
            raise e
        print("\nFinish Checking------------------")

def main():
    
    if len(sys.argv) < 1:
        print("python scraping.py\n")
        exit(0)
    
    VendorName = input("Please enter a vendor name>")
    VendorId = input("Please enter a vendor id>")
    
    # E.g. 'C:\Program Files\Google\Chrome\Application\chromedriver.exe'
    chrome_path = input("Full path to your chromedriver>")
    chrome_Path = Path(chrome_path)
    
    vendor_check = Automated_Vendor_Check()

    if VendorName != '' or VendorId != '':
        try:
            vendor_name, vendor_id, duns_num = vendor_check.get_query(VendorName, VendorId)

        except Exception as e:
            print("Please enter a vaild vendor name and id")
            return e

    elif VendorName == '' and VendorId == '':
        print(vendor_check.get_query())
        print("At least one input is required to complete the query")
        exit()

    try:
        vendor_check.help()
        response = vendor_check.start(vendor_name, vendor_id, duns_num)
    except Exception as e:
        print(e)

if __name__ == '__main__':
    
    # Authorization
    # Connect to the inner database using username, password, domain, and security token
    print("Please first connect to Salesforce")
    UserName = input("Username>")
    Password = input("Password>")
    SecurityToken = input("Security token>")
    Domain = input("Domain>")

    try:
        sf = Salesforce(
            username = UserName,
            password = Password,
            security_token = SecurityToken,
            domain = Domain)
        print("Successfully connected to Salesforce!") 

    except Exception as e:
        print("Connection failed...")
        print(e)
    
    main()
