import selenium.common.exceptions
# from gdrive import GoogleDrive
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service as ChromeService # Similar thing for firefox also!
# from subprocess import CREATE_NO_WINDOW # This flag will only be available in windows
import pandas as pd
from random import choice
import re
from time import sleep
from unicodedata import normalize
from bs4 import BeautifulSoup
import time
import logging
import os
from threading import Thread, Lock, active_count
import urllib.request
import logging
from company import Company, Filing, Document, Contact


class Browsers:
    def __init__(self, ports):
        print(ports)
        self.threads = len(ports)
        self.ports = ports
        self.drivers = [self._launch_selenium(port) for port in self.ports]
        self.free = [driver for driver in self.drivers]
        print(len(self.drivers))
        print(len(self.free))
        # self.print_status("init")

    def _launch_selenium(self, port):
        options = webdriver.ChromeOptions()
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
        # options.page_load_strategy = 'eager'
        options.add_argument('--blink-settings=imagesEnabled=false')
        # options.add_argument('--headless')
        # options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_service = ChromeService('chromedriver')
        # chrome_service.creation_flags = CREATE_NO_WINDOW
        msg = f"\t Launching browser ..."
        logging.info(msg)
        return webdriver.Chrome(options=options)
        # return webdriver.Chrome(options=options, service=chrome_service)

    def close_all(self):
        for driver in self.drivers:
            driver.close()
        # self.print_status("close_all")

    def get_one(self):
        the_driver = choice(self.free)
        self.free.remove(the_driver)
        # self.print_status("get_one")
        return the_driver

    def release(self, the_driver):
        self.free.append(the_driver)
        # self.print_status("release")

    def free(self):
        ret_value = True if len(self.free) > 0 else False
        # self.print_status("free")
        return ret_value

    def print_status(self, fun):
        # print(f"{fun} -->  Free/Total browsers:{len(self.free)}/{len(self.drivers)}")
        pass


class Scraper:
    def __init__(self):
        self.url = "https://ecorp.sos.ga.gov/BusinessSearch"
        self._download_folder = ""
        self._excel_filename = ""
        self.browsers = None
        self.other_threads = 0
        # self.MAX_OTHER = 2
        self.finished = False
        self.current = 0
        self.processing_totals = 0
        self.time_per_company = 0
        self.original_df = None
        self.companies = {}
        # self.gdrive = GoogleDrive()
        self.lock = None
        self.frame = None
        self.working = False
        msg = f"\t Launching Scraper ..."
        logging.info(msg)

    def am_i_connected(self):
        host = 'https://sosnc.gov/'
        try:
            urllib.request.urlopen(host)  # Python 3.x
            return True
        except Exception as e:
            msg = f"Connection Error:{host}, {e}"
            # print(msg)
            logging.error(msg)
            return False

    def _get_excel_filename(self):
        return self._excel_filename

    def _set_excel_filename(self, excel):
        self._excel_filename = excel
        self.original_df = pd.read_excel(excel)
        totals_rows=self.original_df.shape[0]
        msg = self.original_df.shape
        logging.info(msg)
        null_list = ["" for _ in range(totals_rows)]
        if "status" not in self.original_df:
            self.original_df.insert(1, 'status', null_list)
            self.original_df.insert(2, 'chosen_result', null_list)
            self.original_df.insert(3, 'total_results', null_list)
    excel_filename = property(_get_excel_filename, _set_excel_filename)

    def _get_download_folder(self):
        return self._download_folder

    def _set_download_folder(self, download_folder):
        if download_folder == "":
            download_folder = "./"
        self._download_folder = download_folder
        if not os.path.exists(download_folder):
            os.mkdir(download_folder)
    download_folder = property(_get_download_folder, _set_download_folder)

    def _get_company_object(self, index):
        """Instantiates a company objec ued by the scrapper object"""
        return Company(name=self.original_df.iloc[0]['Owner Name'], df_index=index)

    def _find_best_name(self, company_name, companies_list):
        company_name = company_name.lower()
        many_names = [a_name.name.lower() for a_name in companies_list]
        many_names = [re.sub("prev legal.*", "", a_name) for a_name in many_names]
        company_name_parts = company_name.split(" ")
        len_names = []
        for idx in range(len(many_names)):
            cnt = 0
            aname = many_names[idx]
            #Check if the type of the company is of the same kind. Else we shouldn't be bothered with this company
            if (("llc" in company_name.lower()) and ("limited liability company" in companies_list[idx].types.lower()) or
                (("lp" in company_name.lower()) and ("limited partnerships" in companies_list[idx].types.lower())) or
                (("inc" in company_name.lower()) and ("business corporation" in companies_list[idx].types.lower()))):
                for part in company_name_parts:
                    if part != "llc" and part != "lp" and part != "inc":
                        if part in aname:
                            cnt += 1
                remaining_chars = 10000
                #if all parts of the input string are located in the search result then move on
                if cnt == len(company_name_parts)-1:
                    for part in company_name_parts:
                        aname = re.sub(part, "", aname)
                        remaining_chars = len(aname)
            else:
                remaining_chars = 10000
            len_names.append(remaining_chars)
        best = min(len_names)
        return(len_names.index(best))

    def _get_all_companies(self, soup):
        """ Returns a list of Company objects that the search page returned"""
        results = []
        table_rows = soup.find("div", class_="data_pannel").find("tbody").findAll("tr")
        for row in table_rows:
                cells = row.findAll("td")
                company_name = cells[0].getText().strip()
                company_id = cells[1].getText().strip()
                types = cells[2].getText().strip()
                status = cells[5].getText().strip()
                a_Company = Company(company_name)
                a_Company.company_id = company_id
                a_Company.status = status
                a_Company.types = types
                results.append(a_Company)
        return results

    def process_companies(self, start, stop, ports):
        self.finished = False
        self.working = True
        # Get count of non-browser related threads
        self.other_threads = active_count()
        self.browsers = Browsers(ports)

        threads = []
        time1 = time.time()
        self.lock = Lock()
        self.processing_totals = stop-start+1
        # print(f"starting with {threading.active_count()} threads")
        i = start
        while self.working:
            print(f"processing row {i}")
            self.current = i - start + 1
            driver = self.browsers.get_one()
            self.companies[driver] =  self._get_company_object(i)
            # THREADING
            x = Thread(target=self.process_company, args=(driver,))
            threads.append(x)
            x.start()
            
            i += 1
            if i == stop+1:
                self.working = False
            print("end of processing")


            # if the total number of threads is MAX, wait. None have finished yet. Else continue
            # while threading.active_count() >= self.browsers.threads + self.other_threads:
            while not self.browsers.free:
            #     # print(f"No browsers are free {len(self.browsers.free)}/{len(self.browsers.drivers)}")
            #     # print(f"Thread count is {threading.active_count()} >= "
            #     #       f"{self.MAX_DOWN_THREADS+self.OTHER_THREADS}. "
            #     #       f"Free drivers:{len(self.free_drivers)} .Sleeping and checking again...")
                sleep(0.5)
        while active_count() > self.other_threads:
            # print(f"Waiting for all threads to finish {threading.active_count()-self.OTHER_THREADS}. Sleeping and checking again...")
            sleep(0.5)

        self.frame.stop_from_engine()
        time2 = time.time()
        self.finished = True
        self.time_per_company = (time2-time1)/(stop-start)
        msg = f"Download time per company: {(time2-time1)/(stop-start)}"
        logging.info(msg)

    def process_company(self, the_driver):
        msg = f"Searching for company {self.companies[the_driver].name}..."
        if self.results_for_company(the_driver):
            # Choose the best company out of the results
            totals = self.choose_company(the_driver)
            self.companies[the_driver].print()
            msg = "Saving..."
            logging.info(msg)
            self.save_company_to_excel(the_driver)
            print("Releasing browser ...\n-----------------------------------------------------------")
        self.browsers.release(the_driver)

    def results_for_company(self, the_driver):
        """Navigates the assigned browser to the results for the copmany name"""
        result = False
        the_driver.get(self.url)
        elem = the_driver.find_element(By.NAME, "txtBusinessName")
        elem.clear()
        elem.send_keys(self.companies[the_driver].name)
        elem.send_keys(Keys.RETURN)
        try:
            element = WebDriverWait(the_driver, 20).until(EC.any_of(       
                EC.presence_of_element_located((By.ID, "businessSearchResult")),
                EC.presence_of_element_located((By.ID, "errorDialog"))
            ))
        except Exception as error:
            msg = f"Error for company: {self.companies[the_driver].name} --> {error}"
            logging.error(msg)
            result = False
        else:
            which_element = element.get_attribute("id")
            match which_element:
                case "businessSearchResult":
                    result = True
                case "errorDialog":
                    result = False
        return result

    def choose_company(self, the_driver):
        """Chooses the best company and modifies self.companies[the_driver] to current state (set name, index, totals"""
        soup = BeautifulSoup(the_driver.page_source, "lxml")
        totals = 0
        chosen_index = -1
        id = ""
        status = ""
        types = ""
        # Setting default company name
        final_company_name = self.companies[the_driver].name
        print(final_company_name)
        try:
            # Get all the companies that the search page reported
            companies_list = self._get_all_companies(soup)
            totals = len(companies_list)
            msg = f"Totals:{totals}"
            logging.info(msg)
            match totals:
                case 0:
                    chosen_index = -1
                case 1:
                    # Results returned 1 company. Using this
                    chosen_index = 0
                    final_company_name = companies_list[chosen_index].name
                    id = companies_list[chosen_index].company_id
                    status = companies_list[chosen_index].status
                    types = companies_list[chosen_index].types
                case _:
                    # From all results choose the best company
                    chosen_index = self._find_best_name(self.companies[the_driver].name, companies_list)
                    final_company_name = companies_list[chosen_index].name
                    id = companies_list[chosen_index].company_id
                    status = companies_list[chosen_index].status
                    types = companies_list[chosen_index].types

        except Exception as e:
            msg = f"Error 333:{self.companies[the_driver].name}, index:{self.companies[the_driver].df_index:04}, {e}"
            logging.error(msg)
        finally:
            self.companies[the_driver].name = final_company_name
            self.companies[the_driver].company_id = id
            self.companies[the_driver].results_index = chosen_index
            self.companies[the_driver].total_results = totals
            self.companies[the_driver].status = status
            self.companies[the_driver].types = types
            return totals

    def save_company_to_excel(self, the_driver):
        msg = self.original_df.shape
        logging.info(msg)
        i = self.companies[the_driver].df_index
        with self.lock:
            self.original_df.at[i, 'status'] = self.companies[the_driver].status
            self.original_df.at[i, 'chosen_result'] = self.companies[the_driver].name
            self.original_df.at[i, 'total_results'] = self.companies[the_driver].total_results
            self.original_df.to_excel(self._excel_filename, index=False)
