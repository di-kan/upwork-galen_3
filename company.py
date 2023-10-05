import requests
import shutil
import os
from urllib.parse import urlparse
from datetime import datetime


class Contact:
    def __init__(self):
        self.name = ""
        self.title = ""
        self.address = ""
    def __str__(self):
        return f"name:{self.name} title:{self.title} address:{self.address}"

class Document:
    def __init__(self):
        self.description = ""
        self.office_number = ""
        self.office_email = ""
        self.contacts = []

    def print(self):
        msg = f"Description: {self.description}\nemail: {self.office_email}\nTel:{self.office_number}"
        for contact in self.contacts:
            msg += f"\n\tName: {contact.name}\n\tTitle: {contact.title}\n\tAddr: {contact.address}"
        # print(msg)

class Filing:
    def __init__(self, pdf_id, document_id, filing_date, filing_type, document_type):
        self.pdf_id = pdf_id
        self.document_id = document_id
        self._filing_date = None
        self._set_filing_date(filing_date)
        self.filing_type = filing_type
        self.document_type = document_type
        self.document = None
        self.pdf_filename = ""
        self.pdf_url = ""

    def download_pdf(self):
        response = requests.get(self.pdf_url, stream=True)
        with open(self.pdf_filename, 'wb') as out_file:
            shutil.copyfileobj(response.raw, out_file)

    def delete_pdf(self):
        os.remove(self.pdf_filename)

    def upload_pdf(self, folder_id):
        pass

    def get_filename(url):
        parsed_url = urlparse(url)
        path = parsed_url.path
        filename = os.path.basename(path)
        return filename

    def find_url(self, index):
        ret_value = ""
        if self.pdf_id != "":
            try:
                # i_id = self.pdf_id.replace("bz", "z")
                headers = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/98.0.4758.102 Safari/537.36",
                           "Accept-Language": "en-US,en;q=0.9"}
                post_url = "https://www.sosnc.gov/online_services/imaging/download_ivault_pdf"
                data = {'ID': self.pdf_id}
                response = requests.post(post_url, data=data, headers=headers)
                return_data = response.json()

                if return_data.get('fileName') and len(return_data['fileName']) > 0:
                    ret_value = f"https://www.sosnc.gov/online_services/imaging/download/{return_data['fileName']}"
                    self.pdf_filename = f"{index+2}-{return_data['fileName']}.pdf"
            except Exception as e:
                # Log the error or take necessary actions
                # print("Error: 67", e)
                pass
        self.pdf_url = ret_value


    def __lt__(self, other):
        return self.filing_date < other.filing_date

    def _get_filing_date(self):
        return self._filing_date

    def _set_filing_date(self, filing_date):
        self._filing_date = datetime.strptime(filing_date, '%m/%d/%Y').date()
    filing_date = property(_get_filing_date, _set_filing_date)

    def __str__(self):
        return f"{self.filing_date}: {self.filing_type} ({self.document_id})"

class Company:
    def __init__(self, name, df_index=-1):
        self.df_index = df_index
        self.name = name
        self.link = ""
        self.status = ""
        self.types = ""
        self.company_id = ""
        self.annual_report = None
        self.total_results = 0
        self.results_index = 0
        self.dissolved_date = ""

    def __str__(self):
        return f"{self.name}|{self.total_results}|{self.status}|{self.dissolved_date}"
    
    def print(self):
        msg = f"{self.name}|{self.total_results}|{self.status}|{self.dissolved_date}"
        print(msg)

    def keep_last_annual_report(self, all_filings):
        all_filings.sort(reverse=True)
        for index, filing in enumerate(all_filings):
            if "annual " in filing.filing_type.lower():
                if "notice" not in filing.filing_type.lower():
                    found = index
                    self.annual_report = filing
                    self.annual_report.find_url(self.df_index)
                    break
