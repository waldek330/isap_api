# Copyright (c) for Waldemar Łusiak NKS GROUP. All rights reserved.

import smtplib, ssl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import COMMASPACE, formatdate
from email import encoders
from email.message import EmailMessage
import time
import datetime
import selenium
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains as actions
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui as py

def send_email(dokumenty_nazwa, *args, **kwargs):
        #Email Variables
        SMTP_SERVER = '' #Email Server (don't change!)
        SMTP_PORT = 587 #Server Port (don't change!)
        GMAIL_USERNAME = '' #change this to match your gmail account
        GMAIL_PASSWORD = ''  #change this to match your gmail app-password
        
        class Emailer:
            def sendmail(self, recipient, subject, content):

                #Create Headers
                headers = ["From: " + GMAIL_USERNAME, "Subject: " + subject, "To: " + recipient,
                        "MIME-Version: 1.0", "Content-Type: text/html"]
                headers = "\r\n".join(headers)
                
                # Create the message
                message = EmailMessage()
                message['From'] = GMAIL_USERNAME
                message['To'] = recipient
                message['Subject'] = subject
                message.set_content(content)
                #message.add_attachment(attachment)

                #Connect to Gmail Server
                session = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
                session.ehlo()
                session.starttls()
                session.ehlo()

                #Login to Gmail
                session.login(GMAIL_USERNAME, GMAIL_PASSWORD)

                #Send Email & Exit
                session.send_message(message)
                session.quit

        sender = Emailer()

        sendTo = ''
        emailSubject = "Automatically generated email from Internetowy System Aktów Prawnych"
        emailContent = "Szukane dokumenty to: \n\n {}".format(dokumenty_nazwa)
        # 
        #Sends an email to the "sendTo" address with the specified "emailSubject" as the subject and "emailContent" as the email content.
        sender.sendmail(sendTo, emailSubject, emailContent)


"""
def run_code():       
    #dane do wyszukania po API
    publisher = "DU"
    years = (2020,2021,2022,2023)
    email_contents = []
    titles_dict = {}
    #pętla do wyciągania requestów
    for year in years:
        url = f"http://api.sejm.gov.pl/eli/acts/{publisher}/{year}"
        response = requests.get(url)
        if response.status_code == 200:
            data = json.loads(response.text.encode("utf-8", errors='ignore'))
            results = set()  # zmieniamy listę na zbiór
            for item in data['items']:
                if "ochrona środowiska" in item['title'] or "BHP" in item['title'] or 'ochronie środowiska' in item['title']:
                    item_tuple = tuple(item.items())  # zamieniamy słownik na krotkę
                    results.add(item_tuple)  # dodajemy krotkę do zbioru
                    for result in list(results):  # zamieniamy zbiór na listę
                        result_dict = dict(result)  # zamieniamy krotkę na słownik
                        dokumenty_nazwa = result_dict['title']  # pobieramy tytuł ze słownika przez klucz 'title'
                        if "Obwieszczenie" not in dokumenty_nazwa:
                            display_address = result_dict['displayAddress']
                            dokumenty_nazwa = f"Nr. {display_address}: {dokumenty_nazwa}"
                            if dokumenty_nazwa not in titles_dict.values(): 
                                titles_dict[display_address] = dokumenty_nazwa
                                email_contents.append(dokumenty_nazwa + "\n")
                    
    #złączenie dokumentów w jedno body dla emaila
    email_body = "\n".join(email_contents)                 
    send_email(email_body)
    return email_contents

                    

def main_code():
    while True:
        # Pobierz aktualny czas
        now = datetime.datetime.now()

        # Określ docelową godzinę i minutę
        target_hour = 24
        target_minute = 00
        
        # Jeśli aktualna godzina jest po docelowej godzinie lub aktualna godzina jest równa docelowej godzinie, ale aktualna minuta jest po docelowej minucie, 
        # ustaw docelową godzinę na następny dzień
        if (now.hour > target_hour) or (now.hour == target_hour and now.minute >= target_minute):
            target_date = now + datetime.timedelta(days=1)
        else:
            target_date = now

        # Ustaw docelową godzinę i minutę na podaną wartość
        target_date = target_date.replace(hour=target_hour, minute=target_minute, second=0, microsecond=0)

        # Oblicz różnicę pomiędzy aktualnym czasem a docelowym czasem i zamień ją na sekundy
        wait_time_seconds = (target_date - now).total_seconds()

        # Zatrzymaj program na podaną ilość sekund
        time.sleep(wait_time_seconds)
        
        run_code()
        
if __name__ == '__main__':
    main_code()
"""


def glpi_ticket_creation(email_contents):
    content = email_contents
    title = "ISAP- Internetowy System Aktów Prawnych, dokumenty do sprawdzenia"
    responsible_person = "Łusiak Waldemar"
    login_glpi_api = ""
    passw_glpi_api = ""
    url = "https://glpistg10.oneumbrella.pl/front/central.php"
    #prd = "https://glpi.oneumbrella.pl/front/central.php"
    try:
        options = Options()
        options.BinaryLocation = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        driver_path = r"C:\\Users\\waldemar.lusiak\Downloads\\chromedriver.exe"
        driver = webdriver.Chrome(options=options, service=Service(driver_path))
        driver.get(url)
        time.sleep(3)
        login_field = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[2]/div/form/div/div[1]/div[2]/input")
        login_field.send_keys(login_glpi_api)
        time.sleep(1)    
        password_field = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[2]/div/form/div/div[1]/div[3]/input")
        password_field.send_keys(passw_glpi_api)
        time.sleep(1)
        login_source = driver.find_element(By.CSS_SELECTOR, "#dropdown_auth1 > option:nth-child(1)")
        login_source.click()
        time.sleep(1) 
        login_button = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[2]/div/form/div/div[1]/div[6]/button")
        login_button.click()
        time.sleep(1) 
        assistance_form = driver.find_element(By.XPATH, "/html/body/div[2]/aside/div/div[2]/ul/li[2]/a")
        assistance_form.click()
        time.sleep(1) 
        forms_link = driver.find_element(By.XPATH, "/html/body/div[2]/aside/div/div[2]/ul/li[2]/div/div/div[2]/a/span")
        forms_link.click()
        time.sleep(1)
        ticket_link = driver.find_element(By.XPATH, "/html/body/div[2]/div/div/main/div[1]/div[2]/div/div[4]/div[1]/div[4]/a")
        ticket_link.click()
        time.sleep(1) 
        
        #ticket creation
        quest_title = driver.find_element(By.NAME, "formcreator_field_45")
        quest_title.send_keys(title)
        time.sleep(2)        
        responsible_title = driver.find_element(By.CSS_SELECTOR, "#form-group-field-46 > div > span > span.selection > span > ul > li > input")#CSS_SELECTOR, "#form-group-field-46 > div > span > span.selection > span > ul > li > input"
        responsible_title.send_keys(responsible_person)
        py.press('enter')
        time.sleep(4)
        task_detail = driver.find_element(By.CLASS_NAME, "tox-edit-area") #formcreator_field_50470413779_ifr
        driver.switch_to.frame(task_detail.find_element(By.CLASS_NAME, "tox-edit-area__iframe"))
        text_area_iframe = driver.find_element(By.ID, "tinymce")
        text_area_iframe.send_keys(content)
        driver.switch_to.default_content()
        time.sleep(2)
        send_task = driver.find_element(By.CSS_SELECTOR, "#plugin_formcreator_form > div.center > button")
        send_task.click()
        time.sleep(1500)
        driver.quit()
    except (Exception,StaleElementReferenceException) as e:
        raise(e)


email_contents = " "    
glpi_ticket_creation(email_contents)

"""

URL = 'https://glpistg10.oneumbrella.pl/marketplace/formcreator/front/formdisplay.php?id=16/apirest.php'
APPTOKEN = ''
USERTOKEN = ''

try:
    with glpi_api.connect(URL, APPTOKEN, USERTOKEN, verify_certs=False) as glpi:
        #print(glpi.get_config())
         # Utwórz ticket
        ticket_data = {
            'name': 'Internetowy System Aktów Prawnych- dodanie nowych dokumentów',  # tytuł ticketu
            'content': 'Treść ticketu',  # treść ticketu
            'itilcategories_id': 2,  # ID kategorii ticketu (jeśli jest wymagane)
            'urgency': 3  # pilność ticketu (od 1 do 5)
        }
        ticket = glpi.add('ticket', ticket_data)
        print(ticket)
        print(f'Utworzono ticket o ID {ticket["id"]}')
except glpi_api.GLPIError as err:
    print(str(err))

import requests

URL = "https://glpistg10.oneumbrella.pl/marketplace/formcreator/front/formdisplay.php?id=16/apirest.php"
APPTOKEN = ""
USERTOKEN = ""

try:
    headers = {
        "App-Token": APPTOKEN,
        "User-Token": USERTOKEN,
    }
    response = requests.get(URL, headers=headers, verify=False)
    soup = BeautifulSoup(response.text, "html.parser")
    text = soup.get_text()
    print(text)
except Exception as e:
    print(e)
"""
