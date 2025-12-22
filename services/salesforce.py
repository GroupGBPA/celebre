from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
import time
import os
from datetime import datetime
from dotenv import load_dotenv

# Gerando variaveis 
url = os.getenv('url_salesforce')
rpa_email = os.getenv("rpa_email")
rpa_password = os.getenv("rpa_password")

# Configurando o driver e abrindo o salesforce

options = Options()
options.add_argument("--start-maximized")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
driver.get(url)

