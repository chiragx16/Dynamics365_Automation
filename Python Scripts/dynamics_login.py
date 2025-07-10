from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import os
import configparser

# Load the configuration
config = configparser.ConfigParser()
config.read('__path__')   # config file path


# Read email and password from the [QA] section
qa_email = config.get('Prod', 'dynamics_email')
qa_password = config.get('Prod', 'dynamics_password')
login_web = config.get('Prod', 'login_web')


def dynamics_login():
    dfg = os.path.join(os.getcwd(), "__path__")  # chromedriver path
    service = Service(executable_path=dfg)

    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    options = Options()
    # options.set_preference('browser.sessionstore.resume_from_crash', False)
    #driver = webdriver.Firefox(options=options) # You need to have GeckoDriver instaclearlled
    # Login to CEIPAL
    driver.get(login_web)


    # Enter Email Address
    time.sleep(5)
    email_input = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.NAME, "loginfmt"))
    )
    email_input.send_keys(qa_email)
    email_input.send_keys(Keys.RETURN)


    # Enter Password
    time.sleep(5)
    pass_input = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.NAME, "passwd"))
    )
    pass_input.send_keys(qa_password)
    pass_input.send_keys(Keys.RETURN)


    # Clcik Yes if it Appears
    time.sleep(5)
    try:
        submit_button = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//input[@id="idSIButton9"]'))
        )
        submit_button.click()
    except:
        pass


    # Go to site
    time.sleep(5)
    return driver
