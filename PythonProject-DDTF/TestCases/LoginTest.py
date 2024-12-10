"""
LoginTest.py

Program : DDTF main executing file
"""

import pytest
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from Homepage import LoginPage

def read_test_data():
    return pd.read_excel('test_data.xlsx')

def write_test_result(index, result):
    df = pd.read_excel('test_data.xlsx')
    df.at[index, 'Test Result'] = result
    df.to_excel('test_data.xlsx', index=False)

@pytest.mark.parametrize('index, username, password',
                         [(i, row['Username'], row['Password']) for i, row in read_test_data().iterrows()])
def test_login(index, username, password):
    # Setup WebDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get('https://opensource-demo.orangehrmlive.com/web/index.php/auth/login')

    # Initialize LoginPage
    login_page = LoginPage(driver)

    # Perform login
    login_page.login(username, password)

    # Check if login was successful
    if login_page.is_login_successful():
        result = 'Passed'
    else:
        result = 'Failed'

    # Write the result to Excel
    write_test_result(index, result)

    # Cleanup
    driver.quit()

