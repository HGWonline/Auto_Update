import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
from bs4 import BeautifulSoup
import re
from datetime import datetime

# Chrome 설정
chrome_service = Service(ChromeDriverManager().install())
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--window-size=1920,1080')

driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

def wait_for_page_load(timeout=10):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )

def login_and_navigate(driver):
    print("Navigating to login page...")
    driver.get('https://www.hangawee.com.au/login')

    username_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'username'))
    )
    print("Entering login code...")
    username_input.send_keys('245554073198')

    sign_in_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'kt_login_signin_submit'))
    )
    print("Clicking sign in button...")
    sign_in_button.click()

    # Outlet 선택
    booragoon_label = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//label[contains(., 'Booragoon')]"))
    )
    print("Selecting Booragoon outlet...")
    booragoon_label.click()

    next_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'next-step'))
    )
    print("Clicking Next button...")
    next_button.click()

    confirm_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'goIndex'))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", confirm_button)
    time.sleep(1)
    print("Clicking Confirm button...")
    confirm_button.click()

    # 홈 페이지의 특정 요소가 로드될 때까지 대기
    print("Waiting for home page to fully load...")
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'sales-volume-chart'))
        )
    except:
        print("Home page chart not found, proceeding anyway.")

    # item list 페이지로 이동
    print("Navigating to item list page...")
    driver.get('https://www.hangawee.com.au/?page=mitems')

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'button.btn.dropdown-toggle.btn-light'))
    )

    print("Setting item count to 700 using JavaScript...")
    dropdown_script = '''
    document.querySelector('button.dropdown-toggle').click();
    document.querySelectorAll('.dropdown-menu a')[5].click();
    '''
    driver.execute_script(dropdown_script)

    print("Waiting for the table to load with 700 items...")
    WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//table/tbody/tr')))

def go_to_next_page():
    try:
        print("Clicking next page button...")
        next_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.datatable-pager-link-next'))
        )
        next_button.click()

        print("Waiting for next page to load...")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//table/tbody/tr'))
        )
    except Exception as e:
        print(f"Error navigating to next page: {e}")
        driver.quit()

# 재고 정보 처리 함수 그대로 사용
def process_inventory_data(cell, discount_expiration=None, discount_price_global=None):
    stock, expiration, discount_price = None, None, None
    stock_info = cell.get('aria-label')
    if stock_info:
        discount_info = re.search(r'\$(\d+\.\d+)', stock_info)
        if discount_info:
            temp_discount_price = float(discount_info.group(1))
            expiration_stock_info = re.findall(r'(\d{2}/\d{2}/\d{4})\[(\d+)\]', stock_info)
            if expiration_stock_info:
                temp_expiration, temp_stock = expiration_stock_info[0]
                temp_stock = int(temp_stock)
            else:
                leading_stock_info = re.match(r'^(\d+)', stock_info)
                temp_stock = int(leading_stock_info.group(1)) if leading_stock_info else 0
                temp_expiration = None

            if 'MD:' in stock_info:
                temp_expiration = None

            if temp_stock > 0:
                stock = temp_stock
                expiration = temp_expiration
                discount_price = temp_discount_price
                discount_expiration = expiration
                discount_price_global = discount_price
            else:
                stock = temp_stock
        else:
            if discount_expiration and discount_price_global:
                expiration_stock_info = re.findall(r'(\d{2}/\d{2}/\d{4})\[(\d+)\]', stock_info)
                temp_stock = 0
                for exp_date, stk in expiration_stock_info:
                    if exp_date == discount_expiration:
                        temp_stock = int(stk)
                        break
                stock = temp_stock
                expiration = discount_expiration if temp_stock > 0 else None
                discount_price = discount_price_global if temp_stock > 0 else None
            else:
                leading_stock_info = re.match(r'^(\d+)', stock_info)
                stock = int(leading_stock_info.group(1)) if leading_stock_info else 0

    return stock or 0, discount_price, expiration, discount_expiration, discount_price_global

def scrape_data_and_save_to_excel():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Scraped Data"

    sheet.append(['Barcode', 'Description', 'CR_Stock', 'CR_Discount', 'CR_Expiration', 'NB_Stock', 'NB_Discount', 'NB_Expiration', 'IN_Stock', 'IN_Discount', 'IN_Expiration', 'BR_Stock', 'BR_Discount', 'BR_Expiration'])

    max_pages = 4
    for page_number in range(1, max_pages + 1):
        print(f"Scraping page {page_number}...")
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        rows = soup.select('table tbody tr')

        for row in rows:
            discount_expiration = None
            discount_price_global = None
            cells = row.find_all('td')
            if len(cells) > 6:
                barcode = cells[0].get_text(strip=True)
                description = re.sub(r'\s{2,}', ' ', cells[1].get_text(strip=True))

                cr_stock, cr_discount, cr_expiration, discount_expiration, discount_price_global = process_inventory_data(cells[2], discount_expiration, discount_price_global)
                nb_stock, nb_discount, nb_expiration, _, _ = process_inventory_data(cells[3], discount_expiration, discount_price_global)
                in_stock, in_discount, in_expiration, _, _ = process_inventory_data(cells[4], discount_expiration, discount_price_global)
                br_stock, br_discount, br_expiration, _, _ = process_inventory_data(cells[5], discount_expiration, discount_price_global)

                sheet.append([barcode, description, cr_stock, cr_discount, cr_expiration, nb_stock, nb_discount, nb_expiration, in_stock, in_discount, in_expiration, br_stock, br_discount, br_expiration])

        if page_number < max_pages:
            go_to_next_page()

    file_name = f"Scraped_{datetime.now().strftime('%y%m%d')}.xlsx"
    workbook.save(file_name)
    print(f"Data saved to '{file_name}'.")

# 실행
try:
    login_and_navigate(driver)
    scrape_data_and_save_to_excel()
finally:
    driver.quit()
