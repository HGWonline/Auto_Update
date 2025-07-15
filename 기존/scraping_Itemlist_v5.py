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

# ===================== 1) 크롬 옵션 설정 =====================
chrome_options = Options()

# 최신 Headless 모드 (Chrome 109+) 
# 구버전 환경이면 '--headless' 그대로 쓰셔도 됩니다.
chrome_options.add_argument('--headless=new')

# 기초 옵션
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--start-maximized')  # 창 최대로
# 또는 아래 처럼 직접 크기 지정
# chrome_options.add_argument('--window-size=1920,1080')

# Anti-bot 감지 우회 (간단)
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
chrome_options.add_experimental_option('useAutomationExtension', False)

# ===================== 2) WebDriver 시작 =====================
chrome_service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

# 필요하다면 수동으로 창 크기 지정
driver.set_window_size(1920,1080)

def wait_for_page_load(timeout=20):
    """ 페이지 로드가 완료될 때까지 대기하는 함수 (타임아웃 20초) """
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )

def login_and_navigate(driver):
    print("Navigating to login page...")
    # Headless라 maximize_window() 무의미할 수 있으나, 일단 유지
    driver.maximize_window()

    driver.get('https://www.hangawee.com.au/login')
    wait_for_page_load()  # 페이지 로드 대기

    # username 필드
    username_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'username'))
    )
    print("Entering login code...")
    username_input.send_keys('245554073198')

    sign_in_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'kt_login_signin_submit'))
    )
    print("Clicking sign in button...")
    sign_in_button.click()

    # Booragoon 라벨
    booragoon_label = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.XPATH, "//label[contains(., 'Booragoon')]"))
    )
    print("Selecting Booragoon outlet...")
    # 필요하다면 JS로 클릭
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", booragoon_label)
    time.sleep(0.5)
    booragoon_label.click()

    next_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'next-step'))
    )
    print("Clicking Next button...")
    next_button.click()

    confirm_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'goIndex'))
    )
    print("Clicking Confirm button...")
    confirm_button.click()

    wait_for_page_load()
    print("Login and navigation completed successfully.")

    # 홈 화면 -> itemlist 페이지
    print("Navigating to item list page...")
    driver.get('https://www.hangawee.com.au/?page=mitems')
    wait_for_page_load()

    # 드롭다운 로드 확인
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'button.btn.dropdown-toggle.btn-light'))
        )
    except Exception as e:
        print("Dropdown button not found or took too long to appear.")
        raise e

    # JS로 항목 수를 700개로 설정
    print("Setting item count to 700 using JavaScript...")
    dropdown_script = """
    const ddBtn = document.querySelector('button.dropdown-toggle.btn-light');
    if(ddBtn) {
      ddBtn.click();
      const items = document.querySelectorAll('.dropdown-menu a');
      if(items && items[5]) {
        items[5].click();
      }
    }
    """
    driver.execute_script(dropdown_script)

    # 700개 항목이 로드되도록 대기
    print("Waiting for the table to load with 700 items...")
    time.sleep(2)  # JS가 반영될 시간 약간 확보
    WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.XPATH, '//table/tbody/tr'))
    )

def go_to_next_page():
    try:
        print("Clicking next page button...")
        # next 버튼 다시 locate
        next_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.datatable-pager-link-next'))
        )
        # 스크롤
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
        time.sleep(0.5)
        next_button.click()

        print("Waiting for next page to load...")
        WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.XPATH, '//table/tbody/tr'))
        )

    except Exception as e:
        print(f"Error navigating to next page: {e}")
        driver.quit()

def process_inventory_data(cell, discount_expiration=None, discount_price_global=None):
    """
    할인 정보와 재고 정보 추출 함수
    """
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
                if leading_stock_info:
                    temp_stock = int(leading_stock_info.group(1))
                    temp_expiration = None
                else:
                    temp_stock = 0
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
                expiration = None
                discount_price = None
        else:
            if discount_expiration and discount_price_global:
                expiration_stock_info = re.findall(r'(\d{2}/\d{2}/\d{4})\[(\d+)\]', stock_info)
                temp_stock = 0
                if expiration_stock_info:
                    for exp_date, stk in expiration_stock_info:
                        if exp_date == discount_expiration:
                            temp_stock = int(stk)
                            break
                else:
                    leading_stock_info = re.match(r'^(\d+)', stock_info)
                    if leading_stock_info:
                        temp_stock = int(leading_stock_info.group(1))

                if temp_stock > 0:
                    stock = temp_stock
                    expiration = discount_expiration
                    discount_price = discount_price_global
                else:
                    stock = temp_stock
                    expiration = None
                    discount_price = None
            else:
                leading_stock_info = re.match(r'^(\d+)', stock_info)
                if leading_stock_info:
                    stock = int(leading_stock_info.group(1))

    if stock is None:
        stock = 0

    return stock, discount_price, expiration, discount_expiration, discount_price_global

def scrape_data_and_save_to_excel():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Scraped Data"
    sheet.append([
        'Barcode', 'Description', 
        'CR_Stock', 'CR_Discount', 'CR_Expiration', 
        'NB_Stock', 'NB_Discount', 'NB_Expiration', 
        'IN_Stock', 'IN_Discount', 'IN_Expiration', 
        'BR_Stock', 'BR_Discount', 'BR_Expiration'
    ])

    max_pages = 4
    for page_number in range(1, max_pages + 1):
        print(f"Scraping page {page_number} data...")

        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')
        rows = soup.select('table tbody tr')

        for row in rows:
            discount_expiration = None
            discount_price_global = None
            cells = row.find_all('td')

            if len(cells) > 6:
                barcode = cells[0].get_text(strip=True)
                description = cells[1].get_text(strip=True)
                description = re.sub(r'\s{2,}', ' ', description)
                barcode = str(barcode)

                cr_stock, cr_discount, cr_expiration, discount_expiration, discount_price_global = process_inventory_data(cells[2], discount_expiration, discount_price_global)
                nb_stock, nb_discount, nb_expiration, _, _ = process_inventory_data(cells[3], discount_expiration, discount_price_global)
                in_stock, in_discount, in_expiration, _, _ = process_inventory_data(cells[4], discount_expiration, discount_price_global)
                br_stock, br_discount, br_expiration, _, _ = process_inventory_data(cells[5], discount_expiration, discount_price_global)

                sheet.append([
                    barcode, description, 
                    cr_stock, cr_discount, cr_expiration, 
                    nb_stock, nb_discount, nb_expiration, 
                    in_stock, in_discount, in_expiration, 
                    br_stock, br_discount, br_expiration
                ])

        if page_number < max_pages:
            go_to_next_page()
            # 페이지 전환 후 잠시 대기
            time.sleep(2)

    current_date = datetime.now().strftime('%y%m%d')
    file_name = f'Scraped_{current_date}.xlsx'
    workbook.save(file_name)
    print(f"Data successfully saved to '{file_name}'.")

# ============== 메인 실행 흐름 =================
login_and_navigate(driver)
scrape_data_and_save_to_excel()
driver.quit()
