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

# ChromeDriver 경로 설정 및 자동 설치
chrome_service = Service(ChromeDriverManager().install())

# 헤드리스 모드 활성화 (속도 향상을 위해)
chrome_options = Options()
chrome_options.add_argument('--headless')  # 헤드리스 모드 활성화
chrome_options.add_argument('--disable-gpu')  # GPU 비활성화 (Windows에서 필요)
chrome_options.add_argument('--window-size=1920x1080')  # 가상 브라우저 크기 설정

# WebDriver 시작 (service 및 options 분리)
driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

def wait_for_page_load(timeout=10):
    """ 페이지 로드가 완료될 때까지 대기하는 함수 """
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )

# 로그인 및 네비게이션
def login_and_navigate():
    print("Navigating to login page...")
    driver.maximize_window()  # 브라우저 전체화면으로 설정
    driver.get('https://www.hangawee.com.au/login')

    # 페이지가 완전히 로드될 때까지 대기
    wait_for_page_load()

    print("Entering login code...")
    code_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'username'))
    )
    code_input.send_keys('245554073198')

    # Sign in 버튼 클릭
    sign_in_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'kt_login_signin_submit'))
    )
    print("Clicking sign in button...")
    sign_in_button.click()

    # Next 버튼 클릭
    next_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'next-step'))
    )
    print("Clicking Next button...")
    next_button.click()

    # 팝업에서 Confirm 버튼 클릭
    confirm_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'goIndex'))
    )
    print("Clicking Confirm button on popup...")
    confirm_button.click()

    # 홈 화면으로 이동한 후, itemlist 페이지로 이동
    print("Navigating to item list page...")
    driver.get('https://www.hangawee.com.au/?page=mitems')

    # 드롭다운이 로드되었는지 확인
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'button.btn.dropdown-toggle.btn-light'))
    )

    # JavaScript로 항목 수를 700개로 설정
    print("Setting item count to 700 using JavaScript...")
    dropdown_script = '''
    document.querySelector('button.dropdown-toggle').click();
    document.querySelectorAll('.dropdown-menu a')[5].click();
    '''
    driver.execute_script(dropdown_script)

    # 700개 항목이 로드되도록 충분한 시간 대기
    print("Waiting for the table to load with 700 items...")
    WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//table/tbody/tr')))

# 다음 페이지로 이동
def go_to_next_page():
    try:
        print("Clicking next page button...")
        next_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.datatable-pager-link-next'))
        )
        next_button.click()

        print("Waiting for next page to load...")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//table/tbody/tr')))

    except Exception as e:
        print(f"Error navigating to next page: {e}")
        driver.quit()

# 전역 변수 설정
discount_expiration = None
discount_price_global = None

# 할인 정보와 재고 정보 추출 함수
def process_inventory_data(cell, discount_expiration=None, discount_price_global=None):
    stock, expiration, discount_price = None, None, None  # 기본값 설정

    # 1. aria-label 속성에서 정보를 모두 추출
    stock_info = cell.get('aria-label')
    if stock_info:
        # 3. 할인가 추출: $로 시작하는 숫자 추출
        discount_info = re.search(r'\$(\d+\.\d+)', stock_info)
        if discount_info:
            discount_price = float(discount_info.group(1))  # 할인가격을 숫자로 변환
            discount_price_global = discount_price  # 현재 매장의 할인가격

            # 할인 정보가 있는 매장의 유통기한과 재고 추출
            expiration_stock_info = re.findall(r'(\d{2}/\d{2}/\d{4})\[(\d+)\]', stock_info)
            if expiration_stock_info:
                expiration, stock = expiration_stock_info[0]
                stock = int(stock)
                discount_expiration = expiration  # 현재 매장의 유통기한
            else:
                # 유통기한과 재고 정보가 없을 때 총 재고 추출
                leading_stock_info = re.match(r'^(\d+)', stock_info)
                if leading_stock_info:
                    stock = int(leading_stock_info.group(1))

            # 'MD:'가 있는 경우 유통기한 표시 안 함
            if 'MD:' in stock_info:
                expiration = None

        else:
            # 할인 정보가 없는 매장 처리
            if discount_expiration and discount_price_global:
                expiration = discount_expiration
                discount_price = discount_price_global

                # 해당 유통기한의 재고 추출
                expiration_stock_info = re.findall(r'(\d{2}/\d{2}/\d{4})\[(\d+)\]', stock_info)
                stock = 0  # 기본값 설정
                if expiration_stock_info:
                    for exp_date, stk in expiration_stock_info:
                        if exp_date == expiration:
                            stock = int(stk)
                            break
            else:
                # 유통기한 없이 총 재고만 추출
                leading_stock_info = re.match(r'^(\d+)', stock_info)
                if leading_stock_info:
                    stock = int(leading_stock_info.group(1))

    # 5. 재고가 None이면 기본값 0으로 설정
    if stock is None:
        stock = 0

    return stock, discount_price, expiration, discount_expiration, discount_price_global

# BeautifulSoup을 사용한 데이터 스크래핑 및 Excel 파일로 저장
def scrape_data_and_save_to_excel():
    # Excel 파일 생성
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Scraped Data"

    # 헤더 작성
    sheet.append(['Barcode', 'Description', 'CR_Stock', 'CR_Discount', 'CR_Expiration', 'NB_Stock', 'NB_Discount', 'NB_Expiration', 'IN_Stock', 'IN_Discount', 'IN_Expiration', 'BR_Stock', 'BR_Discount', 'BR_Expiration'])

    # 계속해서 페이지를 넘기며 데이터 스크래핑 (최대 4페이지)
    max_pages = 4  # 페이지 수 제한
    for page_number in range(1, max_pages + 1):
        print(f"Scraping page {page_number} data...")

        # 페이지 소스를 BeautifulSoup으로 파싱
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')

        # 테이블 행을 파싱하여 데이터를 추출
        rows = soup.select('table tbody tr')
        for row in rows:
            # 각 상품마다 할인 정보 변수 초기화
            discount_expiration = None
            discount_price_global = None

            cells = row.find_all('td')

            # 필요한 데이터가 있는 경우 추출
            if len(cells) > 6:
                barcode = cells[0].get_text(strip=True)
                description = cells[1].get_text(strip=True)

                # 띄어쓰기 두 번을 한 번으로 교체
                description = re.sub(r'\s{2,}', ' ', description)

                # 바코드를 문자열로 변환
                barcode = str(barcode)  # 항상 문자열로 처리

                # 매장별 재고 및 할인 정보 추출
                cr_stock, cr_discount, cr_expiration, discount_expiration, discount_price_global = process_inventory_data(cells[2], discount_expiration, discount_price_global)
                nb_stock, nb_discount, nb_expiration, _, _ = process_inventory_data(cells[3], discount_expiration, discount_price_global)
                in_stock, in_discount, in_expiration, _, _ = process_inventory_data(cells[4], discount_expiration, discount_price_global)
                br_stock, br_discount, br_expiration, _, _ = process_inventory_data(cells[5], discount_expiration, discount_price_global)

                # 엑셀 파일에 한 줄씩 추가
                sheet.append([barcode, description, cr_stock, cr_discount, cr_expiration, nb_stock, nb_discount, nb_expiration, in_stock, in_discount, in_expiration, br_stock, br_discount, br_expiration])

        if page_number < max_pages:
            go_to_next_page()  # 다음 페이지로 이동

    # 현재 날짜 기준으로 파일 이름 작성 (yymmdd 형식)
    current_date = datetime.now().strftime('%y%m%d')
    file_name = f'Scraped_{current_date}.xlsx'

    # Excel 파일 저장
    workbook.save(file_name)
    print(f"Data successfully saved to '{file_name}'.")

# 스크립트 실행
login_and_navigate()

# 데이터 스크래핑 및 Excel 저장 실행
scrape_data_and_save_to_excel()

# 브라우저 종료
driver.quit()
