import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
from bs4 import BeautifulSoup
import re
from datetime import datetime

# ChromeDriver 경로 설정
chrome_driver_path = r'C:\Users\김남빈\chromedriver\chromedriver-win64\chromedriver.exe'
chrome_service = Service(chrome_driver_path)

# 헤드리스 모드 활성화 (속도 향상을 위해)
chrome_options = Options()
chrome_options.add_argument('--headless')  # 헤드리스 모드 활성화
chrome_options.add_argument('--disable-gpu')  # GPU 비활성화 (Windows에서 필요)
chrome_options.add_argument('--window-size=1920x1080')  # 가상 브라우저 크기 설정

# WebDriver 시작
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
    # 코드 입력 필드를 id로 찾기
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
        # 페이지 맨 아래에 도달한 후 다음 페이지 화살표 클릭
        next_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.datatable-pager-link-next'))
        )
        next_button.click()

        # 페이지 로딩 대기
        print("Waiting for next page to load...")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//table/tbody/tr')))

    except Exception as e:
        print(f"Error navigating to next page: {e}")
        driver.quit()

# 데이터 가공 함수
def process_inventory_data(text):
    # 총 재고 추출 (모든 숫자를 더하여 총 재고 계산)
    stocks = re.findall(r'\[(\d+)\]', text)
    total_stock = sum(map(int, stocks)) if stocks else 0

    # 할인율 및 할인가 추출
    discount_info = re.search(r'\[(.*?)\]:(\d+.\d+%) \$(\d+.\d+)', text)
    discount_rate = discount_price = None
    if discount_info:
        discount_rate = float(discount_info.group(2).replace('%', ''))  # 할인율을 숫자 형식으로 변환
        discount_price = float(discount_info.group(3))  # 할인가를 숫자 형식으로 변환

    return total_stock, discount_rate, discount_price

# BeautifulSoup을 사용한 데이터 스크래핑 및 Excel 파일로 저장
def scrape_data_and_save_to_excel():
    # Excel 파일 생성
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Scraped Data"

    # 헤더 작성
    sheet.append(['BarCode', 'Description', 'CAROUSEL Stock', 'NORTHBRIDGE Stock', 'INNALOO Stock', 'BOORAGOON Stock', 'CFC Stock', 'Discount Rate', 'Discount Price'])

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
            cells = row.find_all('td')

            # 필요한 데이터가 있는 경우 추출
            if len(cells) > 6:
                barcode = cells[0].get_text(strip=True)
                description = cells[1].get_text(strip=True)
                
                # 띄어쓰기 두 번을 한 번으로 교체
                description = re.sub(r'\s{2,}', ' ', description)

                # 바코드를 숫자 형식으로 변환
                if barcode.isdigit():
                    barcode = int(barcode)

                # 지역별 재고 정보 추출
                carousel_stock = process_inventory_data(cells[2].get_text(strip=True))[0]
                northbridge_stock = process_inventory_data(cells[3].get_text(strip=True))[0]
                innaloo_stock = process_inventory_data(cells[4].get_text(strip=True))[0]
                booragoon_stock = process_inventory_data(cells[5].get_text(strip=True))[0]
                cfc_stock = process_inventory_data(cells[6].get_text(strip=True))[0]

                # 할인 정보 추출
                _, discount_rate, discount_price = process_inventory_data(cells[5].get_text(strip=True))

                # 엑셀 파일에 한 줄씩 추가
                sheet.append([barcode, description, carousel_stock, northbridge_stock, innaloo_stock, booragoon_stock, cfc_stock, discount_rate, discount_price])

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
