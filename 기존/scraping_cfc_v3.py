import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

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

# 데이터 스크래핑 및 Excel 파일로 저장
def scrape_data_and_save_to_excel():
    # Excel 파일 생성
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Scraped Data"

    # 헤더 작성
    sheet.append(['BarCode', 'Description', 'CAROUSEL', 'NORTHBRIDGE', 'INNALOO', 'BOORAGOON', 'CFC'])

    # 계속해서 페이지를 넘기며 데이터 스크래핑 (최대 4페이지)
    max_pages = 4  # 페이지 수 제한
    for page_number in range(1, max_pages + 1):
        print(f"Scraping page {page_number} data...")

        # 페이지의 모든 행 선택
        rows = driver.find_elements(By.XPATH, '//table/tbody/tr')

        # 각 행에서 데이터를 추출
        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, 'td')  # 각 행의 모든 셀 추출
                
                # 필요한 데이터가 있는 경우 추출
                if len(cells) > 6:  # 열이 충분히 있는지 확인
                    barcode = cells[0].text.strip()  # BarCode
                    description = cells[1].text.strip()  # Description
                    carousel = cells[2].text.strip()  # CAROUSEL
                    northbridge = cells[3].text.strip()  # NORTHBRIDGE
                    innaloo = cells[4].text.strip()  # INNALOO
                    booragoon = cells[5].text.strip()  # BOORAGOON
                    cfc = cells[6].text.strip()  # CFC

                    # 엑셀 파일에 한 줄씩 추가
                    sheet.append([barcode, description, carousel, northbridge, innaloo, booragoon, cfc])

            except Exception as e:
                print(f"Error scraping data from row: {e}")

        if page_number < max_pages:
            go_to_next_page()  # 다음 페이지로 이동

    # Excel 파일 저장
    workbook.save('scraped_data.xlsx')
    print("Data successfully saved to 'scraped_data.xlsx'.")

# 스크립트 실행
login_and_navigate()

# 데이터 스크래핑 및 Excel 저장 실행
scrape_data_and_save_to_excel()

# 브라우저 종료
driver.quit()
