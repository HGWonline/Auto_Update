import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# ChromeDriver 경로 설정
chrome_driver_path = r'C:\Users\김남빈\chromedriver\chromedriver-win64\chromedriver.exe'
chrome_service = Service(chrome_driver_path)

# 헤드리스 모드 설정
chrome_options = Options()
chrome_options.add_argument('--headless')  # 헤드리스 모드
chrome_options.add_argument('--disable-gpu')  # GPU 비활성화 (Windows에서 필요)
chrome_options.add_argument('--window-size=1920x1080')  # 가상 브라우저 크기 설정

# 재시도 로직 추가
def start_driver_with_retries(retry_count=3):
    attempt = 0
    while attempt < retry_count:
        try:
            driver = webdriver.Chrome(service=chrome_service)
            return driver
        except WebDriverException as e:
            print(f"Attempt {attempt + 1}/{retry_count} - WebDriverException: {e}")
            attempt += 1
            time.sleep(3)  # 재시도 전 잠시 대기
    raise WebDriverException("Failed to start WebDriver after several attempts")

driver = start_driver_with_retries()

# 로그인 및 네비게이션
def login_and_navigate():
    driver.maximize_window()  # 브라우저 전체화면으로 설정
    driver.get('https://www.hangawee.com.au/login')

    # 코드 입력 필드를 XPath로 찾기
    try:
        code_input = driver.find_element(By.XPATH, '//input[@type="text"]')
        code_input.send_keys('245554073198')
    except Exception as e:
        print(f"Error finding the code input field: {e}")
        driver.quit()
        return

    # Sign in 버튼 클릭
    try:
        sign_in_button = driver.find_element(By.XPATH, '//button[contains(text(), "Sign In")]')
        sign_in_button.click()
    except Exception as e:
        print(f"Error finding the sign-in button: {e}")
        driver.quit()
        return

    # Next 버튼 클릭 (select 이후)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Next")]')))
    next_button = driver.find_element(By.XPATH, '//button[contains(text(), "Next")]')
    next_button.click()

    # 팝업에서 Confirm 버튼 클릭
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Confirm")]')))
    confirm_button = driver.find_element(By.XPATH, '//button[contains(text(), "Confirm")]')
    confirm_button.click()

    # 홈 화면으로 이동한 후, itemlist 페이지로 이동
    driver.get('https://www.hangawee.com.au/?page=mitems')

    # 테이블이 로드될 때까지 대기
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//table/tbody/tr')))

    # JavaScript로 항목 수를 700개로 변경
    dropdown_script = '''
    document.querySelector('button.dropdown-toggle').click();
    document.querySelectorAll('.dropdown-menu a')[5].click();
    '''
    driver.execute_script(dropdown_script)
    time.sleep(5)  # 변경 후 페이지가 로드될 시간을 기다림

# 다음 페이지 화살표 클릭 및 스크롤 처리
def go_to_next_page():
    try:
        # 페이지 아래로 스크롤
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)  # 스크롤 후 잠시 대기
        
        # 다음 페이지 화살표 클릭
        next_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@class="datatable-pager-link datatable-pager-link-next"]')))
        next_button.click()
        
        # 페이지 로딩 대기
        time.sleep(3)  # 페이지가 완전히 로드될 때까지 기다림
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//table/tbody/tr')))
        time.sleep(3)

        # 페이지 위로 스크롤
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(2)  # 스크롤 후 잠시 대기
    except TimeoutException as e:
        print(f"Error navigating to next page: TimeoutException: {e}")
        driver.quit()
    except WebDriverException as e:
        print(f"Error navigating to next page: WebDriverException: {e}")
        driver.quit()

# 데이터 스크래핑 및 CSV 출력
def scrape_data_and_save_to_csv():
    # CSV 파일을 열고 쓰기 모드로 설정
    with open('scraped_data.csv', mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file)
        
        # CSV 파일에 헤더 작성
        writer.writerow(['BarCode', 'Description', 'CAROUSEL', 'NORTHBRIDGE', 'INNALOO', 'BOORAGOON', 'CFC'])

        # 1페이지부터 4페이지까지 순회하며 데이터 스크래핑
        for page in range(1, 5):
            print(f"Scraping page {page}...")
            if page > 1:
                # 다음 페이지로 이동
                go_to_next_page()

            # 페이지의 모든 행 선택
            rows = driver.find_elements(By.XPATH, '//table/tbody/tr')
            
            # 각 행에서 데이터를 추출
            for row in rows:
                try:
                    # 각 열의 데이터를 정확히 추출
                    barcode = row.find_element(By.XPATH, './td[1]').text.strip()  # BarCode 열
                    description = row.find_element(By.XPATH, './td[2]').text.strip()  # Description 열
                    carousel = row.find_element(By.XPATH, './td[3]').text.strip()  # CAROUSEL 열
                    northbridge = row.find_element(By.XPATH, './td[4]').text.strip()  # NORTHBRIDGE 열
                    innaloo = row.find_element(By.XPATH, './td[5]').text.strip()  # INNALOO 열
                    booragoon = row.find_element(By.XPATH, './td[6]').text.strip()  # BOORAGOON 열
                    cfc = row.find_element(By.XPATH, './td[7]').text.strip()  # CFC 열
                    
                    # 추출한 데이터를 한 줄씩 CSV 파일에 작성
                    writer.writerow([barcode, description, carousel, northbridge, innaloo, booragoon, cfc])
                
                except Exception as e:
                    print(f"Error scraping data from row: {e}")

# 스크립트 실행
login_and_navigate()

# 데이터 스크래핑 및 CSV 저장 실행
scrape_data_and_save_to_csv()

# 브라우저 종료
driver.quit()
