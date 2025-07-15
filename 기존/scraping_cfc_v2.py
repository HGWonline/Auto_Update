import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# ChromeDriver 경로 설정
chrome_driver_path = r'C:\Users\김남빈\chromedriver\chromedriver-win64\chromedriver.exe'
chrome_service = Service(chrome_driver_path)

# WebDriver 시작
driver = webdriver.Chrome(service=chrome_service)

# 로그인 및 네비게이션
def login_and_navigate():
    driver.maximize_window()
    driver.get('https://www.hangawee.com.au/login')

    # 코드 입력 및 로그인 처리
    driver.find_element(By.XPATH, '//input[@type="text"]').send_keys('245554073198')
    driver.find_element(By.XPATH, '//button[contains(text(), "Sign In")]').click()

    # 추가적인 버튼 처리
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Next")]'))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Confirm")]'))).click()

    # itemlist 페이지로 이동
    driver.get('https://www.hangawee.com.au/?page=mitems')

    # JavaScript로 항목 수를 700개로 변경
    try:
        dropdown_toggle = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.dropdown-toggle')))
        dropdown_toggle.click()
        
        # 드롭다운 목록의 항목을 클릭 (5번째 항목을 선택)
        dropdown_item = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.dropdown-menu a:nth-child(6)')))
        dropdown_item.click()
        
    except Exception as e:
        print(f"Error in dropdown menu selection: {e}")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//table/tbody/tr')))

# 페이지 전환
def go_to_next_page():
    try:
        # 페이지 아래로 스크롤 및 다음 페이지 버튼 클릭
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@class="datatable-pager-link datatable-pager-link-next"]'))
        ).click()

        # 다음 페이지 로드 대기
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//table/tbody/tr')))
    except Exception as e:
        print(f"Error navigating to next page: {e}")
        driver.quit()

# 데이터 스크래핑 및 CSV 출력
def scrape_data_and_save_to_csv():
    with open('scraped_data.csv', mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file)
        writer.writerow(['BarCode', 'Description', 'CAROUSEL', 'NORTHBRIDGE', 'INNALOO', 'BOORAGOON', 'CFC'])

        for page in range(1, 5):
            if page > 1:
                go_to_next_page()

            rows = driver.find_elements(By.XPATH, '//table/tbody/tr')

            for row in rows:
                try:
                    data = [row.find_element(By.XPATH, f'./td[{i}]').text.strip() for i in range(1, 8)]
                    writer.writerow(data)
                except Exception as e:
                    print(f"Error scraping data from row: {e}")

# 스크립트 실행
login_and_navigate()
scrape_data_and_save_to_csv()
driver.quit()
