import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
import openpyxl
import re
from datetime import datetime
import traceback

chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--window-size=1920x1080')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

def wait_for_page_load(timeout=40):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )

def login_and_navigate_to_purchase(driver):
    # 로그인 페이지 접속
    driver.maximize_window()
    driver.get('https://www.hangawee.com.au/login')

    # 로그인
    username_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'username'))
    )
    username_input.send_keys('245554073198')

    sign_in_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'kt_login_signin_submit'))
    )
    sign_in_button.click()

    # Outlet 선택 (예: Carousel)
    booragoon_label = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//label[contains(., 'Carousel')]"))
    )
    booragoon_label.click()

    # Next 버튼 클릭
    next_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'next-step'))
    )
    next_button.click()

    # 팝업 Confirm 버튼 클릭
    confirm_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'goIndex'))
    )
    confirm_button.click()

    # Purchase 페이지로 이동
    driver.get('https://www.hangawee.com.au/?page=purchase')
    wait_for_page_load()

def set_date_range_and_search():
    start_date_obj = datetime.now() - timedelta(days=10)
    start_date = start_date_obj.strftime("%d/%m/%Y")
    end_date = datetime.now().strftime("%d/%m/%Y")

    start_date_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'startDate'))
    )
    end_date_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'endDate'))
    )

    start_date_input.clear()
    start_date_input.send_keys(start_date)
    end_date_input.clear()
    end_date_input.send_keys(end_date)

    # Purchase history 버튼 클릭
    purchase_history_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'kt_search'))
    )
    purchase_history_button.click()

    wait_for_page_load()

def change_view_to_50():
    try:
        # 드롭다운 토글 버튼 클릭
        dropdown_toggle = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".datatable-pager-size .dropdown-toggle"))
        )
        dropdown_toggle.click()  # 드롭다운 열기

        # '50' 옵션 클릭
        view_50_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li/a/span[text()='50']"))
        )
        view_50_element.click()

        wait_for_page_load()
        print("View changed to 50 successfully.")
    except Exception as e:
        print("Error changing view to 50:", e)
        traceback.print_exc()

def scrape_purchase_data():
    # 엑셀 파일 생성
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Purchase Data"
    # Comment 열 추가
    sheet.append(['Description', 'Purchase Date', 'Price', 'User Name', 'Comment'])

    # scraped_data 구조 변경: {description: (purchase_dt, price, user_name, comment)}
    scraped_data = {}
    page = 1
    while True:
        print(f"Scraping page {page}...")
        # 테이블 로딩 대기
        try:
            rows = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'table.datatable-table tbody.datatable-body tr'))
            )
        except:
            # 테이블 행이 없으면 종료
            print("No rows found, stopping.")
            break

        if not rows:
            print("No rows on this page, stopping.")
            break

        for row in rows:
            try:
                description = row.find_element(By.CSS_SELECTOR, 'td[data-field="description"]').text.strip()
                purchase_date = row.find_element(By.CSS_SELECTOR, 'td[data-field="purchaseDate"]').text.strip()
                price = row.find_element(By.CSS_SELECTOR, 'td[data-field="price"] input').get_attribute('value').strip()
                user_name = row.find_element(By.CSS_SELECTOR, 'td[data-field="realName"]').text.strip()

                # Comment 추출
                # input 안의 value로 comment가 존재
                comment_element = row.find_element(By.CSS_SELECTOR, 'td[data-field="comment"] input')
                comment = comment_element.get_attribute('value').strip()

                # 날짜 파싱 (일/월/년)
                try:
                    purchase_dt = datetime.strptime(purchase_date, "%d/%m/%Y")
                except Exception as e:
                    print(f"Date parsing error for {purchase_date}: {e}")
                    purchase_dt = None

                # 중복 제거 로직: description 동일하면 더 최신 날짜의 데이터만 유지
                if description in scraped_data:
                    existing_dt = scraped_data[description][0]
                    if purchase_dt and existing_dt and purchase_dt > existing_dt:
                        scraped_data[description] = (purchase_dt, price, user_name, comment)
                else:
                    scraped_data[description] = (purchase_dt, price, user_name, comment)

            except Exception as e:
                print("Error scraping a row:", e)
                traceback.print_exc()
                continue

        # 다음 페이지 이동 시도
        try:
            next_btn = driver.find_element(By.CSS_SELECTOR, 'a.datatable-pager-link[title="Next"]')
            next_btn_classes = next_btn.get_attribute('class')
            if 'datatable-pager-link-disabled' in next_btn_classes or next_btn.get_attribute('disabled') == 'disabled':
                print("Reached the last page.")
                break

            # 다음 페이지로 이동
            next_btn.click()
            time.sleep(3)
            wait_for_page_load()
            page += 1
        except:
            print("No next button found, stopping.")
            break

    # scraped_data를 엑셀에 저장
    for desc, (p_dt, p_price, p_user, p_comment) in scraped_data.items():
        p_date_str = p_dt.strftime("%d/%m/%Y") if p_dt else ''
        sheet.append([desc, p_date_str, p_price, p_user, p_comment])

    workbook.save(f'Purchase_{datetime.now().strftime("%y%m%d")}.xlsx')
    print("Data successfully saved.")

# 실행
login_and_navigate_to_purchase(driver)
set_date_range_and_search()
change_view_to_50()
scrape_purchase_data()
driver.quit()
