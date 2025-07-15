import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager  # 자동 드라이버 관리
import openpyxl
import re
from datetime import datetime
import traceback  # 추가된 부분: traceback 모듈

# Chrome 옵션 설정
chrome_options = Options()
# chrome_options.add_argument('--headless')  # 브라우저 창을 표시하지 않음 (디버깅 시 주석 처리)
chrome_options.add_argument('--disable-gpu')  # GPU 비활성화
chrome_options.add_argument('--window-size=1920x1080')  # 창 크기 설정

# WebDriver 설정 및 자동 업데이트
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

def wait_for_page_load(timeout=40):
    """ 페이지 로드가 완료될 때까지 대기하는 함수 """
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )
    # 추가적인 요소 확인 (테이블이 로드되었는지 확인)
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'table tbody tr'))
    )

# 활성화 버튼 빠르게 클릭하는 함수
def activate_buttons():
    try:
        print("Activating buttons...")

        # 두 버튼을 클릭하는 스크립트 한번에 처리
        script = """
        document.querySelector("input[name='radios13'][value='1']").click();
        document.querySelector("input[name='radios14'][value='1']").click();
        """
        driver.execute_script(script)

        wait_for_page_load()
        print("Both 'WholesalerInactive' and 'RetailInactive' set to ACTIVE.")
        
    except Exception as e:
        print(f"Error activating buttons: {e}")
        traceback.print_exc()

# 페이지 아래로 스크롤하여 '보기'를 'all'로 변경하는 함수
def change_view_to_all():
    try:
        print("Changing view to 'all'...")

        # select 요소를 직접 찾아 'all' (값이 '999')로 변경
        select_element = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "select.selectpicker.datatable-pager-size"))
        )
        
        # JavaScript로 값을 직접 변경하여 'all'을 선택
        driver.execute_script("arguments[0].value = '999'; arguments[0].dispatchEvent(new Event('change'))", select_element)

        wait_for_page_load()  # 페이지 로딩 대기
        print("View changed to 'all'.")
        
    except Exception as e:
        print(f"Error changing view to 'all': {e}")
        traceback.print_exc()

# 로그인 및 페이지 이동 함수
def login_and_navigate():
    try:
        print("Navigating to login page...")
        driver.get('https://www.hangawee.com.au/login')
        
        # 로그인 과정
        print("Entering login code...")
        code_input = WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.ID, 'username')))
        code_input.send_keys('245554073198')

        sign_in_button = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.ID, 'kt_login_signin_submit')))
        print("Clicking sign in button...")
        sign_in_button.click()

        next_button = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.ID, 'next-step')))
        print("Clicking Next button...")
        next_button.click()

        confirm_button = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.ID, 'goIndex')))
        print("Clicking Confirm button on popup...")
        confirm_button.click()

        # Item Management 페이지로 이동
        print("Navigating to item management page...")
        driver.get('https://www.hangawee.com.au/?page=itemmanagement')
        wait_for_page_load()

        # WholesalerInactive와 Retail Inactive 활성 상태로 변경
        activate_buttons()

        # 페이지 아래로 스크롤하여 '보기'를 'all'로 변경
        change_view_to_all()

    except Exception as e:
        print(f"Error during login and navigation: {e}")
        traceback.print_exc()
        driver.quit()
        return  # 프로그램 종료 대신 함수 종료로 변경

# 괄호 안의 내용을 제거하는 함수
def remove_parentheses_content(product_name):
    """ 괄호 안의 내용을 모두 제거하는 함수 """
    return re.sub(r'\(.*?\)', '', product_name).strip()

# 정보를 스크래핑하여 엑셀로 저장하는 함수
def scrape_data():
    # Excel 파일 생성
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Scraped Data"

    # 헤더 작성 (기존 헤더 유지)
    sheet.append(['상품 이름', 'Barcode', 'wholesalerPrice', 'retailPrice'])

    try:
        page = 1
        while True:  # 더 이상 페이지가 없을 때까지 반복
            print(f"Scraping page {page}...")
            rows = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'table tbody tr')))
            
            if not rows:
                print("No rows found on the page. Exiting.")
                break
            
            for row in rows:
                try:
                    # 숨겨진 정보 스크래핑
                    product_name = row.find_element(By.CSS_SELECTOR, 'td[data-field="wholesalerDescription"]').text
                    xero_barcode = row.find_element(By.CSS_SELECTOR, 'td[data-field="wholesalerBarCode"]').text
                    wholesaler_price = row.find_element(By.CSS_SELECTOR, 'td[data-field="wholesalerPrice"]').get_attribute('aria-label')
                    pak_price = row.find_element(By.CSS_SELECTOR, 'td[data-field="pakPrice"]').get_attribute('aria-label')
                    ea_price = row.find_element(By.CSS_SELECTOR, 'td[data-field="eaPrice"]').get_attribute('aria-label')

                    # 괄호 안의 내용을 모두 제거한 상품 이름으로 변경
                    product_name_cleaned = remove_parentheses_content(product_name)

                    # 매장별로 가격을 분리 및 다수의 가격 중 최빈값 선택
                    pak_prices = re.findall(r'\w{2}:(\d+\.\d+)', pak_price)  # CA, NB, IN, BR 순서로 가격 추출
                    ea_prices = re.findall(r'\w{2}:(\d+\.\d+)', ea_price)

                    # 최빈값 계산
                    pak_price_value = float(max(set(pak_prices), key=pak_prices.count)) if pak_prices else None
                    ea_price_value = float(max(set(ea_prices), key=ea_prices.count)) if ea_prices else None

                    # 바코드 가져오기
                    pak_barcode = row.find_element(By.CSS_SELECTOR, 'td[data-field="pakBarCode"]').get_attribute('aria-label')
                    ea_barcode = row.find_element(By.CSS_SELECTOR, 'td[data-field="eaBarCode"]').get_attribute('aria-label')

                    # 바코드 리스트 생성
                    barcodes = []
                    if pak_barcode and pak_barcode != 'N/A':
                        barcodes.append(str(int(pak_barcode)) if pak_barcode.isdigit() else pak_barcode)
                    if ea_barcode and ea_barcode != 'N/A':
                        barcodes.append(str(int(ea_barcode)) if ea_barcode.isdigit() else ea_barcode)

                    # 가격 리스트 생성
                    prices = []
                    if pak_price_value is not None:
                        prices.append(pak_price_value)
                    if ea_price_value is not None:
                        prices.append(ea_price_value)

                    # wholesaler_price를 float로 변환
                    if wholesaler_price != 'N/A':
                        wholesaler_price = float(wholesaler_price)
                    else:
                        wholesaler_price = None

                    # 바코드와 가격의 개수를 맞추기
                    min_length = min(len(barcodes), len(prices))

                    # 바코드와 가격을 개별 행으로 작성
                    for i in range(min_length):
                        barcode = barcodes[i]
                        retail_price = prices[i]
                        adjusted_wholesaler_price = wholesaler_price  # 초기값 설정

                        # wholesaler_price가 retail_price보다 큰 경우 로직 적용
                        if wholesaler_price is not None and retail_price is not None and wholesaler_price > retail_price:
                            try:
                                # 상품 이름에서 오른쪽부터 숫자 추출
                                numbers_in_name = re.findall(r'\d+', product_name_cleaned)
                                numbers_in_name = [int(num) for num in numbers_in_name]
                                numbers_in_name.reverse()  # 오른쪽부터 숫자를 사용하기 위해 역순으로 정렬

                                temp_wholesaler_price = wholesaler_price  # 임시 변수

                                for num in numbers_in_name:
                                    if num != 0:
                                        temp_wholesaler_price = temp_wholesaler_price / num
                                        if temp_wholesaler_price <= retail_price:
                                            adjusted_wholesaler_price = temp_wholesaler_price
                                            break  # 조건을 만족하면 종료

                                # 모든 숫자를 사용했음에도 wholesaler_price가 여전히 큰 경우
                                if temp_wholesaler_price > retail_price:
                                    adjusted_wholesaler_price = temp_wholesaler_price  # 마지막 계산값 사용

                            except Exception as e:
                                print(f"Error adjusting wholesalerPrice for {product_name_cleaned}: {e}")
                                traceback.print_exc()
                                adjusted_wholesaler_price = wholesaler_price  # 조정 실패 시 원래 값 사용

                        else:
                            adjusted_wholesaler_price = wholesaler_price

                        # 한 줄에 대한 데이터 작성
                        sheet.append([
                            product_name_cleaned,
                            barcode,
                            adjusted_wholesaler_price,
                            retail_price
                        ])

                except Exception as e:
                    print(f"Error scraping row: {e}")
                    traceback.print_exc()
                    continue

            # 다음 페이지로 이동
            try:
                # 다음 페이지 버튼이 활성화되어 있는지 확인
                next_page_button = driver.find_element(By.CSS_SELECTOR, 'a.datatable-pager-link[title="Next"]')
                if 'disabled' in next_page_button.get_attribute('class'):
                    print("Reached last page.")
                    break

                print("Clicking next page button...")
                next_page_button.click()
                time.sleep(3)  # 페이지 전환 후 안정적인 로딩을 위한 대기 추가
                wait_for_page_load()

                page += 1  # 페이지 번호 증가
            except Exception as e:
                print(f"Error navigating to next page: {e}")
                traceback.print_exc()
                break
        
    except Exception as e:
        print(f"Error during scraping: {e}")
        traceback.print_exc()
    
    # 데이터가 있는 경우에만 파일 저장
    if sheet.max_row > 1:
        workbook.save(f'ScrapedM_{datetime.now().strftime("%y%m%d")}.xlsx')
        print("Data successfully saved.")
    else:
        print("No data to save.")

# 실행
login_and_navigate()
scrape_data()
driver.quit()
