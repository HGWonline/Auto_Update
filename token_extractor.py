import json
import requests
from playwright.sync_api import sync_playwright

### 🛎️ 슬랙 알림 함수
def send_slack_message(text):
    webhook_url = "https://hooks.slack.com/services/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"  # 👈 슬랙 Webhook URL 입력
    payload = {"text": text}
    try:
        response = requests.post(webhook_url, json=payload)
        if response.status_code != 200:
            print(f"❗ 슬랙 알림 실패: {response.status_code}")
        else:
            print("📨 슬랙 알림 전송 성공")
    except Exception as e:
        print(f"❗ 슬랙 알림 오류: {e}")

def extract_tokens():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        try:
            print("🔐 로그인 페이지 접속 중...")
            page.goto("https://www.hangawee.com.au/login")

            print("🔢 로그인 코드 입력...")
            page.fill('input#username', '245554073198')
            page.click('button#kt_login_signin_submit')

            print("🏪 Outlet 선택 중...")
            page.wait_for_selector("label:has-text('Booragoon')", timeout=10000)
            page.click("label:has-text('Booragoon')")
            page.click("button#next-step")
            page.wait_for_selector("button#goIndex", timeout=10000)
            page.click("button#goIndex")

            print("⏳ 요청 감지 대기 중...")

            token_found = {"csrf": None, "jsession": None}

            def handle_request(request):
                headers = request.headers
                if "x-csrf-token" in headers and not token_found["csrf"]:
                    token_found["csrf"] = headers["x-csrf-token"]
                    cookies = context.cookies()
                    for cookie in cookies:
                        if cookie["name"] == "JSESSIONID":
                            token_found["jsession"] = cookie["value"]

            page.on("request", handle_request)
            page.wait_for_timeout(5000)

            context.close()
            browser.close()

            if token_found["csrf"] and token_found["jsession"]:
                with open("token.json", "w") as f:
                    json.dump({
                        "JSESSIONID": token_found["jsession"],
                        "X_CSRF_TOKEN": token_found["csrf"]
                    }, f)
                print("✅ token.json 저장 완료")
                send_slack_message("✅ token.json 저장 완료되었습니다.")
            else:
                print("❌ 토큰 추출 실패: 요청이 감지되지 않았거나 토큰 없음")
                send_slack_message("❌ 토큰 추출 실패: 요청이 감지되지 않았거나 토큰 없음")

        except Exception as e:
            print(f"❌ 예외 발생: {e}")
            send_slack_message(f"❌ 예외 발생 중단됨:\n```\n{e}\n```")

if __name__ == "__main__":
    extract_tokens()
