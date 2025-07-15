from flask import Flask, redirect, request, session, url_for
import os
import requests
from urllib.parse import urlencode
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = 'super_secret_key'  # 세션 보안용

# 환경변수
API_KEY = os.getenv('SHOPIFY_API_KEY')
API_SECRET = os.getenv('SHOPIFY_API_SECRET')
SHOP = os.getenv('SHOP')  # 예: your-store.myshopify.com
SCOPES = "read_products,write_products,read_custom_collections,write_custom_collections,read_inventory"
REDIRECT_URI = os.getenv('REDIRECT_URI')  # 예: https://abcd.ngrok-free.app/callback

# 🔹 루트 경로: 기본 페이지
@app.route('/')
def index():
    return '''
        <h2>✅ AutoSortify 작동 중</h2>
        <a href="/auth">👉 Shopify 앱 인증 시작하기</a>
    '''

# 🔹 Shopify Admin에서 앱 클릭 시 진입하는 경로
@app.route('/apps/AutoSortify')
def app_launch():
    return redirect('/auth')

# 🔹 인증 시작
@app.route('/auth')
def auth():
    install_url = f"https://{SHOP}/admin/oauth/authorize?" + urlencode({
        "client_id": API_KEY,
        "scope": SCOPES,
        "redirect_uri": REDIRECT_URI,
    })
    return redirect(install_url)

# 🔹 인증 후 콜백: access_token 발급
@app.route('/callback')
def callback():
    code = request.args.get("code")
    payload = {
        "client_id": API_KEY,
        "client_secret": API_SECRET,
        "code": code
    }
    response = requests.post(f"https://{SHOP}/admin/oauth/access_token", json=payload)
    access_token = response.json().get('access_token')
    session['access_token'] = access_token

    if access_token:
        return redirect('/products')
    else:
        return "❌ 인증 실패"

# 🔹 제품 목록 가져오기 (테스트용 API 호출)
@app.route('/products')
def get_products():
    access_token = session.get('access_token')
    if not access_token:
        return redirect('/auth')

    headers = {
        "X-Shopify-Access-Token": access_token
    }
    response = requests.get(f"https://{SHOP}/admin/api/2023-10/products.json", headers=headers)
    products = response.json().get('products', [])
    return f"📦 총 {len(products)}개의 제품이 있습니다."

if __name__ == '__main__':
    app.run(debug=True, port=5000)
