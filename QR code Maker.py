import qrcode

# QR 코드에 담을 데이터
data = "https://hangaweemarket.com/search?q=wheat+flour+cake&options%5Bprefix%5D=last"

# QR 코드 인스턴스 생성 (버전, 오류 수정, 박스 크기, 테두리 설정)
qr = qrcode.QRCode(
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_L,
    box_size=10,
    border=4,
)

# 데이터 추가 및 QR 코드 최적화
qr.add_data(data)
qr.make(fit=True)

# QR 코드 이미지 생성 (색상: 검정/흰색)
img = qr.make_image(fill_color="black", back_color="white")

# 생성된 QR 코드 이미지 파일로 저장
img.save("QR Code.png")

print("QR 코드 이미지가 example_qr.png 로 저장되었습니다.")
