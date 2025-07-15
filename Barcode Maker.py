import barcode
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont

def generate_label(product_label: str, barcode_value: str, output_file="label.png"):
    # 1) 바코드 생성
    Code128 = barcode.get_barcode_class('code128')
    code128 = Code128(barcode_value, writer=ImageWriter())
    code128.default_writer_options["module_width"]  = 1.0
    code128.default_writer_options["module_height"] = 8.0
    code128.default_writer_options["quiet_zone"]    = 3.0
    code128.default_writer_options["font_size"]     = 10
    code128.default_writer_options["text_distance"] = 2.0
    code128.default_writer_options["write_text"]    = False

    temp_barcode_path = code128.save("temp_barcode")

    # 2) 4cm×2cm @300dpi → (472×236)
    label_width, label_height = 472, 236
    label_img = Image.new("RGB", (label_width, label_height), "white")
    draw = ImageDraw.Draw(label_img)

    # 3) 폰트 설정 (상품명용 24, 번호용 18)
    font_title = ImageFont.truetype("arial.ttf", 24)
    font_num   = ImageFont.truetype("arial.ttf", 18)

    # 4) (위) 상품명
    title_bbox = draw.textbbox((0, 0), product_label, font=font_title)
    title_w = title_bbox[2] - title_bbox[0]
    title_h = title_bbox[3] - title_bbox[1]

    title_x = (label_width - title_w) // 2
    title_y = 30  # 위쪽 여백 30px
    draw.text((title_x, title_y), product_label, font=font_title, fill="black")

    # 5) 바코드
    barcode_img = Image.open(temp_barcode_path)
    bw, bh = barcode_img.size
    barcode_x = (label_width - bw) // 2
    barcode_y = title_y + title_h + 10  # 상품명 아래 10px
    label_img.paste(barcode_img, (barcode_x, barcode_y))

    # 6) (아래) 바코드 번호
    num_text = barcode_value
    num_bbox = draw.textbbox((0,0), num_text, font=font_num)
    num_w = num_bbox[2] - num_bbox[0]
    num_h = num_bbox[3] - num_bbox[1]

    num_x = (label_width - num_w) // 2
    num_y = barcode_y + bh + 10   # 바코드 아래 10px
    # 혹시 화면 아래를 넘어서진 않는지 체크
    if num_y + num_h > label_height:
        num_y = label_height - num_h - 5

    draw.text((num_x, num_y), num_text, font=font_num, fill="black")

    # 7) 저장
    label_img.save(output_file, "PNG")

if __name__ == "__main__":
    generate_label(
        "Keyring $9.99",
        "112005400",
        "Barcode_final.png"
    )
    print("라벨 이미지 생성 완료!")
