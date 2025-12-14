import pandas as pd
import os
import qrcode
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display

# --------- إعداد المجلدات ---------
os.makedirs("pages", exist_ok=True)
os.makedirs("qrcards", exist_ok=True)

# --------- إعدادات المستطيل السفلي ---------
BOTTOM_BOX_HEIGHT = 150
BOTTOM_BOX_COLOR = (255, 255, 255)
TEXT_COLOR = (0, 0, 0)

# --------- إعدادات الإطار ---------
BORDER_COLOR = (100, 100, 100)
BORDER_WIDTH = 8

# --------- الخطوط ---------
font_title = ImageFont.truetype("Amiri-BoldItalic copy.ttf", 45)
font_small = ImageFont.truetype("ScheherazadeNew-Bold.ttf", 22)

# --------- دالة إصلاح العربية ---------
def fix_arabic(text):
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)

# --------- تنظيف اسم الملف ---------
def safe_filename(text):
    invalid = '<>:"/\\|?*'
    for c in invalid:
        text = text.replace(c, "_")
    return text

# --------- قراءة ملف الإكسل ---------
df = pd.read_excel("invites_2.xlsx")

# --------- تحميل الخلفية (بدون تغيير الحجم) ---------
background = Image.open("logo_3.jpg").convert("RGB")
bg_width, bg_height = background.size

# ======================================================
#                   تنفيذ السكربت
# ======================================================

for index, row in df.iterrows():
    serial = index + 1
    name = str(row['Name']).strip()
    guests = row['Table']
    table = row['Guests']

    # --------- رابط الصفحة ---------
    page_link = f"https://mosabkassar.github.io/pages/{serial}.html"
    video_file = "video.mp4"  # ضع الفيديو داخل مجلد pages/assets/


    # --------- إنشاء صفحة HTML ---------
   
    html_content = f"""
    <!DOCTYPE html>
    <html lang="ar" dir="rtl">
    <head>
    <meta charset="UTF-8">
    <title>دعوة {name}</title>
    <style>
      body {{ font-family: Arial, sans-serif; text-align: center; margin: 50px; background-color: #f9f9f9; }}
      h1 {{ color: #333; }}
      video {{ width: 80%; max-width: 600px; border: 3px solid #ccc; border-radius: 10px; }}
    </style>
    </head>
    <body>
    <h1>مرحباً {name}</h1>
    <p>رقم الطاولة: {table} | عدد الأشخاص: {guests}</p>
    <video controls>
      <source src="{video_file}" type="video/mp4">
      متصفحك لا يدعم عرض الفيديو.
    </video>
    </body>
    </html>
    """
    with open(f"pages/{serial}.html", "w", encoding="utf-8") as f:
        f.write(html_content)
    # --------- إنشاء QR (بدون تكبير) ---------
    qr_info = qrcode.make(
        f"الاسم: {name}\nرقم الطاولة: {table}\nعدد الأشخاص: {guests}\n المؤشر : {index} "
    ).convert("RGB")

    qr_link = qrcode.make(page_link).convert("RGB")

    qr_size = 140
    qr_info = qr_info.resize((qr_size, qr_size), Image.NEAREST)
    qr_link = qr_link.resize((qr_size, qr_size), Image.NEAREST)

    # --------- إنشاء كارد جديد بالحجم الأصلي ---------
    card_width = bg_width
    card_height = bg_height + BOTTOM_BOX_HEIGHT

    card = Image.new("RGB", (card_width, card_height), (255, 255, 255))
    card.paste(background, (0, 0))
    draw = ImageDraw.Draw(card)

    # --------- المستطيل السفلي ---------
    box_top = bg_height
    draw.rectangle(
        [0, box_top, card_width, card_height],
        fill=BOTTOM_BOX_COLOR
    )

    # --------- الاسم (داخل الصورة الأصلية) ---------
    draw.text(
        (card_width // 2, box_top + BOTTOM_BOX_HEIGHT // 2),
        fix_arabic(name),
        font=font_title,
        fill="black",
        anchor="mm"
    )



    # --------- QR داخل المستطيل ---------
    margin = 25
    qr_y = box_top + (BOTTOM_BOX_HEIGHT - qr_size) // 2

    card.paste(qr_info, (margin, qr_y))
    card.paste(qr_link, (card_width - qr_size - margin, qr_y))

    # --------- الإطار الخارجي ---------
    draw.rectangle(
        [2, 2, card_width - 2, card_height - 2],
        outline=BORDER_COLOR,
        width=BORDER_WIDTH
    )

    # --------- حفظ البطاقة بدقة عالية ---------
    safe_name = safe_filename(name)
    card.save(
        f"qrcards/{serial}_{safe_name}.png",
        dpi=(300, 300)
    )

    print(f"{serial} → {name}")

print("\n🎉 تم إنشاء البطاقات مع الحفاظ الكامل على الدقة!")
