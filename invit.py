import pandas as pd
import os
import qrcode
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display

# --------- إعداد المجلدات ---------
os.makedirs("pages", exist_ok=True)
os.makedirs("qrcards", exist_ok=True)


# --------- إعدادات الإطار الفخم ---------
# --------- إعدادات التصميم ---------
GOLD = (212, 175, 55)  # لون ذهبي
grey = (100,100,100)
BORDER_COLOR = grey
BORDER_WIDTH = 4

RADIUS = 18        # استدارة الزوايا
SHADOW_OFFSET = 4 # الظل




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
df = pd.read_excel("invites.xlsx")  # Name | Table | Guests

# --------- تحميل الخلفية ---------
background = Image.open("log_2.jpg").convert("RGB")
background = background.resize((450, 500))

# --------- تحميل الخط ---------
font_title = ImageFont.truetype("Amiri-Bold.ttf", 25)
font_name = ImageFont.truetype("Amiri-BoldItalic.ttf", 25)

# ======================================================
#                   تنفيذ السكربت
# ======================================================

for index, row in df.iterrows():
    serial = index + 1
    name = str(row['Name']).strip()
    table = row['Table']
    guests = row['Guests']
    card = background.copy()
    draw = ImageDraw.Draw(card)

    # --------- رابط الصفحة ---------
    page_link = f"https://mosabkassar.github.io/{serial}.html"

# رابط الفيديو (يمكن أن يكون ملف mp4 محلي أو رابط خارجي)
    video_file = "video.mp4"  # أو رابط يوتيوب أو رابط مباشر للفيديو

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


    # --------- QR البيانات ---------
    qr_data = f"""الاسم: {name}
رقم الطاولة: {guests}
عدد الأشخاص: {table}"""
    qr_info = qrcode.make(qr_data).resize((65, 65))

    # --------- QR الرابط ---------
    qr_link = qrcode.make(page_link).resize((65, 65))

    # --------- إنشاء البطاقة ---------
    card = background.copy()
    draw = ImageDraw.Draw(card)

    # --------- النص الرسمي (منتصف البطاقة) ---------
    y_center = card.height // 2 + 210

    line1 = fix_arabic(f"{name}") 
    draw.rectangle(
        [2, 2, card.width - 2, card.height - 2],
        outline=BORDER_COLOR,
        width=BORDER_WIDTH
    )

   

    # --------- النص ---------
    draw.text(
        (card.width // 2, y_center),
        line1,
        font=font_title,
        fill="black",
        anchor="mm"
    )

    # --------- أماكن QR ---------
    margin = 6
    y_qr = card.height - qr_info.height - margin

    card.paste(qr_info, (margin, y_qr))
    card.paste(qr_link, (card.width - qr_link.width - margin, y_qr))

    # --------- حفظ البطاقة ---------
    safe_name = safe_filename(name)
    card.save(f"qrcards/{serial}_{safe_name}.png")

    print(f"{serial} → {name}")

    
    margin = 6
    y_qr = card.height - qr_info.height - margin

    # يسار
    card.paste(qr_info, (margin, y_qr))

    # يمين
    card.paste(qr_link, (card.width - qr_link.width - margin, y_qr))

    # --------- حفظ البطاقة ---------
    safe_name = safe_filename(name)
    card.save(f"qrcards/{serial}_{safe_name}.png")

    print(f"{serial} → {name}")

print("\n🎉 تم إنشاء دعوات رسمية مع QR في الزوايا بنجاح!")
