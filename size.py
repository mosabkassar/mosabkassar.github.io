from PIL import Image

img = Image.open("logo_3.jpg")
print(img.size)   # (العرض, الارتفاع)
print(img.info.get("dpi"))
