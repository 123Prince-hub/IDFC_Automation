import pytesseract
from PIL import Image
from time import sleep

# img = Image.open('img.jpg')
# imgGray = img.convert('L')
# imgGray.save('img.jpg')
sleep(1)

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
a = pytesseract.image_to_string(r'img.png').strip()
a = a.replace("(", "j")
print(a)
