import os
import xlsxwriter
import pytesseract
from PIL import Image
import cv2
from tqdm import tqdm, tqdm_gui
from helpers import *
from processor import *

if os.name == 'nt':
    pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'

lineHeight = 4
rowHeight = 28
rowCount = 100


excel = initWorksheet()
worksheet = excel['worksheet']
workbook = excel['workbook']


files = getFiles()
rowCount = 1

try:
    for path in files:

        y = 394
        x = 97
        w = 2896
        
        img = cv2.imread(path)
        img = cv2.resize(img, (w, 4096))

        deskew = deskew(img)
        img = get_grayscale(deskew)


        for i in tqdm(range(1,100+1)):
        
            rowImage = img[y:y+rowHeight, x:x+w]
            row = Row(rowImage)
            row.getCols()

            colCount = 0
            for col in row.data:
                worksheet.write(rowCount, colCount, col)
                colCount += 1

            y += rowHeight+6

            if i%20 == 0:
                y-=3

            rowCount += 1

    workbook.close()
    input("Press enter to exit...")

except:
    workbook.close()
    input("Press enter to exit...")
    exit()


