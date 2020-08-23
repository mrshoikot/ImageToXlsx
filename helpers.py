import os
import pytesseract
from PIL import Image
import cv2
import xlsxwriter
import tkinter
from tkinter import filedialog



lineHeight = 4
rowHeight = 28
rowCount = 100

class Row:

    def __init__(self, photo):
        self.photo = photo
        self.data = []
        # cv2.imshow("rowImage", photo)
        # cv2.waitKey(0)

    def getCols(self):
        coords = [223,411,679,816,1123,1400,1584,1751,1936,2061,2357,2451,2577,2718]
        pos = 0
        
        for coord in coords:

            colImage = self.photo[0:rowHeight, pos:coord-lineHeight]

            # cv2.imshow("rowImage", colImage)
            # cv2.waitKey(0)

            text = pytesseract.image_to_string(colImage, config=r'--oem 3 --psm 6')
            self.data.append(text.strip())

            pos = coord


def initWorksheet():
    workbook = xlsxwriter.Workbook('./result/data.xlsx')
    worksheet = workbook.add_worksheet('Data')

    worksheet.freeze_panes(1, 0)
    worksheet.set_column(0, 0, 15)
    worksheet.set_column(1, 1, 20)
    worksheet.set_column(2, 2, 25)
    worksheet.set_column(3, 3, 10)
    worksheet.set_column(4, 4, 35)
    worksheet.set_column(5, 5, 35)
    worksheet.set_column(6, 6, 30)
    worksheet.set_column(7, 7, 30)
    worksheet.set_column(8, 8, 30)
    worksheet.set_column(9, 9, 20)
    worksheet.set_column(10, 10, 50)
    worksheet.set_column(11, 11, 10)
    worksheet.set_column(12, 12, 10)
    worksheet.set_column(13, 13, 15)

    bold = workbook.add_format({'bold': True})

    worksheet.write(0, 0, "RecordNo", bold)
    worksheet.write(0, 1, "Policy Data", bold)
    worksheet.write(0, 2, "PolicyNo", bold)
    worksheet.write(0, 3, "Medical Card", bold)
    worksheet.write(0, 4, "First Name", bold)
    worksheet.write(0, 5, "Last Name", bold)
    worksheet.write(0, 6, "City Name", bold)
    worksheet.write(0, 7, "State Name", bold)
    worksheet.write(0, 8, "Phone", bold)
    worksheet.write(0, 9, "Martial", bold)
    worksheet.write(0, 10, "gpcode", bold)
    worksheet.write(0, 11, "Hosp Days", bold)
    worksheet.write(0, 12, "Paid", bold)
    worksheet.write(0, 13, "Net amt", bold)

    return {'worksheet': worksheet, 'workbook': workbook}


def getFiles():
    root = tkinter.Tk()
    filez = tkinter.filedialog.askopenfilenames(parent=root,title='Choose a file')
    return root.tk.splitlist(filez)
