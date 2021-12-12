import os
import cv2
import jpype
import shutil
import weasyprint
from bs4 import BeautifulSoup

jpype.startJVM()
from asposecells.api import *

def generatePDF(XLSXPath, OutPath):
    workbook = Workbook(XLSXPath)
    workbook.save(f"sheet.html", SaveFormat.HTML)

    with open(f'./sheet_files/sheet001.htm') as f:
        htmlDoc = f.read()

    soup = BeautifulSoup(htmlDoc, 'html.parser')
    table = soup.find_all('table')[0]

    with open(f'./sheet_files/stylesheet.css') as f:
        styles = f.read()

    with open(f'out.html', 'w') as f:
        f.write(f'''
            <style>
                {styles}
                @page {{size: A4; margin:0;}}
                table {{margin:auto; margin-top: 5mm;}}
                table, tr, td {{border: 1px solid #000 !important;}}
            </style>
        ''')
        f.write(str(table.prettify()))

    weasyprint.HTML('out.html').write_pdf(OutPath)

def cleanPDF(OutPath):
    shutil.rmtree('./sheet_files')
    os.remove('./out.html')
    os.remove('./sheet.html')
    os.remove(OutPath)

def generatePNG(XLSXPath, OutPath):
    workbook = Workbook(XLSXPath)
    workbook.save(f"sheet.png", SaveFormat.PNG)
    img = cv2.imread("sheet.png")
    cropped = img[20:-100]
    cv2.imwrite(OutPath, cropped)

def cleanPNG(OutPath):
    os.remove('./sheet.png')
    os.remove(OutPath)