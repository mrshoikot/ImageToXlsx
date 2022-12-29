# ImageToXlsx

Extract table content from given image(jpg, png) and put them into an xlsx file.
The process uses OpenCV for processing the image and pytesseract as an OCR tool for reading the text content from the input image.
From the extracted text, the system recognizes patterns and separates rows and columns and fianally put them into workbook.

## install Dependenceis
`pip install -r requirements.txt`

## Run the tool
`python tool.py`

It'll open up a GUI by which the user will be able to choose the input image. There is also a exe file available in the `dist` directory named as `tool.exe` which can be used in windows systems and it does require any prior environment setup.

The final result will be stored in the `result` directory.

## Dependecies
- xlsxwriter
- pytesseract
- Pillow
- opencv-python
- tqdm
- numpy

## Useful links

- [Tesseract OCR](https://tesseract-ocr.github.io/)
- [PyTesseract](https://pypi.org/project/pytesseract/)