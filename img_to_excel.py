from img2table.document import PDF
from img2table.ocr import TesseractOCR
import os
from img2table.ocr import SuryaOCR


# Instantiation of the pdf
pdf = PDF(src="Кинеф Потребность 2024.pdf", pages=[0, 1, 2])

# Instantiation of the OCR, Tesseract, which requires prior installation
ocr = TesseractOCR(lang='eng+rus')


tables_per_page = []
for page_num in range(3):
    pdf = PDF(src="Кинеф Потребность 2024.pdf", pages=[page_num])
    extracted_tables = pdf.extract_tables(ocr=ocr, min_confidence=50)  # Экстракция таблиц
    tables_per_page.append(extracted_tables)
    pdf_table = pdf.extract_tables(ocr=ocr)
    pdf.to_xlsx(f'tables{page_num}_surya.xlsx', ocr=ocr)

# Table identification and extraction
pdf_tables = pdf.extract_tables(ocr=ocr)


# We can also create an excel file with the tables
pdf.to_xlsx('tables_surya.xlsx', ocr=ocr)
