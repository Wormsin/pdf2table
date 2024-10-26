import fitz  # PyMuPDF
import os 
from PIL import Image
import pytesseract
import pandas as pd
from img2table.document import Image
from img2table.ocr import TesseractOCR


def extract_images_from_pdf(pdf_path, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    pdf_document = fitz.open(pdf_path)
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = pdf_document.extract_image(xref)
            image_bytes = base_image["image"]
            image_filename = f"{output_folder}/page_{page_num + 1}_img_{img_index + 1}.png"
            with open(image_filename, "wb") as image_file:
                image_file.write(image_bytes)

def ocr_image_to_text(image_path):
    image = Image.open(image_path)
    text = pytesseract.image_to_string(image, lang='rus')
    return text

def process_images(image_folder):
    for image_file in os.listdir(image_folder):
        image_path = os.path.join(image_folder, image_file)
        text = ocr_image_to_text(image_path)
        with open(f"{image_path}.txt", "w") as text_file:
            text_file.write(text)



#extract_images_from_pdf("Кинеф Потребность 2024.pdf", "output_images")
#process_images("output_images")

def text_to_dataframe(text_file):
    # Пример обработки текста
    with open(text_file, "r") as file:
        lines = file.readlines()
    data = [line.strip().split() for line in lines]
    df = pd.DataFrame(data[1:], columns=data[0])
    return df

def save_dataframe_to_excel(df, excel_file):
    df.to_excel(excel_file, index=False)

df = text_to_dataframe("output_images/page_3_img_1.png.txt")
save_dataframe_to_excel(df, "output.xlsx")



