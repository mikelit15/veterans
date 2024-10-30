import fitz  
import os
import time

input_directory = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted"
output_directory = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted - After Scaling"

target_width = 842  
target_height = 560

retry_delay = 3  
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

def resize_pdf_with_retry(input_pdf_path, output_pdf_path, target_width, target_height):
    while True:  
        try:
            pdf_document = fitz.open(input_pdf_path)
            output_pdf = fitz.open()
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                rect = page.rect
                new_page = output_pdf.new_page(width=target_width, height=target_height)
                scale_x = target_width / rect.width
                scale_y = target_height / rect.height
                scale_factor = min(scale_x, scale_y)
                scaled_width = rect.width * scale_factor
                scaled_height = rect.height * scale_factor
                x_offset = (target_width - scaled_width) / 2
                y_offset = (target_height - scaled_height) / 2
                matrix = fitz.Matrix(scale_factor, scale_factor).pretranslate(x_offset, y_offset)
                new_page.show_pdf_page(new_page.rect, pdf_document, page_num, matrix)
            output_pdf.save(output_pdf_path)
            output_pdf.close()
            pdf_document.close()
            print(f"Successfully resized: {input_pdf_path}")
            break  
        except Exception as e:
            print(f"Error processing {input_pdf_path}: {e}")
            print(f"Retrying in {retry_delay} seconds...")
            time.sleep(retry_delay)  

def process_directory(input_dir, output_dir):
    for root, dirs, files in os.walk(input_dir):
        for file_name in files:
            if file_name.endswith(".pdf"):
                input_pdf_path = os.path.join(root, file_name)
                relative_path = os.path.relpath(input_pdf_path, input_dir)
                output_pdf_path = os.path.join(output_dir, relative_path)
                output_subdir = os.path.dirname(output_pdf_path)
                if not os.path.exists(output_subdir):
                    os.makedirs(output_subdir)
                print(f"Resizing {input_pdf_path} to {target_width}x{target_height} points (295mm x 205mm)...")
                resize_pdf_with_retry(input_pdf_path, output_pdf_path, target_width, target_height)

process_directory(input_directory, output_directory)