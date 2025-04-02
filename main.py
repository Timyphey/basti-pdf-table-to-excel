import pandas as pd
import os
import re
import numpy as np

import fitz  # PyMuPDF
import pytesseract
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def pdf_to_excel(pdf_path, excel_path):
    print("Attempting OCR for tables...")
    doc = fitz.open(pdf_path)
    
    # Create a Pandas Excel writer
    os.makedirs(os.path.dirname(excel_path), exist_ok=True)
    excel_writer = pd.ExcelWriter(excel_path, engine='openpyxl')
    
    # Create text file for debug purposes
    txt_path = excel_path.replace(".xlsx", ".txt")
    txt_file = open(txt_path, "w", encoding="utf-8")
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=900)  # High DPI for better quality
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Preprocessing for better OCR
        img = img.convert("L")  # Grayscale
        img = img.point(lambda x: 0 if x < 128 else 255, '1')  # Binarize
        
        # Custom OCR config for German language and table structure
        # PSM 11 (sparse text) or 6 (uniform block) with OSD
        custom_config = r'--oem 3 --psm 11 -l deu'
        
        # Get text with hocr output to preserve positional data
        hocr_data = pytesseract.image_to_pdf_or_hocr(img, extension='hocr', config=custom_config)
        
        # Alternatively, use Tesseract's built-in table detection (TSV)
        tsv_output = pytesseract.image_to_data(img, config=custom_config, output_type=pytesseract.Output.DATAFRAME)
        
        # Filter out low-confidence detections and empty text
        tsv_output = tsv_output[tsv_output['conf'] > 40]
        tsv_output = tsv_output[tsv_output['text'].str.len() > 0]
        
        # Write raw OCR to text file for debugging
        text = pytesseract.image_to_string(img, config=custom_config)
        txt_file.write(f"--- PAGE {page_num+1} ---\n{text}\n\n")
        
        # Process table structure
        try:
            # First, identify text blocks that span multiple lines
            block_data = {}
            for _, row in tsv_output.iterrows():
                block_id = row['block_num']
                if block_id not in block_data:
                    block_data[block_id] = {
                        'lines': {},
                        'top': float('inf'),
                        'left': float('inf'),
                        'bottom': 0,
                        'right': 0
                    }
                
                line_id = row['line_num']
                if line_id not in block_data[block_id]['lines']:
                    block_data[block_id]['lines'][line_id] = []
                
                # Add word to the line
                block_data[block_id]['lines'][line_id].append({
                    'text': row['text'],
                    'left': row['left'],
                    'top': row['top'],
                    'width': row['width'],
                    'height': row['height']
                })
                
                # Update block boundaries
                block_data[block_id]['top'] = min(block_data[block_id]['top'], row['top'])
                block_data[block_id]['left'] = min(block_data[block_id]['left'], row['left'])
                block_data[block_id]['bottom'] = max(block_data[block_id]['bottom'], row['top'] + row['height'])
                block_data[block_id]['right'] = max(block_data[block_id]['right'], row['left'] + row['width'])
            
            # Convert blocks to a structured format for table creation
            structured_blocks = []
            for block_id, block in block_data.items():
                # Join all lines in the block with proper ordering
                block_text = []
                for line_id in sorted(block['lines'].keys()):
                    # Sort words in line by position
                    line_words = sorted(block['lines'][line_id], key=lambda x: x['left'])
                    line_text = ' '.join(word['text'] for word in line_words)
                    block_text.append(line_text)
                
                # Join all lines with newline character to preserve multi-line structure
                final_text = '\n'.join(block_text)
                
                structured_blocks.append({
                    'block_id': block_id,
                    'text': final_text,
                    'top': block['top'],
                    'left': block['left'],
                    'bottom': block['bottom'],
                    'right': block['right']
                })
            
            # Sort blocks by vertical position (top to bottom)
            structured_blocks.sort(key=lambda x: x['top'])
            
            # Group blocks into rows based on vertical position
            # Blocks whose vertical positions overlap significantly are considered part of the same row
            rows = []
            current_row = []
            last_bottom = 0
            
            for block in structured_blocks:
                # If this block starts below the bottom of the previous row (with some overlap tolerance)
                # or if it's the first block, start a new row
                if not current_row or block['top'] > last_bottom - 10:  # 10 pixels tolerance for overlap
                    if current_row:
                        rows.append(sorted(current_row, key=lambda x: x['left']))  # Sort row by horizontal position
                    current_row = [block]
                    last_bottom = block['bottom']
                else:
                    current_row.append(block)
                    last_bottom = max(last_bottom, block['bottom'])
            
            # Add the last row if it exists
            if current_row:
                rows.append(sorted(current_row, key=lambda x: x['left']))
            
            # Create a table from the rows
            table_data = []
            for row in rows:
                table_data.append([block['text'] for block in row])
            
            # Also save to CSV
            csv_folder = excel_path.replace(".xlsx", "_csv")
            os.makedirs(csv_folder, exist_ok=True)
            csv_path = os.path.join(csv_folder, f"Page_{page_num+1}.csv")

            # When processing each page, store table data for combined CSV
            if table_data:
                # Handle uneven row lengths
                max_cols = max(len(row) for row in table_data)
                padded_rows = [row + [''] * (max_cols - len(row)) for row in table_data]
                
                # Only add header row from first page with data
                if page_num == 0 or not hasattr(pdf_to_excel, 'combined_data'):
                    pdf_to_excel.combined_data = padded_rows
                else:
                    # Skip header row (first row) for subsequent pages
                    pdf_to_excel.combined_data.extend(padded_rows[1:])
                
                # Write combined data to single CSV after last page
                if page_num == len(doc) - 1:
                    combined_csv_path = excel_path.replace(".xlsx", "_combined.csv")
                    pd.DataFrame(pdf_to_excel.combined_data).to_csv(combined_csv_path, index=False, header=False)
                    print(f"All tables saved to combined CSV: {combined_csv_path}")
            
            # Create DataFrame and save to Excel
            if table_data:
                # Handle uneven row lengths
                max_cols = max(len(row) for row in table_data)
                padded_rows = [row + [''] * (max_cols - len(row)) for row in table_data]
                
                table_df = pd.DataFrame(padded_rows)
                # Save to Excel sheet
                sheet_name = f"Page_{page_num+1}"
                table_df.to_excel(excel_writer, sheet_name=sheet_name, index=False, header=False)
                print(f"Table extracted for page {page_num+1} with multi-line text blocks")
            else:
                print(f"No table data found on page {page_num+1}")
                
        except Exception as e:
            print(f"Error processing page {page_num+1}: {str(e)}")
            # Add empty sheet for failed page
            pd.DataFrame().to_excel(excel_writer, sheet_name=f"Page_{page_num+1}", index=False)
    
    # Save Excel file
    excel_writer.close()
    txt_file.close()
    
    print(f"OCR data saved to {txt_path}")
    print(f"Tables saved to Excel: {excel_path}")

def main():
    pdf_folder = "./pdfs"
    excel_folder = "./excels"
    
    # Ensure folders exist
    os.makedirs(pdf_folder, exist_ok=True)
    os.makedirs(excel_folder, exist_ok=True)
    
    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith(".pdf")]

    if not pdf_files:
        print("No PDF files found in the folder.")
        return

    print("Available PDF files:")
    for i, pdf_file in enumerate(pdf_files, start=1):
        print(f"{i}: {pdf_file}")

    choice = int(input("Enter the number of the PDF file you want to use: ")) - 1
    if choice < 0 or choice >= len(pdf_files):
        print("Invalid choice.")
        return

    pdf_path = os.path.join(pdf_folder, pdf_files[choice])
    excel_path = os.path.join(excel_folder, pdf_files[choice].replace(".pdf", ".xlsx"))
    pdf_to_excel(pdf_path, excel_path)

if __name__ == "__main__":
    main()
