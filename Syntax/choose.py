import os
import pandas as pd

def process_excel_file(file_path):

    sheet1 = pd.read_excel(file_path, sheet_name=0)  
    # Pilih kolom urutan ke-1, 2, dst
    sheet1 = sheet1.iloc[:, [0, 1, 2, 3,5,6,7,8]]

    # Baca sheet 2 (urutan kedua)
    sheet2 = pd.read_excel(file_path, sheet_name=1) 
    # Pilih kolom urutan ke-1, 2, dst
    sheet2 = sheet2.iloc[:, [0, 1, 2, 3, 4, 5,7,8,9,10]]

    # Tulis kembali hasil ke file Excel yang sama
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        sheet1.to_excel(writer, sheet_name='Sheet1', index=False)
        sheet2.to_excel(writer, sheet_name='Sheet2', index=False)

def process_directory(directory_path):
    for filename in os.listdir(directory_path):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(directory_path, filename)
            print(f"Processing file: {file_path}")
            process_excel_file(file_path)

directory_path = r'D:\Semester 7\Magang\Publikasi ST\Tanaman Perkebunan\backup 210824'
process_directory(directory_path)
